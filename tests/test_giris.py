"""Domen (Outlook) hesabı ilə giriş: kuki, yönləndirmə, parol təxmininə qarşı qoruma, ayrı sessiyalar."""
import os
import pathlib
import sys
import tempfile
import time

os.environ.setdefault("WA_DATA", tempfile.mkdtemp())
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
from fastapi.testclient import TestClient  # noqa: E402

from assistant import giris as G  # noqa: E402
from assistant.app import app  # noqa: E402

DATA = pathlib.Path(__file__).parent / "data"
ISTIFADECILER = {"aynur": ("Parol1", "Aynur Məmmədova"), "resad": ("Parol2", "Rəşad Əliyev")}


def _sahte(login, parol):
    sam, _ = G.normallasdir(login)
    if sam in ISTIFADECILER and ISTIFADECILER[sam][0] == parol:
        return {"login": sam, "ad": ISTIFADECILER[sam][1], "mail": ""}
    return None


def _qur():
    from assistant import state
    from assistant.modullar import bank as B
    G.LDAP_URL, G.DOMEN, G.YOXLAYICI = "ldaps://dc.test.local:636", "asf.gov.az", _sahte
    G._UGURSUZ.clear()
    B.HESABAT_QOVLUGU, B.QAYDA_FAYLI = state.HESABATLAR, state.DATA / "bank_qaydalari.json"  # başqa testlər dəyişə bilər


def _gir(c, login, parol):
    return c.post("/giris", data={"login": login, "parol": parol}, follow_redirects=False)


def test_girissiz_bagli():
    _qur()
    c = TestClient(app)
    r = c.get("/", follow_redirects=False)
    assert r.status_code == 303 and r.headers["location"] == "/giris"
    assert c.get("/api/veziyyet").status_code == 401
    assert c.get("/saglamliq").status_code == 200
    assert "Outlook" in c.get("/giris").text


def test_giris_ve_cixis():
    _qur()
    c = TestClient(app)
    assert _gir(c, "aynur", "sehv").status_code == 401
    r = _gir(c, "ASF\\Aynur", "Parol1")                      # domen\ad formatı da olur
    assert r.status_code == 303 and G.KUKI in r.cookies
    assert "httponly" in r.headers["set-cookie"].lower() and "samesite=lax" in r.headers["set-cookie"].lower()
    v = c.get("/api/veziyyet").json()
    assert v["istifadeci"] == {"ad": "Aynur Məmmədova", "login": "aynur"}
    c.get("/cixis", follow_redirects=False)
    c.cookies.clear()
    assert c.get("/api/veziyyet").status_code == 401


def test_saxta_ve_kohne_kuki():
    _qur()
    k = G.kuki_yarat({"login": "aynur", "ad": "A"})
    assert G.kuki_oxu(k)["l"] == "aynur"
    yuk, imza = k.split(".")
    assert G.kuki_oxu(G._b64(b'{"l":"resad","exp":9999999999}') + "." + imza) is None   # imza uyğun gəlmir
    saat = G.SESSIYA_SAAT
    G.SESSIYA_SAAT = -1
    assert G.kuki_oxu(G.kuki_yarat({"login": "aynur"})) is None                         # vaxtı keçib
    G.SESSIYA_SAAT = saat


def test_parol_texmini_bloklanir():
    _qur()
    c = TestClient(app)
    for _ in range(G.MAKS_CEHD):
        assert _gir(c, "resad", "sehv").status_code == 401
    assert _gir(c, "resad", "Parol2").status_code == 429     # düzgün parol da 10 dəqiqə qəbul olunmur


def test_istifadeci_adi_yoxlamasi():
    G.DOMEN = "asf.gov.az"
    assert G.normallasdir("Aynur.M") == ("aynur.m", "aynur.m@asf.gov.az")
    assert G.normallasdir("aynur@asf.gov.az") == ("aynur", "aynur@asf.gov.az")
    assert G.normallasdir("*)(uid=*") == (None, None)                        # LDAP injection
    assert G.ad_yoxla("aynur", "") is None                                   # boş parol (anonim bind) qəbul olunmur


def test_domen_serveri_elcatmaz():
    G.LDAP_URL = "ldaps://127.0.0.1:1"
    try:
        G.ad_yoxla("aynur", "x")
        raise AssertionError("xəta gözlənilirdi")
    except G.GirisXetasi as x:
        assert "qoşulmaq" in str(x)


def test_her_istifadecinin_oz_sessiyasi():
    from assistant.modullar import bank as B
    _qur()
    a, b = TestClient(app), TestClient(app)
    _gir(a, "aynur", "Parol1")
    _gir(b, "resad", "Parol2")
    with open(DATA / "bank_kapital_test.xlsx", "rb") as f:
        j = a.post("/api/yukle", files=[("fayl", ("cixaris.xlsx", f))]).json()
    for _ in range(100):
        st = a.get(f"/api/yukle/{j['is']}").json()
        if st["status"] != "isleyir":
            break
        time.sleep(0.1)
    assert st["status"] == "hazir" and st["netice"]["bank"]["daxilolma"] == 9
    assert b.get(f"/api/yukle/{j['is']}").status_code == 404           # başqasının işini görmür
    assert a.get("/api/veziyyet").json()["hazir"] is True
    assert b.get("/api/veziyyet").json()["hazir"] is False             # Rəşad Aynurun faylını görmür
    B._qaydalar_yaz({**B._qaydalar(), "madde": {k["kod"]: k["teklif"] for k in B.KATEQORIYALAR}})   # mühasib maddələri deyib
    r = a.post("/api/sohbet", json={"mesaj": "1C üçün Excel hazırla"}).json()
    assert r["hadiseler"][0]["ui"]["tip"] == "sual"                       # maddə hər doldurmada soruşulur
    r = a.post("/api/sohbet", json={"mesaj": "Sığorta haqları maddəsi \"Sığorta haqları\" olsun"}).json()
    url = r["hadiseler"][0]["ui"]["fayllar"][0]["url"]
    assert a.get(url).status_code == 200
    assert b.get(url).status_code == 404                                 # başqasının hesabatını yükləyə bilmir


# ---------------- IMAP (XDMX mail.gov.az) ----------------
class _SahteIMAP:
    hesablar = {"mirsaghli.shiraliyev@asf.gov.az": "DogruParol"}

    def __init__(self, host, port, ssl_context=None, timeout=None):
        assert host == "mail.gov.az" and port == 993 and ssl_context is not None   # yalnız SSL

    def login(self, ad, parol):
        import imaplib
        if self.hesablar.get(ad) != parol:
            raise imaplib.IMAP4.error("AUTHENTICATIONFAILED")

    def logout(self):
        pass


def _imap(monkeypatch, icazeli=""):
    import imaplib
    monkeypatch.setattr(imaplib, "IMAP4_SSL", _SahteIMAP)
    monkeypatch.setattr(G, "IMAP_HOST", "mail.gov.az")
    monkeypatch.setattr(G, "IMAP_PORT", 993)
    monkeypatch.setattr(G, "IMAP_DOMENLER", ["asf.gov.az"])
    monkeypatch.setenv("WA_ICAZELI", icazeli)
    monkeypatch.setenv("WA_DATA", tempfile.mkdtemp())


def test_imap_giris(monkeypatch):
    _imap(monkeypatch)
    assert G.epoct("Mirsaghli.Shiraliyev") == "mirsaghli.shiraliyev@asf.gov.az"    # domen avtomatik əlavə olunur
    u = G.imap_yoxla("mirsaghli.shiraliyev", "DogruParol")
    assert u == {"login": "mirsaghli.shiraliyev@asf.gov.az", "ad": "Mirsaghli Shiraliyev", "mail": "mirsaghli.shiraliyev@asf.gov.az"}
    assert G.imap_yoxla("mirsaghli.shiraliyev@asf.gov.az", "sehv") is None
    assert G.imap_yoxla("mirsaghli.shiraliyev@asf.gov.az", "") is None


def test_imap_basqa_qurum_ve_icazeli_siyahi(monkeypatch):
    _imap(monkeypatch)
    try:
        G.imap_yoxla("kimse@mia.gov.az", "x")                  # mail.gov.az-da başqa qurumun işçisi girə bilməz
        raise AssertionError
    except G.GirisXetasi as x:
        assert "@asf.gov.az" in str(x)
    _imap(monkeypatch, icazeli="muhasib@asf.gov.az")
    try:
        G.imap_yoxla("mirsaghli.shiraliyev@asf.gov.az", "DogruParol")   # siyahıda yoxdur
        raise AssertionError
    except G.GirisXetasi as x:
        assert "icazəniz yoxdur" in str(x)


def test_imap_server_elcatmaz(monkeypatch):
    _imap(monkeypatch)
    import imaplib
    monkeypatch.setattr(imaplib, "IMAP4_SSL", _Elcatmaz)
    try:
        G.imap_yoxla("mirsaghli.shiraliyev", "x")
        raise AssertionError
    except G.GirisXetasi as x:
        assert "mail.gov.az" in str(x)


class _Elcatmaz:
    def __init__(self, *a, **k):
        raise TimeoutError("timed out")


class _Sock:
    def __init__(self, der):
        self.der = der

    def getpeercert(self, binary_form=False):
        return self.der


def test_imap_sertifikat_sancmasi(monkeypatch):
    import hashlib
    import ssl
    _imap(monkeypatch)
    der = b"mail.gov.az-in-sertifikati"
    yaxalanan = {}

    class Pinli(_SahteIMAP):
        def __init__(self, host, port, ssl_context=None, timeout=None):
            yaxalanan["ctx"] = ssl_context
            self.sock = _Sock(der)

        def shutdown(self):
            pass
    import imaplib
    monkeypatch.setattr(imaplib, "IMAP4_SSL", Pinli)
    monkeypatch.setattr(G, "IMAP_PIN", True)
    monkeypatch.setattr(G, "IMAP_PINLER", {"00" * 32, hashlib.sha256(der).hexdigest()})   # bir neçə node
    assert G.imap_yoxla("mirsaghli.shiraliyev", "DogruParol")["login"] == "mirsaghli.shiraliyev@asf.gov.az"
    assert yaxalanan["ctx"].verify_mode == ssl.CERT_NONE                 # vaxtı keçmiş sertifikat qəbul olunur...
    monkeypatch.setattr(G, "IMAP_PINLER", {"00" * 32})                    # ...amma yalnız məhz o sertifikat
    try:
        G.imap_yoxla("mirsaghli.shiraliyev", "DogruParol")
        raise AssertionError
    except G.GirisXetasi as x:
        assert "Parol göndərilmədi" in str(x)


def test_imap_sertifikat_adi(monkeypatch):
    import ssl
    _imap(monkeypatch)
    yaxalanan = {}
    sert = {"subject": ((("commonName", "*.mail.gov.az"),),), "subjectAltName": (("DNS", "*.mail.gov.az"),)}

    class Adli(_SahteIMAP):
        def __init__(self, host, port, ssl_context=None, timeout=None):
            yaxalanan["ctx"] = ssl_context
            self.sock = type("S", (), {"getpeercert": lambda self, binary_form=False: sert})()

        def shutdown(self):
            pass
    import imaplib
    monkeypatch.setattr(imaplib, "IMAP4_SSL", Adli)
    monkeypatch.setattr(G, "IMAP_PIN", False)
    monkeypatch.setattr(G, "IMAP_SERT_ADI", "*.mail.gov.az")
    assert G.imap_yoxla("mirsaghli.shiraliyev", "DogruParol")
    assert yaxalanan["ctx"].verify_mode == ssl.CERT_REQUIRED          # zəncir və vaxt yoxlanmaqda davam edir
    monkeypatch.setattr(G, "IMAP_SERT_ADI", "mail.example.az")
    try:
        G.imap_yoxla("mirsaghli.shiraliyev", "DogruParol")
        raise AssertionError
    except G.GirisXetasi as x:
        assert "Parol göndərilmədi" in str(x)


def test_sohbetler_istifadeciye_gore():
    _qur()
    a, b = TestClient(app), TestClient(app)
    _gir(a, "aynur", "Parol1")
    _gir(b, "resad", "Parol2")
    m = [{"rol": "user", "metn": "Xülasə"}, {"rol": "ai", "cavab": "58 daxilolma", "hadiseler": []}]
    assert a.put("/api/sohbetler/abc1", json={"ad": "Oktyabr çıxarışı", "mesajlar": m}).status_code == 200
    assert [s["ad"] for s in a.get("/api/sohbetler").json()] == ["Oktyabr çıxarışı"]
    assert a.get("/api/sohbetler/abc1").json()["mesajlar"] == m
    assert b.get("/api/sohbetler").json() == []                     # Rəşad Aynurun söhbətlərini görmür
    assert b.get("/api/sohbetler/abc1").status_code == 404
    assert a.put("/api/sohbetler/../x", json={"ad": "", "mesajlar": []}).status_code in (400, 404)
    c = TestClient(app)                                              # başqa kompüterdən eyni istifadəçi
    _gir(c, "aynur@asf.gov.az", "Parol1")
    assert [s["id"] for s in c.get("/api/sohbetler").json()] == ["abc1"]
    a.delete("/api/sohbetler/abc1")
    assert a.get("/api/sohbetler").json() == []


def test_adla_salamlasir():
    _qur()
    a = TestClient(app)
    _gir(a, "aynur", "Parol1")
    assert a.post("/api/sohbet", json={"mesaj": "Salam"}).json()["cavab"].startswith("Salam, Aynur")
    assert "**Məmmədova**" in a.post("/api/sohbet", json={"mesaj": "Mən kiməm?"}).json()["cavab"]


def test_ad_soyad_e_poctdan():
    import types
    from assistant import sohbet as SB
    s = types.SimpleNamespace(istifadeci_adi="Mirsaghli Shiraliyev", istifadeci_login="mirsaghli.shiraliyev@asf.gov.az",
                              fayl_var=False, excel=None)
    for sual in ["Mənim adım nədir", "Soyadım nədir?", "Məni tanıyırsan?", "Mən kiməm"]:
        c = SB.cavab(s, sual)
        assert c and "**Mirsaghli**" in c and "**Shiraliyev**" in c and "mirsaghli.shiraliyev@asf.gov.az" in c, sual
    assert SB.cavab(s, "Kür Avto-nun adı ilə borcu nə qədərdir?") is None   # iş sualı şablona düşmür


def _yukle_gozle(c, fayllar, rejim=""):
    files = [("fayl", (ad, open(yol, "rb"))) for yol, ad in fayllar]
    j = c.post("/api/yukle", files=files, data={"rejim": rejim}).json()
    for _ in range(100):
        st = c.get(f"/api/yukle/{j['is']}").json()
        if st["status"] != "isleyir":
            return st
        time.sleep(0.1)


def test_fayllarim_yadda_qalir_ve_secilir():
    _qur()
    a, b = TestClient(app), TestClient(app)
    _gir(a, "aynur", "Parol1")
    _gir(b, "resad", "Parol2")
    st = _yukle_gozle(a, [(DATA / "bank_kapital_test.xlsx", "Oktyabr çıxarışı.xlsx")])
    assert st["status"] == "hazir"
    fl = a.get("/api/fayllar").json()
    assert [f["ad"] for f in fl] == ["Oktyabr çıxarışı.xlsx"]
    assert b.get("/api/fayllar").json() == []                          # Rəşad Aynurun fayllarını görmür
    assert b.post(f"/api/fayllar/{fl[0]['id']}/ac").status_code == 404
    # başqa faylla işləyib sonra siyahıdan geri açmaq — yenidən yükləmədən
    _yukle_gozle(a, [(DATA / "ixtiyari_excel.xlsx", "satislar.xlsx")])
    assert a.get("/api/veziyyet").json()["bank"] is None
    j = a.post(f"/api/fayllar/{fl[0]['id']}/ac").json()
    for _ in range(100):
        st = a.get(f"/api/yukle/{j['is']}").json()
        if st["status"] != "isleyir":
            break
        time.sleep(0.1)
    assert st["netice"]["bank"]["daxilolma"] == 9 and len(a.get("/api/fayllar").json()) == 2


def test_tekrar_yukleme_ve_ustune_yazma():
    _qur()
    a = TestClient(app)
    _gir(a, "aynur", "Parol1")
    yol = DATA / "bank_kapital_test.xlsx"
    _yukle_gozle(a, [(yol, "cixaris.xlsx")])
    k = a.post("/api/fayllar/yoxla", json={"fayllar": [{"ad": "Cixaris.xlsx", "olcu": yol.stat().st_size}]}).json()
    assert len(k) == 1 and k[0]["eyni_olcu"] is True                     # interfeys «saxlanmışı açım?» soruşur
    assert a.post("/api/fayllar/yoxla", json={"fayllar": [{"ad": "basqa.xlsx", "olcu": 1}]}).json() == []
    _yukle_gozle(a, [(yol, "cixaris.xlsx")])                               # eyni məzmun: təkrar saxlanmır
    assert len([f for f in a.get("/api/fayllar").json() if f["ad"] == "cixaris.xlsx"]) == 1
    _yukle_gozle(a, [(DATA / "bank_cixarisi_test.xlsx", "cixaris.xlsx")], rejim="uzerine")   # fərqli məzmun, üstünə
    ss = [f for f in a.get("/api/fayllar").json() if f["ad"] == "cixaris.xlsx"]
    assert len(ss) == 1 and ss[0]["olcu"] == (DATA / "bank_cixarisi_test.xlsx").stat().st_size
    _yukle_gozle(a, [(yol, "cixaris.xlsx")], rejim="yeni")                  # ayrıca saxla
    assert len([f for f in a.get("/api/fayllar").json() if f["ad"] == "cixaris.xlsx"]) == 2
    fid = a.get("/api/fayllar").json()[0]["id"]
    assert a.delete(f"/api/fayllar/{fid}").status_code == 200


# ---------------- Gmail (Google ilə daxil ol) ----------------
def _google(monkeypatch, siyahi="test.istifadeci@gmail.com"):
    _qur()
    monkeypatch.setattr(G, "GOOGLE_ID", "test-id.apps.googleusercontent.com")
    monkeypatch.setattr(G, "GOOGLE_SECRET", "sirr")
    monkeypatch.setattr(G, "GOOGLE_REDIRECT", "")
    monkeypatch.setenv("WA_ICAZELI", siyahi)


def _iddia(**x):
    return {"aud": G.GOOGLE_ID, "iss": "https://accounts.google.com", "exp": time.time() + 600,
            "email": "Test.Istifadeci@gmail.com", "email_verified": True, "name": "Test İstifadəçi", **x}


def test_giris_secimi_outlook_ve_gmail(monkeypatch):
    _google(monkeypatch)
    t = TestClient(app).get("/giris").text
    assert 'data-u="outlook"' in t and 'data-u="gmail"' in t                  # iki seçim
    assert 'action="/giris"' in t and 'href="/giris/google"' in t
    monkeypatch.setattr(G, "LDAP_URL", "")                                      # yalnız Gmail qurulub: seçim yoxdur
    monkeypatch.setattr(G, "IMAP_HOST", "")
    t = TestClient(app).get("/giris").text
    assert 'data-u="outlook"' not in t and 'action="/giris"' not in t and "Google ilə daxil ol" in t


def test_gmail_giris_axini(monkeypatch):
    from urllib.parse import parse_qs, urlparse
    _google(monkeypatch)
    alinan = {}

    def sahte(kod, verifier, redirect):
        alinan.update(kod=kod, verifier=verifier, redirect=redirect)
        return G.google_iddia_yoxla(_iddia())
    monkeypatch.setattr(G, "GOOGLE_YOXLAYICI", sahte)
    c = TestClient(app)
    r = c.get("/giris/google", follow_redirects=False)
    u = urlparse(r.headers["location"])
    q = {k: v[0] for k, v in parse_qs(u.query).items()}
    assert u.netloc == "accounts.google.com" and q["client_id"] == G.GOOGLE_ID
    assert q["code_challenge_method"] == "S256" and q["redirect_uri"] == "http://testserver/giris/google/geri"
    assert c.get("/giris/google/geri", params={"code": "k", "state": "basqa"}, follow_redirects=False).status_code == 400  # saxta state
    r = c.get("/giris/google", follow_redirects=False)
    state = parse_qs(urlparse(r.headers["location"]).query)["state"][0]
    r = c.get("/giris/google/geri", params={"code": "kod1", "state": state}, follow_redirects=False)
    assert r.status_code == 303 and G.KUKI in r.cookies
    assert alinan["kod"] == "kod1" and alinan["redirect"] == "http://testserver/giris/google/geri"
    v = c.get("/api/veziyyet").json()
    assert v["istifadeci"]["login"] == "test.istifadeci@gmail.com"


def test_gmail_icaze_yoxlamasi(monkeypatch):
    import pytest
    _google(monkeypatch)
    assert G.google_iddia_yoxla(_iddia())["ad"] == "Test İstifadəçi"
    with pytest.raises(G.GirisXetasi, match="icazəsi yoxdur"):
        G.google_iddia_yoxla(_iddia(email="yad@gmail.com"))                    # siyahıda deyil
    with pytest.raises(G.GirisXetasi, match="təsdiqlənməyib"):
        G.google_iddia_yoxla(_iddia(email_verified=False))
    with pytest.raises(G.GirisXetasi, match="yoxlanmadı"):
        G.google_iddia_yoxla(_iddia(aud="basqa-musteri"))                       # başqa tətbiqin tokeni
    monkeypatch.setenv("WA_ICAZELI", "")
    with pytest.raises(G.GirisXetasi, match="siyahısı qurulmayıb"):            # siyahısız Gmail girişi bağlıdır
        G.google_iddia_yoxla(_iddia())


# ---------------- peşələr: girişdən sonra soruşulur və yadda qalır ----------------
def test_pesheler_sorusulur_ve_yadda_qalir():
    _qur()
    a, b = TestClient(app), TestClient(app)
    _gir(a, "aynur", "Parol1")
    _gir(b, "resad", "Parol2")
    v = a.get("/api/veziyyet").json()
    assert v["profil"] is None and any(p["kod"] == "muhasib" for p in v["peseler"])   # seçilməyib: interfeys soruşur
    assert a.put("/api/profil", json={"peseler": []}).status_code == 400               # boş seçim qəbul olunmur
    assert a.put("/api/profil", json={"peseler": ["satinalma"]}).status_code == 400     # siyahıda olmayan peşə yazıla bilməz
    r = a.put("/api/profil", json={"peseler": ["muhasib", "yoxdur", "muhasib"], "diger": "özü yazdığı"})
    assert r.json()["profil"] == {"peseler": ["muhasib"]}
    h = a.get("/api/veziyyet").json()["heller"]                                         # həllər peşəyə görə
    assert h[0]["s"] == "1C üçün Excel hazırla" and all(x["pese"] == "Mühasib" for x in h)
    assert TestClient(app, cookies=a.cookies).get("/api/veziyyet").json()["profil"]["peseler"] == ["muhasib"]   # yeni brauzerdə də
    assert b.get("/api/veziyyet").json()["profil"] is None                                # başqa istifadəçinin seçimi ayrıdır
    from assistant import profil as PR
    assert PR.modullar({"peseler": ["muhasib", "huquq"]}) == ["bank", "muhasibat", "excel", "muqavile"]
    hl = PR.heller({"peseler": ["muhasib", "rehber"]})
    assert len({x["s"] for x in hl}) == len(hl) and {x["pese"] for x in hl} == {"Mühasib", "Rəhbər"}   # təkrarsız


def test_hellerin_hamisi_basa_dusulur():
    """Peşə həllərinin hər sorğusu işləyən bir alətə gedir (boş vəd yoxdur): ya modulun öz sorğusudur,
    ya da klinika qəbul jurnalı nümunəsində (tests/data/klinika_test.xlsx) həqiqətən nəticə verir."""
    from assistant import profil as PR
    from assistant.modullar import AKTIV
    from assistant.tools import norm
    isleyen = {norm(t["s"]) for m in AKTIV for t in m.MODUL["teklifler"]}
    _qur()
    c = TestClient(app)
    _gir(c, "aynur", "Parol1")
    with open(DATA / "klinika_test.xlsx", "rb") as f:
        j = c.post("/api/yukle", files=[("fayl", ("klinika.xlsx", f))]).json()
    for _ in range(100):
        if c.get(f"/api/yukle/{j['is']}").json()["status"] != "isleyir":
            break
        time.sleep(0.1)
    for p in PR.PESELER:
        assert p["heller"], p["kod"]
        for h in p["heller"]:
            if norm(h["s"]) in isleyen:
                continue
            r = c.post("/api/sohbet", json={"mesaj": h["s"]}).json()
            alet = r["hadiseler"][0]["alet"] if r.get("hadiseler") else None
            assert alet in {"excel_xulase", "sutunlar", "qrupla", "en_boyuk", "problemler"}, (p["kod"], h["s"], r)
            assert "tapılmadı" not in r["cavab"] and "nan" not in r["cavab"], (h["s"], r["cavab"][:200])


def test_tibb_hellerinin_neticesi():
    from assistant.modullar import excel as EX

    class K:
        excel = EX.ac(DATA / "klinika_test.xlsx", tempfile.mkdtemp())
    m = EX.qrupla(K(), **EX.qayda(K(), "Məbləğ cəmi həkim üzrə")[1])["metn"]
    assert "Dr. Həsənov" in m and "(boş)" in m                                  # boş həkim xanası «nan» deyil
    m = EX.qrupla(K(), **EX.qayda(K(), "Məbləğ cəmi xidmət üzrə")[1])["metn"]
    assert m.index("Implant") < m.index("Plomb")                                # ən çox gəlir gətirən öndədir
    assert "Təkrar sətirlər" in EX.problemler(K())["metn"]

def test_hekim_teleb_yazir_rehber_tesdiq_edir():
    """Həkim fayl yükləmədən stoka tələb yazır; rəhbər (peşəsi «Diş klinikasının rəhbəri») görür və təsdiq edir."""
    _qur()
    hekim, rehber = TestClient(app), TestClient(app)
    _gir(hekim, "aynur", "Parol1")
    _gir(rehber, "resad", "Parol2")
    hekim.put("/api/profil", json={"peseler": ["dis_hekimi"]})
    rehber.put("/api/profil", json={"peseler": ["klinika_rehber"]})
    r = hekim.post("/api/sohbet", json={"mesaj": "Stoka tələb: Kompozit plomb A2 — 5 ədəd, təcili"}).json()
    assert "qeydə alındı" in r["cavab"] and "fayl lazımdır" not in r["cavab"]
    no = r["cavab"].split("Tələb ")[1].split(" ")[0]
    r = hekim.post("/api/sohbet", json={"mesaj": f"{no} təsdiq"}).json()
    assert "yalnız klinika rəhbəri" in r["cavab"]
    r = rehber.post("/api/sohbet", json={"mesaj": "Stok tələblərini göstər"}).json()
    setirler = r["hadiseler"][0]["ui"]["bloklar"][0]["setirler"]
    assert any(x[0] == no and x[5] == "Aynur Məmmədova" and x[3] == "təcili" for x in setirler)
    r = rehber.post("/api/sohbet", json={"mesaj": f"{no} təsdiq"}).json()
    assert r["cavab"].startswith(f"Oldu: {no} — təsdiqləndi")
    r = hekim.post("/api/sohbet", json={"mesaj": "Stok tələblərini göstər"}).json()
    assert any(x[0] == no and x[4] == "təsdiqləndi" for x in r["hadiseler"][0]["ui"]["setirler"])   # həkim nəticəni görür


def test_peseye_gore_telimat():
    """Peşə seçiləndə təlimat qaytarılır; «Təlimat» yazanda yenidən açılır; nümunə sorğuların hamısı işləyir."""
    from assistant import profil as PR
    from assistant.modullar import AKTIV
    from assistant.tools import norm
    assert set(PR.TELIMATLAR) == {p["kod"] for p in PR.PESELER}                    # hər peşənin təlimatı var
    isleyen = {norm(t["s"]) for m in AKTIV for t in m.MODUL["teklifler"]} | {norm(h["s"]) for p in PR.PESELER for h in p["heller"]}
    for kod, t in PR.TELIMATLAR.items():
        assert t["giris"] and t["fayl"] and t["addimlar"], kod
        for a in t["addimlar"]:
            assert a["b"] and a["m"], kod
            assert "s" not in a or norm(a["s"]) in isleyen, (kod, a["s"])
    _qur()
    c = TestClient(app)
    _gir(c, "aynur", "Parol1")
    r = c.put("/api/profil", json={"peseler": ["dis_hekimi", "klinika_rehber"]}).json()
    assert [p["ad"] for p in r["telimat"]["peseler"]] == ["Diş həkimi", "Həkim", "Diş klinikasının rəhbəri", "Rəhbər"]   # üst peşələr də
    r = c.post("/api/sohbet", json={"mesaj": "Təlimat"}).json()                       # fayl yüklənməyib, yenə də açılır
    ui = r["hadiseler"][0]["ui"]
    assert ui["tip"] == "telimat" and ui["peseler"][0]["addimlar"][0]["s"] == "Stoka tələb yaz"


def test_dis_hekimi_hekimin_altindadir():
    """Diş həkimi «Həkim»in alt peşəsidir: seçiləndə həkimin həlləri, modulları və təlimatı da əlavə olunur."""
    from assistant import profil as PR
    assert next(p for p in PR.PESELER if p["kod"] == "dis_hekimi")["ust"] == "hekim"
    prf = {"peseler": ["dis_hekimi"]}
    assert PR.effektiv(prf) == ["dis_hekimi", "hekim"]
    hl = PR.heller(prf)
    assert {"Diş həkimi", "Həkim"} == {h["pese"] for h in hl}
    assert "Sütunları göstər" in [h["s"] for h in hl]                                  # həkimin həlli
    assert hl[0]["pese"] == "Diş həkimi"                                              # öz həlləri əvvəl
    assert len({h["s"] for h in hl}) == len(hl)                                       # təkrarsız
    assert [p["ad"] for p in PR.telimat(prf)["peseler"]] == ["Diş həkimi", "Həkim"]
    assert PR.effektiv({"peseler": ["hekim"]}) == ["hekim"]                           # tərsinə deyil
    assert PR.effektiv({"peseler": ["hekim", "dis_hekimi"]}) == ["hekim", "dis_hekimi"]
    _qur()
    c = TestClient(app)
    _gir(c, "aynur", "Parol1")
    assert c.put("/api/profil", json={"peseler": ["dis_hekimi"]}).json()["profil"] == {"peseler": ["dis_hekimi"]}
    v = c.get("/api/veziyyet").json()
    assert v["pese_effektiv"] == ["dis_hekimi", "hekim"] and any(h["pese"] == "Həkim" for h in v["heller"])


def test_klinika_rehberi_rehberin_altindadir():
    """Diş klinikasının rəhbəri «Rəhbər»in alt peşəsidir: rəhbərin həlləri və təlimatı da gəlir, rəhbər hüquqları qalır."""
    from assistant import profil as PR
    from assistant.modullar import teleb as TELEB
    prf = {"peseler": ["klinika_rehber"]}
    assert PR.effektiv(prf) == ["klinika_rehber", "rehber"]
    hl = PR.heller(prf)
    assert hl[0]["pese"] == "Diş klinikasının rəhbəri" and {"Vəziyyət necədir?", "Debitor borcları göstər"} <= {h["s"] for h in hl}
    assert [p["ad"] for p in PR.telimat(prf)["peseler"]] == ["Diş klinikasının rəhbəri", "Rəhbər"]
    assert PR.effektiv({"peseler": ["rehber"]}) == ["rehber"]                         # adi rəhbər klinika rəhbəri olmur
    _qur()
    c = TestClient(app)
    _gir(c, "resad", "Parol2")
    c.put("/api/profil", json={"peseler": ["rehber"]})

    class S:
        istifadeci_login = "resad"
    assert not TELEB.rehberdir(S())                                                  # tələbləri yalnız klinika rəhbəri təsdiq edir
    c.put("/api/profil", json={"peseler": ["klinika_rehber"]})
    assert TELEB.rehberdir(S())


def test_modullar_peseye_aiddir():
    """Yan paneldə yalnız istifadəçinin peşəsinə aid işləyən modullar görünür."""
    _qur()
    c = TestClient(app)
    _gir(c, "aynur", "Parol1")
    kodlar = lambda: [m["kod"] for m in c.get("/api/veziyyet").json()["modullar"]]  # noqa: E731
    from assistant import profil as PR
    from assistant import tools as T
    hamisi = [m["kod"] for m in PR.modul_siyahisi(T.modullar(), None)]
    assert {"bank", "qebul", "stok"} <= set(hamisi) and "muqavile" not in hamisi      # peşə yoxdur: bütün işləyənlər
    c.put("/api/profil", json={"peseler": ["muhasib"]})
    assert kodlar() == ["bank", "muhasibat", "excel"]
    c.put("/api/profil", json={"peseler": ["dis_hekimi"]})
    assert kodlar() == ["qebul", "teleb", "excel"]                                   # həkimdən excel gəlir, bank yoxdur
    c.put("/api/profil", json={"peseler": ["huquq"]})
    assert kodlar() == ["excel"]                                                     # «Müqavilələr» hələ işləmir
    c.put("/api/profil", json={"peseler": ["klinika_rehber"]})
    assert kodlar() == ["qebul", "stok", "teleb", "excel", "bank", "muhasibat"]      # rəhbərdən bank, mühasibat
