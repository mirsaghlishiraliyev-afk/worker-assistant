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
    r = a.post("/api/sohbet", json={"mesaj": "1C üçün Excel hazırla"}).json()
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
    assert "Aynur Məmmədova" in a.post("/api/sohbet", json={"mesaj": "Mən kiməm?"}).json()["cavab"]
