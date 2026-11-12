"""Mobil giriş və qeydiyyat: telefon nömrəsi + SMS/WhatsApp kodu, ad soyad, peşə məcburidir; nömrə ilə tanınma."""
import os
import pathlib
import sys
import tempfile
import time

import pytest

os.environ.setdefault("WA_DATA", tempfile.mkdtemp())
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
from fastapi.testclient import TestClient  # noqa: E402

from assistant import giris as G  # noqa: E402
from assistant import hesab as H  # noqa: E402
from assistant.app import app  # noqa: E402

pytestmark = pytest.mark.qeydiyyat
GELEN = {}


@pytest.fixture(autouse=True)
def muhit(monkeypatch):
    monkeypatch.setattr(H, "FAYL", pathlib.Path(tempfile.mkdtemp()) / "hesablar.json")
    monkeypatch.setattr(H, "_KODLAR", {})
    monkeypatch.setattr(H, "KOD_ARA", 0)                                            # testlərdə dərhal yenidən göndərmək olsun
    monkeypatch.setattr(G, "LDAP_URL", "ldaps://dc.test.local:636")
    monkeypatch.setattr(G, "DOMEN", "asf.gov.az")
    monkeypatch.setattr(G, "YOXLAYICI", lambda l, p: (
        {"login": "aynur", "ad": "Aynur Məmmədova", "mail": ""} if p == "Parol1" else
        {"login": l.lower(), "ad": "Leyla Həsənova", "mail": l.lower()} if p == "PoctParol" else None))
    monkeypatch.setattr(G, "YERLI", False)
    from assistant.modullar import qebul as QB
    monkeypatch.setattr(QB, "TOKEN", "t")                                              # WhatsApp kanalı «qurulub»
    monkeypatch.setattr(QB, "PHONE_ID", "1")
    monkeypatch.setenv("WA_SMS_URL", "https://sms.test/send?to={tel}&text={metn}")
    monkeypatch.setitem(H.GONDERICI, "sms", lambda tel, kod: GELEN.__setitem__(tel, kod))
    monkeypatch.setitem(H.GONDERICI, "whatsapp", lambda tel, kod: GELEN.__setitem__(tel, kod))
    G._UGURSUZ.clear()
    GELEN.clear()


def _kod(c, tel, kanal="sms"):
    j = c.post("/giris/kod", data={"tel": tel, "kanal": kanal}).json()
    assert j["ok"], j
    return GELEN["994501112233"]


def test_telefonla_ilk_giris_qeydiyyat_ve_tekrar_giris():
    c = TestClient(app)
    assert 'data-u="mobil"' in c.get("/giris").text                                  # Mobil bölməsi
    kod = _kod(c, "050 111 22 33")
    r = c.post("/giris/mobil", data={"tel": "050 111 22 33", "kod": kod}, follow_redirects=False)
    assert r.headers["location"] == "/qeydiyyat"                                      # ilk dəfə: qeydiyyat
    s = c.get("/qeydiyyat").text
    assert "tapa bilmədim" in s and "Nömrə təsdiqlənib" in s and 'name="peseler"' in s
    assert "Peşə məcburidir" in c.post("/qeydiyyat", data={"ad": "Leyla Həsənova", "mail": "leyla@asf.gov.az"}).text
    assert "Ad soyadı yazın" in c.post("/qeydiyyat", data={"ad": "", "peseler": ["muhasib"]}).text
    assert "Poçt ünvanı məcburidir" in c.post("/qeydiyyat", data={"ad": "Leyla Həsənova", "peseler": ["muhasib"]}).text
    r = c.post("/qeydiyyat", data={"ad": "Leyla Həsənova", "mail": "Leyla@asf.gov.az", "peseler": ["muhasib"]}, follow_redirects=False)
    assert r.status_code == 303 and G.KUKI in r.cookies
    v = c.get("/api/veziyyet").json()
    assert v["istifadeci"] == {"ad": "Leyla Həsənova", "login": "994501112233"} and v["profil"] == {"peseler": ["muhasib"]}
    # ikinci dəfə: sistem nömrəni tanıyır, qeydiyyat yoxdur
    d = TestClient(app)
    kod = _kod(d, "+994 50 111 22 33", "whatsapp")
    r = d.post("/giris/mobil", data={"tel": "+994501112233", "kod": kod}, follow_redirects=False)
    assert r.headers["location"] == "/" and d.get("/api/veziyyet").json()["istifadeci"]["ad"] == "Leyla Həsənova"


def test_poctla_giren_telefon_ve_pese_olmadan_davam_edemir_parolu_unudanda_telefonla_girir():
    c = TestClient(app)
    c.post("/giris", data={"login": "aynur", "parol": "Parol1"})
    assert c.get("/api/veziyyet").status_code == 403                                    # qeydiyyat tamamlanmayıb
    assert c.get("/", follow_redirects=False).headers["location"] == "/qeydiyyat"
    s = c.get("/qeydiyyat").text
    assert 'value="Aynur Məmmədova"' in s and "Təsdiq kodu göndər" in s                # ad poçtdan tapılır, dəyişmək olar
    assert "Nömrəni təsdiqləyin" in c.post("/qeydiyyat", data={"ad": "Aynur Məmmədova", "tel": "0501112233", "mail": "aynur@asf.gov.az",
                                                               "peseler": ["muhasib"]}).text
    kod = _kod(c, "0501112233")
    r = c.post("/qeydiyyat", data={"ad": "Aynur Məmmədova", "tel": "0501112233", "kod": kod, "mail": "aynur@asf.gov.az",
                                   "peseler": ["dis_hekimi"]}, follow_redirects=False)
    assert r.status_code == 303
    assert H.al("aynur")["mail"] == "aynur@asf.gov.az"                                # kod yoxlananda poçt itmir
    assert c.get("/api/veziyyet").status_code == 200
    # poçtun parolunu unudub: telefonla girir — sistem onu «aynur» hesabına salır
    t = TestClient(app)
    kod = _kod(t, "050 111 22 33")
    t.post("/giris/mobil", data={"tel": "0501112233", "kod": kod})
    v = t.get("/api/veziyyet").json()
    assert v["istifadeci"]["login"] == "aynur" and v["profil"] == {"peseler": ["dis_hekimi"]}


def test_kod_tehlukesizliyi(monkeypatch):
    monkeypatch.setattr(H, "KOD_ARA", 60)
    c = TestClient(app)
    kod = _kod(c, "0501112233")
    assert not c.post("/giris/kod", data={"tel": "0501112233"}).json()["ok"]            # 60 saniyə gözləmək lazımdır
    for _ in range(5):
        assert c.post("/giris/mobil", data={"tel": "0501112233", "kod": "000000" if kod != "000000" else "111111"}).status_code == 401
    assert c.post("/giris/mobil", data={"tel": "0501112233", "kod": kod}).status_code == 429   # çox səhv cəhd
    H._KODLAR["994501112233"].update(hash=H._hash("994501112233", "123456"), bitir=time.time() - 1, cehd=0)
    assert H.kod_yoxla("0501112233", "123456") == (False, "Kodun vaxtı keçib və ya kod göndərilməyib. Yenidən göndərin.")
    assert not H.kod_gonder("12", "sms")["ok"]


def test_xidmet_qurulmayan_server(monkeypatch):
    from assistant.modullar import qebul as QB
    monkeypatch.setattr(QB, "TOKEN", "")
    monkeypatch.delenv("WA_SMS_URL")
    monkeypatch.delenv("WA_KOD_TEST", raising=False)
    c = TestClient(app)
    assert 'data-u="mobil"' not in c.get("/giris").text                                # kod göndərilə bilmir: Mobil yoxdur
    c.post("/giris", data={"login": "aynur", "parol": "Parol1"})
    assert "xidməti qurulmayıb" in c.get("/qeydiyyat").text
    r = c.post("/qeydiyyat", data={"ad": "Aynur Məmmədova", "tel": "0501112233", "mail": "aynur@asf.gov.az", "peseler": ["muhasib"]},
               follow_redirects=False)
    assert r.status_code == 303 and H.al("aynur")["tel_tesdiq"] is False               # bloklanmır, nömrə təsdiqsiz
    assert H.tel_ile("0501112233") is None                                               # təsdiqsiz nömrə ilə mobil giriş yoxdur


def test_telefonda_kod_ekranda(monkeypatch):
    from assistant.modullar import qebul as QB
    monkeypatch.setattr(QB, "TOKEN", "")
    monkeypatch.delenv("WA_SMS_URL")
    monkeypatch.setattr(G, "YERLI", True)
    j = H.kod_gonder("0501112233", "sms")
    assert j["ok"] and len(j["ekranda"]) == 6 and "sınaq rejimi" in j["metn"]
    assert H.kod_yoxla("0501112233", j["ekranda"]) == (True, "")


def _telefonla_qeydiyyat(c, mail="leyla@asf.gov.az"):
    kod = _kod(c, "0501112233")
    c.post("/giris/mobil", data={"tel": "0501112233", "kod": kod})
    r = c.post("/qeydiyyat", data={"ad": "Leyla Həsənova", "mail": mail, "peseler": ["muhasib"]}, follow_redirects=False)
    assert r.status_code == 303
    c.put("/api/sohbetler/s1", json={"ad": "Oktyabr çıxarışı", "mesajlar": []})       # işin bir hissəsi


def test_telefonla_girib_poctla_qaldigi_yerden_davam():
    """Telefonla qeydiyyat (poçt məcburi) → çıxış → poçtla giriş: hesabın nömrəsinə gələn kodla eyni hesab, eyni söhbətlər."""
    c = TestClient(app)
    _telefonla_qeydiyyat(c)
    c.get("/cixis")
    p = TestClient(app)
    r = p.post("/giris", data={"login": "leyla@asf.gov.az", "parol": "PoctParol"}, follow_redirects=False)
    assert r.headers["location"] == "/giris/bagla"                                      # poçt telefonla yaranmış hesaba bağlıdır
    s = p.get("/giris/bagla").text
    assert "+994 50 *** ** 33" in s and "994501112233" not in s                         # nömrə gizlədilir
    assert p.post("/giris/bagla/kod", data={"kanal": "sms"}).json()["ok"]
    assert p.post("/giris/bagla", data={"kod": "000000" if GELEN["994501112233"] != "000000" else "1"}).status_code == 401
    assert p.post("/giris/bagla/kod", data={"kanal": "whatsapp"}).json()["ok"]
    r = p.post("/giris/bagla", data={"kod": GELEN["994501112233"]}, follow_redirects=False)
    assert r.status_code == 303
    v = p.get("/api/veziyyet").json()
    assert v["istifadeci"]["login"] == "994501112233" and v["profil"] == {"peseler": ["muhasib"]}
    assert [x["ad"] for x in p.get("/api/sohbetler").json()] == ["Oktyabr çıxarışı"]   # qaldığı yerdən
    q = TestClient(app)                                                                  # bağlantı təsdiqləndi: növbəti dəfə kodsuz
    q.post("/giris", data={"login": "leyla@asf.gov.az", "parol": "PoctParol"})
    assert q.get("/api/veziyyet").json()["istifadeci"]["login"] == "994501112233"


def test_bagli_poct_basqasi_terefinden_tutula_bilmez():
    c = TestClient(app)
    _telefonla_qeydiyyat(c, "aynur@asf.gov.az")
    d = TestClient(app)                                                                  # həqiqi sahib poçtla girir
    d.post("/giris", data={"login": "aynur@asf.gov.az", "parol": "PoctParol"})
    assert d.get("/api/veziyyet").status_code == 401                                    # kod olmadan o hesaba girmir
    f = TestClient(app)                                                                  # başqası həmin ünvanı yazmağa çalışır
    H._KODLAR.clear()
    GELEN.clear()
    f.post("/giris/kod", data={"tel": "0701234567", "kanal": "sms"})
    f.post("/giris/mobil", data={"tel": "0701234567", "kod": GELEN["994701234567"]})
    r = f.post("/qeydiyyat", data={"ad": "Yad Adam", "mail": "aynur@asf.gov.az", "peseler": ["muhasib"]})
    assert "başqa hesaba bağlıdır" in r.text                                             # artıq bağlı ünvanı yazmaq olmur


def test_telefonda_poctla_birbasa(monkeypatch):
    monkeypatch.setattr(G, "YERLI", True)
    c = TestClient(app)
    _telefonla_qeydiyyat(c)
    p = TestClient(app)
    p.post("/giris", data={"login": "leyla@asf.gov.az", "usul": "outlook"})
    assert p.get("/api/veziyyet").json()["istifadeci"]["login"] == "994501112233"       # telefonda kod lazım deyil
