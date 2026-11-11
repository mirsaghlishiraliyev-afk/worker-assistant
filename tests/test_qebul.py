"""Diş klinikası: qəbul jurnalı və WhatsApp mesajları (təsdiq, xatırlatma, həkimə xəbərdarlıq)."""
import datetime as dt
import pathlib
import sys
import tempfile
from urllib.parse import unquote

import pytest

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
from assistant.modullar import excel as EX  # noqa: E402
from assistant.modullar import qebul as QB  # noqa: E402

FAYL = pathlib.Path(__file__).parent / "data" / "qebul_test.xlsx"
INDI = dt.datetime(2026, 11, 11, 9, 0)


class S:
    def __init__(self):
        self.excel = EX.ac(FAYL, tempfile.mkdtemp())


@pytest.fixture(autouse=True)
def muhit(monkeypatch):
    monkeypatch.setattr(QB, "QOVLUQ", pathlib.Path(tempfile.mkdtemp()))
    monkeypatch.setattr(QB, "_indi", lambda: INDI)
    monkeypatch.setattr(QB, "TOKEN", "")
    monkeypatch.setattr(QB, "PHONE_ID", "")
    monkeypatch.setattr(QB, "XATIRLATMA_SAAT", [2.0])


def _api(monkeypatch):
    gonderilen = []
    monkeypatch.setattr(QB, "TOKEN", "t")
    monkeypatch.setattr(QB, "PHONE_ID", "123")
    monkeypatch.setattr(QB, "GONDERICI", lambda tel, sablon, p: gonderilen.append((tel, sablon, p)))
    return gonderilen


def test_telefon():
    assert QB.telefon("050 123 45 67") == "994501234567"
    assert QB.telefon("+994 70 555 66 77") == "994705556677"
    assert QB.telefon(501234567.0) == "994501234567"            # Excel nömrəni rəqəmə çevirir (sıfır itir)
    assert QB.telefon("00994551112233") == "994551112233"
    assert QB.telefon("12345") == "" and QB.telefon(None) == ""


def test_jurnal_oxunur():
    d = QB.qebullar(S())
    q = {x["xeste"]: x for x in d["qebullar"]}
    assert len(q) == 4
    a = q["Əliyev Murad"]
    assert a["tel"] == "994501234567" and a["vaxt"] == "12.11.2026, saat 10:00" and a["yazilma"] == "05.11.2026"
    assert a["hekim_tel"] == "994552223344" and a["problem"] == "Diş ağrısı (sol alt çənə)"
    assert q["Məmmədov Elçin"]["vaxt"] == "13.11.2026, saat 14:00"         # saat mətn kimi yazılıb
    assert q["Hüseynova Leyla"]["problemler"] == ["xəstənin telefonu yoxdur və ya oxunmur"]


def test_tesdiq_mesaji_linklerle():
    n = QB.qebul_mesajlari(S(), "tesdiq")
    k = n["ui"]["kartlar"]
    assert n["ui"]["tip"] == "whatsapp" and len(k) == 4 and "API qurulmayıb" in n["metn"]
    m = k[0]["metn"]
    for parca in ["Əliyev Murad", "05.11.2026 tarixində", "Dr. Həsənova həkimin qəbuluna", "12.11.2026, saat 10:00",
                  "Diş ağrısı", "Gecikəcəyiniz", "+994 55 222 33 44"]:
        assert parca in m, parca
    assert k[0]["link"].startswith("https://wa.me/994501234567?text=") and unquote(k[0]["link"].split("text=")[1]) == m
    assert next(x for x in k if x["ad"] == "Hüseynova Leyla")["status"] == "telefon yoxdur"
    assert "klinikanın qəbul şöbəsi" in next(x for x in k if x["ad"] == "Məmmədov Elçin")["metn"]   # həkimin nömrəsi yoxdur


def test_xatirlatma_pasiyent_ve_hekim(monkeypatch):
    monkeypatch.setattr(QB, "_indi", lambda: dt.datetime(2026, 11, 12, 8, 0))
    k = QB.qebul_mesajlari(S(), "xatirlatma", saat=4)["ui"]["kartlar"]   # 10:00 və 11:30 qəbulları
    assert [x["kime"] for x in k] == ["Pasiyent", "Həkim", "Pasiyent", "Həkim"]
    assert "Gecikəcəksinizsə" in k[0]["metn"] and "həkimi dərhal məlumatlandırın" in k[0]["metn"]
    h = k[1]["metn"]
    assert "Dr. Həsənova" in h and "Əliyev Murad qəbulunuzdadır" in h and "yubanarsa" in h and "+994 50 123 45 67" in h
    assert k[1]["link"].startswith("https://wa.me/994552223344")


def test_hekim_telefonu_yadda_qalir(monkeypatch):
    s = S()
    plan = QB.qayda(s, "Dr. Rzayev telefonu 050 999 88 77")
    assert plan == ("hekim_telefonu", {"hekim": "Dr. Rzayev", "telefon_nomresi": "050 999 88 77"})
    QB.hekim_telefonu(s, **plan[1])
    q = {x["xeste"]: x for x in QB.qebullar(s)["qebullar"]}
    assert q["Məmmədov Elçin"]["hekim_tel"] == "994509998877"


def test_api_ile_tesdiq_ve_gonderme(monkeypatch):
    gonderilen = _api(monkeypatch)
    s = S()
    n = QB.qebul_mesajlari(s, "tesdiq")
    assert not gonderilen and n["ui"]["sual"]["variantlar"][0]["mesaj"] == "Təsdiq mesajlarını göndər"   # əvvəl təsdiq
    plan = QB.qayda(s, "Təsdiq mesajlarını göndər")
    assert plan == ("qebul_mesajlari", {"nov": "tesdiq", "gonder": True})
    n = QB.qebul_mesajlari(s, **plan[1])
    assert len(gonderilen) == 3 and "3 göndərildi" in n["metn"]          # telefonu olmayan getmir
    tel, sablon, p = gonderilen[0]
    assert sablon == "qebul_tesdiq" and p[:3] == ["Əliyev Murad", "05.11.2026", "Dr. Həsənova"]
    n = QB.qebul_mesajlari(s, "tesdiq", gonder=True)                     # təkrar göndərmir
    assert len(gonderilen) == 3 and "3 əvvəl göndərilmişdi" in n["metn"]
    assert QB.qayda(s, "Mesajları göndərmə") == ("qebul_siyahi", {})


def test_api_xetasi_gosterilir(monkeypatch):
    _api(monkeypatch)

    def xeta(*a):
        raise RuntimeError("WhatsApp API 400: Template name does not exist")
    monkeypatch.setattr(QB, "GONDERICI", xeta)
    n = QB.qebul_mesajlari(S(), "tesdiq", gonder=True)
    assert "3 göndərilmədi" in n["metn"] and "Template name" in n["ui"]["kartlar"][0]["status"]


def test_avtomatik_xatirlatma(monkeypatch):
    gonderilen = _api(monkeypatch)
    s = S()
    assert "aktivdir: 4 qəbul" in QB.xatirlatma_aktiv(s)["metn"]
    assert QB.xatirlatmalari_isle(dt.datetime(2026, 11, 12, 7, 0)) == 0           # hələ tezdir (3 saat qalıb)
    assert QB.xatirlatmalari_isle(dt.datetime(2026, 11, 12, 8, 30)) == 2          # 10:00 qəbulu: pasiyent + həkim
    assert {g[1] for g in gonderilen} == {"qebul_xatirlatma", "hekim_xatirlatma"}
    assert QB.xatirlatmalari_isle(dt.datetime(2026, 11, 12, 8, 31)) == 0          # təkrar getmir
    assert QB.xatirlatmalari_isle(dt.datetime(2026, 11, 12, 10, 1)) == 2          # 11:30 qəbulu
    QB.xatirlatma_aktiv(s, aktiv=False)
    assert QB.xatirlatmalari_isle(dt.datetime(2026, 11, 13, 12, 30)) == 0         # dayandırılıb


def test_avtomatik_api_olmadan():
    assert "WhatsApp Business API qurulmalıdır" in QB.xatirlatma_aktiv(S())["metn"]


def test_sablonlar_senedde():
    """Meta-da yaradılacaq şablonların mətni sənəddə koddakı ilə eynidir."""
    doc = (pathlib.Path(__file__).resolve().parents[1] / "deploy" / "whatsapp_sablonlar.md").read_text(encoding="utf-8")
    for ad, s in QB.SABLONLAR.items():
        assert ad in doc and s["metn"] in doc, ad


def test_qebul_jurnali_olmayan_fayl():
    class B:
        excel = EX.ac(pathlib.Path(__file__).parent / "data" / "ixtiyari_excel.xlsx", tempfile.mkdtemp())
    assert QB.qayda(B(), "Təsdiq mesajlarını hazırla") is None
    assert "qəbul jurnalı tapılmadı" in QB.qebul_mesajlari(B())["metn"]


def test_xatirlatma_yaxin_qebulu_deyir():
    s = S()                                                                     # indi 11.11 09:00; ilk qəbul 12.11 10:00
    n = QB.qebul_mesajlari(s, "xatirlatma")
    assert "Ən yaxın qəbul: 12.11.2026, saat 10:00 — Əliyev Murad" in n["metn"]
    plan = QB.qayda(s, "72 saatlıq xatırlatmaları hazırla")
    assert plan[1]["saat"] == 72 and len(QB.qebul_mesajlari(s, **plan[1])["ui"]["kartlar"]) == 8   # 4 qəbul × (pasiyent + həkim)


def test_fayla_gore_hazir_sorgular():
    from assistant import teklif as TK
    s = S()
    s.bank, s.hazir, s.son_niyyet = None, False, None
    t = TK.hazirla(s)
    assert t[:3] == ["Qəbulları göstər", "Təsdiq mesajlarını hazırla", "Xatırlatmaları hazırla"]
