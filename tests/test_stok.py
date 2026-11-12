"""Diş klinikası: stok nəzarəti, rəhbərə bildiriş, təchizatçılara WhatsApp sifarişi."""
import datetime as dt
import pathlib
import sys
import tempfile

import pytest

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
from assistant.modullar import excel as EX  # noqa: E402
from assistant.modullar import qebul as QB  # noqa: E402
from assistant.modullar import stok as ST  # noqa: E402
from assistant.modullar import teleb as TELEB  # noqa: E402

FAYL = pathlib.Path(__file__).parent / "data" / "stok_test.xlsx"


class S:
    def __init__(self):
        self.excel = EX.ac(FAYL, tempfile.mkdtemp())


@pytest.fixture(autouse=True)
def muhit(monkeypatch):
    monkeypatch.setattr(QB, "QOVLUQ", pathlib.Path(tempfile.mkdtemp()))
    monkeypatch.setattr(QB, "_indi", lambda: dt.datetime(2026, 11, 11, 9, 0))
    monkeypatch.setattr(TELEB, "QOVLUQ", pathlib.Path(tempfile.mkdtemp()))
    monkeypatch.setattr(QB, "TOKEN", "")
    monkeypatch.setattr(QB, "PHONE_ID", "")
    monkeypatch.delenv("WA_REHBER_TEL", raising=False)


def test_azalanlar_ve_sifaris_miqdari():
    d = ST.stok(S())
    m = {x["mehsul"]: x for x in d["mehsullar"]}
    assert m["Birdəfəlik əlcək (M)"]["veziyyet"] == "bitib"
    assert m["Kompozit plomb materialı (A2)"]["veziyyet"] == "azdır" and m["Kompozit plomb materialı (A2)"]["sifaris"] == 10  # sütundan
    assert m["Anestetik (artikain) 1,7 ml"]["sifaris"] == 32                       # sütun boş: 2×20 − 8
    assert m["Tibbi maska"]["veziyyet"] == "normal" and m["Tibbi maska"]["sifaris"] is None
    az = [x["mehsul"] for x in ST._azalanlar(d)]
    assert az[0] == "Birdəfəlik əlcək (M)" and "Tibbi maska" not in az and len(az) == 4   # bitən birinci


def test_stok_veziyyeti_duymelerle():
    n = ST.stok_veziyyeti(S())
    assert "4 məhsul" in n["metn"] and "1-i bitib" in n["metn"]
    cedvel, sual = n["ui"]["bloklar"]
    assert cedvel["setirler"][0][3] == "⛔ bitib"
    assert [v["mesaj"] for v in sual["variantlar"]] == ["Rəhbərə stok bildirişi hazırla", "Təchizatçılara sifariş hazırla"]


def test_rehbere_bildiris():
    s = S()
    n = ST.stok_rehbere(s)
    assert "Rəhbərin WhatsApp nömrəsi yoxdur" in n["metn"]
    plan = ST.qayda(s, "Rəhbərin telefonu 050 777 88 99")
    assert plan == ("rehber_telefonu", {"telefon_nomresi": "050 777 88 99"})
    ST.rehber_telefonu(s, **plan[1])
    k = ST.stok_rehbere(s)["ui"]["kartlar"][0]
    assert k["link"].startswith("https://wa.me/994507778899")
    assert "4 məhsul minimum həddən azdır" in k["metn"] and "Birdəfəlik əlcək (M) — qalıq 0 qutu (min 4)" in k["metn"]


def test_techizatcilara_sifaris():
    k = {x["ad"]: x for x in ST.stok_sifaris(S())["ui"]["kartlar"]}
    assert set(k) == {"DentalMed MMC", "MedSupply", "ImplantPro"}                    # hər təchizatçıya ayrıca
    assert "Kompozit plomb materialı (A2) — 10 ədəd; Anestetik (artikain) 1,7 ml — 32 ampul" in k["DentalMed MMC"]["metn"]
    assert "Endodontik" not in k["MedSupply"]["metn"] and "Birdəfəlik əlcək (M) — 6 qutu" in k["MedSupply"]["metn"]
    assert k["ImplantPro"]["status"] == "telefon yoxdur"
    assert k["DentalMed MMC"]["link"].startswith("https://wa.me/994501112233")


def test_api_ile_tesdiq_sonra_gonderir(monkeypatch):
    gonderilen = []
    monkeypatch.setattr(QB, "TOKEN", "t")
    monkeypatch.setattr(QB, "PHONE_ID", "1")
    monkeypatch.setattr(QB, "GONDERICI", lambda tel, sablon, p: gonderilen.append((tel, sablon, p)))
    s = S()
    n = ST.stok_sifaris(s)
    assert not gonderilen and n["ui"]["sual"]["variantlar"][0]["mesaj"] == "Təchizatçılara sifarişi göndər"
    plan = ST.qayda(s, "Təchizatçılara sifarişi göndər")
    assert plan == ("stok_sifaris", {"gonder": True})
    n = ST.stok_sifaris(s, **plan[1])
    assert len(gonderilen) == 2 and "2 göndərildi" in n["metn"]                     # ImplantPro-nun telefonu yoxdur
    assert all(g[1] == "stok_sifaris" for g in gonderilen)
    ST.stok_sifaris(s, gonder=True)
    assert len(gonderilen) == 2                                                     # eyni gün eyni sifariş təkrar getmir
    assert ST.qayda(s, "Stok mesajlarını göndərmə") == ("stok_veziyyeti", {})


def test_lokal_sorgular_ve_hazir_sorgular():
    from assistant import teklif as TK
    s = S()
    assert ST.qayda(s, "Azalan məhsulları göstər") == ("stok_veziyyeti", {"yalniz_azalan": True})
    assert ST.qayda(s, "Rəhbərə stok bildirişi hazırla") == ("stok_rehbere", {"gonder": False})
    s.bank, s.hazir, s.son_niyyet = None, False, None
    assert TK.hazirla(s)[:4] == ["Stok vəziyyəti", "Azalan məhsulları göstər (4)", "Rəhbərə stok bildirişi hazırla",
                                 "Təchizatçılara sifariş hazırla"]


def test_stok_olmayan_fayl():
    class B:
        excel = EX.ac(pathlib.Path(__file__).parent / "data" / "qebul_test.xlsx", tempfile.mkdtemp())
    assert ST.qayda(B(), "Stok vəziyyəti") is None
    assert "stok cədvəli tapılmadı" in ST.stok_veziyyeti(B())["metn"]


def test_sablonlar_senedde():
    doc = (pathlib.Path(__file__).resolve().parents[1] / "deploy" / "whatsapp_sablonlar.md").read_text(encoding="utf-8")
    for ad, s in ST.SABLONLAR.items():
        assert ad in doc and s["metn"] in doc, ad


def test_tesdiqlenmis_telebler_sifarise_dusur():
    s = S()
    TELEB.stok_teleb_yaz(s, "Tibbi maska", 10, "qutu")                  # stokda var, azalmayıb: həkim istəyir
    TELEB.stok_teleb_yaz(s, "kompozit plomb", 15, "ədəd")              # azalan məhsul: miqdarın böyüyü
    TELEB.stok_teleb_yaz(s, "Rentgen plyonkası", 50, "ədəd")           # stokda yoxdur
    TELEB.stok_teleb_yaz(s, "Endodontik fayl dəsti", 2, "dəst")         # rədd ediləcək
    TELEB.stok_teleb_qerar(s, "T-004", "rədd")
    TELEB.stok_teleb_qerar(s, "hamisi", "təsdiq")
    n = ST.stok_sifaris(s)
    k = {x["ad"]: x["metn"] for x in n["ui"]["kartlar"]}
    assert "Tibbi maska — 10 qutu" in k["MedSupply"] and "Endodontik" not in k["MedSupply"]
    assert "Kompozit plomb materialı (A2) — 15 ədəd" in k["DentalMed MMC"]
    assert "Rentgen plyonkası — 50 ədəd" in k["Təchizatçı göstərilməyib"]
    assert "T-001, T-002, T-003" in n["metn"]


def test_gonderilen_sifarisin_telebleri_isarelenir(monkeypatch):
    monkeypatch.setattr(QB, "TOKEN", "t")
    monkeypatch.setattr(QB, "PHONE_ID", "1")
    monkeypatch.setattr(QB, "GONDERICI", lambda *a: None)
    s = S()
    TELEB.stok_teleb_yaz(s, "Tibbi maska", 10, "qutu")
    TELEB.stok_teleb_qerar(s, "T-001", "təsdiq")
    ST.stok_sifaris(s, gonder=True)
    assert TELEB.hamisi()[0]["status"] == "sifaris"


def test_stok_fayli_olmadan_tesdiqlenmis_telebler():
    """Rəhbər tələbi təsdiq edib, stok faylı açıq deyil: sifariş tələblərdən hazırlanır, təchizatçı yaddaşdan."""
    class Bos:
        excel = None
    bos = Bos()
    assert "Təsdiqlənmiş stok tələbi yoxdur" in ST.stok_sifaris(bos)["metn"]
    TELEB.stok_teleb_yaz(bos, "Kompozit plomb", 7, "ədəd")
    TELEB.stok_teleb_yaz(bos, "Rentgen plyonkası", 50, "ədəd")
    TELEB.stok_teleb_qerar(bos, "hamisi", "təsdiq")
    plan = ST.qayda(bos, "Təchizatçılara sifariş hazırla")                     # fayl yoxdur, yenə də başa düşülür
    assert plan == ("stok_sifaris", {"gonder": False})
    n = ST.stok_sifaris(bos, **plan[1])                                        # təchizatçı hələ bilinmir
    assert n["metn"].startswith("Stok faylı açıq deyil") and "Təchizatçı: " in n["metn"]
    assert [k["ad"] for k in n["ui"]["kartlar"]] == ["Təchizatçı göstərilməyib"]
    ST.stok(S())                                                               # stok faylı bir dəfə yükləndi — yadda qaldı
    n = ST.stok_sifaris(bos)
    k = {x["ad"]: x for x in n["ui"]["kartlar"]}
    assert "Kompozit plomb materialı (A2) — 7 ədəd" in k["DentalMed MMC"]["metn"]
    assert k["DentalMed MMC"]["link"].startswith("https://wa.me/994501112233")
    assert "Rentgen plyonkası — 50 ədəd" in k["Təchizatçı göstərilməyib"]["metn"]


def test_techizatcini_sohbetde_yazmaq():
    """Təchizatçı göstərilməyib → «Təchizatçı: Ad Soyad 055 ...» yazılır, yadda qalır, sifariş ona hazırlanır."""
    class Bos:
        excel = None
    bos = Bos()
    TELEB.stok_teleb_yaz(bos, "Əlcək", 5, "paket")
    TELEB.stok_teleb_yaz(bos, "Ağrı kəsici", 2, "qutu")
    TELEB.stok_teleb_qerar(bos, "hamisi", "təsdiq")
    assert [k["ad"] for k in ST.stok_sifaris(bos)["ui"]["kartlar"]] == ["Təchizatçı göstərilməyib"]
    plan = ST.qayda(bos, "Təchizatçı: Qafarov Anar 055 710 43 34")
    assert plan == ("techizatci_teyin", {"techizatci": "Qafarov Anar", "telefon": "055 710 43 34"})
    n = ST.techizatci_teyin(bos, **plan[1])
    assert n["metn"].startswith("Oldu: Əlcək, Ağrı kəsici — təchizatçı Qafarov Anar")
    k = n["ui"]["kartlar"]
    assert [x["ad"] for x in k] == ["Qafarov Anar"] and k[0]["link"].startswith("https://wa.me/994557104334")
    assert "Salam, Qafarov Anar!" in k[0]["metn"]
    assert ST.qayda(bos, "Təchizatçı Qafarov Anar 0557104334")[0] == "techizatci_teyin"
    assert ST.qayda(bos, "Təchizatçılara sifariş hazırla")[0] == "stok_sifaris"   # adi əmr qarışmır
    assert [x["ad"] for x in ST.stok_sifaris(bos)["ui"]["kartlar"]] == ["Qafarov Anar"]   # yadda qaldı


def test_telefonu_olmayan_techizatciya_telefon():
    s = S()
    n = ST.stok_sifaris(s)
    telsiz = [k["ad"] for k in n["ui"]["kartlar"] if k.get("status") == "telefon yoxdur" and k["ad"] != "Təchizatçı göstərilməyib"]
    if not telsiz:
        pytest.skip("test faylında telefonsuz təchizatçı yoxdur")
    n = ST.techizatci_teyin(s, telsiz[0], "050 999 88 77")
    k = {x["ad"]: x for x in n["ui"]["kartlar"]}
    assert k[telsiz[0]]["link"].startswith("https://wa.me/994509998877")
