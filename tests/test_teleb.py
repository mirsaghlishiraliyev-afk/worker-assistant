"""Diş həkimlərinin stoka tələbləri: mətndən oxuma, qeydiyyat, rəhbərin qərarı, icazələr."""
import datetime as dt
import pathlib
import sys
import tempfile

import pytest

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
from assistant.modullar import qebul as QB  # noqa: E402
from assistant.modullar import teleb as TELEB  # noqa: E402


class S:
    pass


@pytest.fixture(autouse=True)
def muhit(monkeypatch):
    monkeypatch.setattr(TELEB, "QOVLUQ", pathlib.Path(tempfile.mkdtemp()))
    monkeypatch.setattr(QB, "QOVLUQ", pathlib.Path(tempfile.mkdtemp()))
    monkeypatch.setattr(QB, "_indi", lambda: dt.datetime(2026, 11, 11, 9, 30))
    monkeypatch.setattr(QB, "TOKEN", "")
    monkeypatch.setattr(QB, "PHONE_ID", "")


@pytest.mark.parametrize("metn,mehsul,miqdar,vahid,tecili", [
    ("Stoka tələb: Kompozit plomb A2 — 5 ədəd, təcili", "Kompozit plomb A2", 5, "ədəd", True),
    ("tələb yaz anestetik artikain 20 ampul", "anestetik artikain", 20, "ampul", False),
    ("Implant (4.0x10) 2 ədəd lazımdır", "Implant (4.0x10)", 2, "ədəd", False),
    ("Birdəfəlik əlcək M - 3", "Birdəfəlik əlcək M", 3, "", False),
    ("Stoka tələb: Rentgen plyonkası bitib, təcili", "Rentgen plyonkası", None, "", True),
    ("Tibbi maska 2,5 qutu", "Tibbi maska", 2.5, "qutu", False),
])
def test_metnden_teleb(metn, mehsul, miqdar, vahid, tecili):
    assert TELEB.tehlil(metn) == {"mehsul": mehsul, "miqdar": miqdar, "vahid": vahid, "tecili": tecili}


def test_yaz_ve_siyahi():
    s = S()
    plan = TELEB.qayda(s, "Stoka tələb: Kompozit plomb A2 — 5 ədəd, təcili")
    assert plan[0] == "stok_teleb_yaz"
    n = TELEB.stok_teleb_yaz(s, **plan[1])
    assert n["metn"].startswith("Tələb T-001 qeydə alındı: Kompozit plomb A2 — 5 ədəd (təcili)")
    assert n["ui"]["variantlar"][0]["mesaj"] == "T-001 tələbini rəhbərə bildir"
    assert TELEB.qayda(s, "Stok tələblərini göstər") == ("stok_telebleri", {})
    n = TELEB.stok_telebleri(s)
    assert "1-i qərar gözləyir" in n["metn"]


def test_iki_addimla_yazmaq():
    s = S()
    plan = TELEB.qayda(s, "Stoka tələb yaz")
    n = TELEB.stok_teleb_yaz(s, **plan[1])
    assert n["metn"].startswith("Nə lazımdır?") and s.teleb_gozle
    plan = TELEB.qayda(s, "Anestetik 20 ampul")                                # növbəti cavab tələbdir
    assert plan == ("stok_teleb_yaz", {"mehsul": "Anestetik", "miqdar": 20, "vahid": "ampul", "tecili": False})
    TELEB.stok_teleb_yaz(s, **plan[1])
    assert not s.teleb_gozle and TELEB.qayda(s, "Qəbulları göstər") is None


def test_qerar_yalniz_rehber(monkeypatch):
    hekim, rehber = S(), S()
    hekim.istifadeci_login, hekim.istifadeci_adi = "aynur", "Aynur"
    rehber.istifadeci_login, rehber.istifadeci_adi = "resad", "Rəşad"
    monkeypatch.setattr(TELEB, "rehberdir", lambda s: s is rehber)
    TELEB.stok_teleb_yaz(hekim, "Plomb", 5, "ədəd")
    TELEB.stok_teleb_yaz(rehber, "Maska", 2, "qutu")
    assert "yalnız klinika rəhbəri" in TELEB.stok_teleb_qerar(hekim, "T-001", "təsdiq")["metn"]
    assert len(TELEB.stok_telebleri(hekim)["ui"]["setirler"]) == 1                # həkim yalnız özününkünü görür
    n = TELEB.stok_telebleri(rehber)
    assert len(n["ui"]["bloklar"][0]["setirler"]) == 2                            # rəhbər hamısını
    assert TELEB.qayda(rehber, "T-001 təsdiq") == ("stok_teleb_qerar", {"nomre": "T-001", "qerar": "tesdiq"})
    assert TELEB.qayda(rehber, "t2 rədd") == ("stok_teleb_qerar", {"nomre": "T-002", "qerar": "redd"})
    TELEB.stok_teleb_qerar(rehber, "T-001", "tesdiq")
    assert [t["status"] for t in TELEB.hamisi()] == ["tesdiq", "gozleyir"]


def test_rehbere_whatsapp():
    from assistant.modullar import stok as ST
    s = S()
    TELEB.stok_teleb_yaz(s, "Kompozit plomb A2", 5, "ədəd", tecili=True)
    assert "nömrəsi sistemdə yoxdur" in TELEB.stok_teleb_rehbere(s, "T-001")["metn"]
    ST.rehber_telefonu(s, "050 777 88 99")
    plan = TELEB.qayda(s, "T-001 tələbini rəhbərə bildir")
    assert plan == ("stok_teleb_rehbere", {"nomre": "T-001", "gonder": False})
    k = TELEB.stok_teleb_rehbere(s, **plan[1])["ui"]["kartlar"][0]
    assert "stoka tələb yazıb: T-001 Kompozit plomb A2 — 5 ədəd (təcili)" in k["metn"]
    assert k["link"].startswith("https://wa.me/994507778899")


def test_sablon_senedde():
    doc = (pathlib.Path(__file__).resolve().parents[1] / "deploy" / "whatsapp_sablonlar.md").read_text(encoding="utf-8")
    for ad, s in TELEB.SABLONLAR.items():
        assert ad in doc and s["metn"] in doc, ad


def test_bir_nece_teleb_birlikde():
    """«T-001,T-002 tələbini rəhbərə bildir» — hər iki tələb bir mesajda; «T-001, T-002 təsdiq» — hər ikisi."""
    from assistant.modullar import stok as ST
    s = S()
    TELEB.stok_teleb_yaz(s, "Əlcək", 5, "paket", tecili=True)
    TELEB.stok_teleb_yaz(s, "Ağrı kəsici", 2, "qutu")
    TELEB.stok_teleb_yaz(s, "Maska", 1, "qutu")
    ST.rehber_telefonu(s, "055 274 22 67")
    plan = TELEB.qayda(s, "T-001,T-002 tələbini rəhbərə bildir")
    assert plan == ("stok_teleb_rehbere", {"nomre": "T-001,T-002", "gonder": False})
    k = TELEB.stok_teleb_rehbere(s, **plan[1])["ui"]["kartlar"]
    assert len(k) == 1 and "T-001 Əlcək — 5 paket (təcili); T-002 Ağrı kəsici — 2 qutu" in k[0]["metn"] and "T-003" not in k[0]["metn"]
    plan = TELEB.qayda(s, "T-001, T-002 təsdiq")
    assert plan == ("stok_teleb_qerar", {"nomre": "T-001,T-002", "qerar": "tesdiq"})
    TELEB.stok_teleb_qerar(s, **plan[1])
    assert [t["status"] for t in TELEB.hamisi()] == ["tesdiq", "tesdiq", "gozleyir"]
    assert TELEB.qayda(s, "t3 rədd")[1]["nomre"] == "T-003"


def test_gozleme_rejiminde_emr_teleb_sayilmir():
    """«Stoka tələb yaz» düyməsindən sonra «T-001 və T-002 ni rəhbərə göndər» yeni tələb kimi yazılmamalıdır."""
    s = S()
    TELEB.stok_teleb_yaz(s, "Əlcək", 5, "paket")
    TELEB.stok_teleb_yaz(s, "Ağrı kəsici", 2, "qutu")
    TELEB.stok_teleb_yaz(s, "")                                                    # «Nə lazımdır?» — cavab gözlənilir
    assert s.teleb_gozle
    plan = TELEB.qayda(s, "T-001 və T-002 ni rəhbərə göndər")
    assert plan == ("stok_teleb_rehbere", {"nomre": "T-001,T-002", "gonder": True})
    assert not s.teleb_gozle                                                       # gözləmə bitdi
    assert TELEB.qayda(s, "Maska 3 qutu") is None                                  # artıq tələb kimi tutulmur
    assert len(TELEB.hamisi()) == 2


def test_oz_telebini_legv_edir(monkeypatch):
    hekim, yad = S(), S()
    hekim.istifadeci_login, hekim.istifadeci_adi = "gunel", "Günel"
    yad.istifadeci_login, yad.istifadeci_adi = "aynur", "Aynur"
    monkeypatch.setattr(TELEB, "rehberdir", lambda s: False)
    TELEB.stok_teleb_yaz(hekim, "Səhv yazılmış", None)
    assert "tapılmadı" in TELEB.stok_teleb_qerar(yad, "T-001", "legv")["metn"]     # başqasının tələbini ləğv edə bilmir
    plan = TELEB.qayda(hekim, "T-001 ləğv et")
    assert plan == ("stok_teleb_qerar", {"nomre": "T-001", "qerar": "legv"})
    assert TELEB.stok_teleb_qerar(hekim, **plan[1])["metn"] == "Oldu: T-001 — ləğv edildi."
    assert TELEB.stok_telebleri(hekim)["ui"]["setirler"][0][4] == "ləğv edildi"


def test_oz_telebini_silir(monkeypatch):
    s = S()
    s.istifadeci_login, s.istifadeci_adi = "gunel", "Günel"
    monkeypatch.setattr(TELEB, "rehberdir", lambda x: False)
    TELEB.stok_teleb_yaz(s, "Əlcək", 5, "paket")
    TELEB.stok_teleb_yaz(s, "T-001 və T-002 ni rəhbərə göndər", None)            # səhvən yazılmış tələb
    plan = TELEB.qayda(s, "T-002 sil")
    assert plan == ("stok_teleb_qerar", {"nomre": "T-002", "qerar": "sil"})
    assert TELEB.stok_teleb_qerar(s, **plan[1])["metn"] == "Oldu: T-002 — silindi."
    assert [x[0] for x in TELEB.stok_telebleri(s)["ui"]["setirler"]] == ["T-001"]   # siyahıda görünmür
    assert TELEB.stok_teleb_yaz(s, "Maska", 1, "qutu")["metn"].startswith("Tələb T-003")   # nömrə təkrarlanmır


def test_legv_edilmisi_silmek(monkeypatch):
    s = S()
    s.istifadeci_login, s.istifadeci_adi = "gunel", "Günel"
    monkeypatch.setattr(TELEB, "rehberdir", lambda x: False)
    TELEB.stok_teleb_yaz(s, "Səhv", None)
    TELEB.stok_teleb_qerar(s, "T-001", "legv")
    assert TELEB.stok_teleb_qerar(s, "T-001", "sil")["metn"] == "Oldu: T-001 — silindi."   # ləğv ediləndən sonra da silinir
    TELEB.stok_teleb_yaz(s, "Maska", 1, "qutu")
    TELEB.status_deyis({"T-002"}, "tesdiq")
    assert "silinmir" in TELEB.stok_teleb_qerar(s, "T-002", "sil")["metn"]                 # təsdiqlənmiş sifarişə düşüb


def test_gozleme_rejiminde_techizatci_teleb_sayilmir():
    s = S()
    s.teleb_gozle = True
    assert TELEB.qayda(s, "Təchizatçı: Qafarov Anar 055 710 43 34") is None
