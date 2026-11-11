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
