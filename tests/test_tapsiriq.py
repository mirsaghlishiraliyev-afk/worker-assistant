"""Tapşırıqlar: rəhbər tapşırıq verir və izləyir, işçi icra edir, aralarında təsdiq və məlumat sorğuları."""
import datetime as dt
import pathlib
import sys
import tempfile

import pytest

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
from assistant.modullar import qebul as QB  # noqa: E402
from assistant.modullar import tapsiriq as TP  # noqa: E402

INDI = dt.datetime(2026, 11, 11, 10, 0)


class S:
    def __init__(self, login, ad):
        self.istifadeci_login, self.istifadeci_adi = login, ad


@pytest.fixture(autouse=True)
def muhit(monkeypatch):
    monkeypatch.setattr(TP, "QOVLUQ", pathlib.Path(tempfile.mkdtemp()))
    monkeypatch.setattr(QB, "_indi", lambda: INDI)
    monkeypatch.setattr(TP, "rehberdir", lambda s: s.istifadeci_login == "resad@asf.gov.az")
    TP.qeyd_et("resad@asf.gov.az", "Rəşad Əliyev")
    TP.qeyd_et("aynur@asf.gov.az", "Aynur Məmmədova")
    TP.qeyd_et("aysel@asf.gov.az", "Aysel Quliyeva")


REHBER, ISCI = (lambda: S("resad@asf.gov.az", "Rəşad Əliyev")), (lambda: S("aynur@asf.gov.az", "Aynur Məmmədova"))


def _icra(s, mesaj):
    plan = TP.qayda(s, mesaj)
    assert plan, mesaj
    return getattr(TP, plan[0])(s, **plan[1])


@pytest.mark.parametrize("metn,gozlenen", [
    ("15.11.2026", dt.date(2026, 11, 15)), ("15.11", dt.date(2026, 11, 15)), ("01.02", dt.date(2027, 2, 1)),
    ("3 gün", dt.date(2026, 11, 14)), ("2 həftə", dt.date(2026, 11, 25)), ("sabah", dt.date(2026, 11, 12)),
    ("bu gün", dt.date(2026, 11, 11)), ("birigün", dt.date(2026, 11, 13)), ("nə vaxtsa", None)])
def test_muddet(metn, gozlenen):
    assert TP.muddet_oxu(metn) == gozlenen


def test_isci_adi_ile_tapilir():
    assert TP.istifadeci_tap("Aynur Məmmədova") == ("aynur@asf.gov.az", "Aynur Məmmədova")
    assert TP.istifadeci_tap("Aynura") == ("aynur@asf.gov.az", "Aynur Məmmədova")                     # yönlük şəkilçisi
    assert TP.istifadeci_tap("Məmmədovaya")[0] == "aynur@asf.gov.az"
    assert TP.istifadeci_tap("yeni.isci@asf.gov.az") == ("yeni.isci@asf.gov.az", "yeni.isci@asf.gov.az")  # hələ girməyib
    TP.qeyd_et("aynur.h@asf.gov.az", "Aynur Həsənova")
    with pytest.raises(ValueError, match="bir neçə"):                                                  # iki Aynur
        TP.istifadeci_tap("Aynur")
    assert TP.istifadeci_tap("Aynur Həsənova")[0] == "aynur.h@asf.gov.az"
    with pytest.raises(ValueError, match="tapılmadı"):
        TP.istifadeci_tap("Kamran")


def test_tam_axin():
    r, i = REHBER(), ISCI()
    # rəhbər tapşırıq verir (adla, müddətlə)
    n = _icra(r, "Aynura tapşırıq: Oktyabr hesabatını hazırla, müddət 15.11.2026")
    assert n["metn"].startswith("Tapşırıq TP-001 verildi: Aynur Məmmədova — «Oktyabr hesabatını hazırla», müddət 15.11.2026")
    # işçi görür, icraya başlayır, addım yazır
    n = _icra(i, "Tapşırıqlarım")
    assert n["ui"]["bloklar"][0]["basliq"] == "Mənə verilən tapşırıqlar" and "1 yeni" in n["metn"]
    assert TP.bildirisler("aynur@asf.gov.az")["yeni"] == 1
    _icra(i, "TP-001 icraya başladım")
    _icra(i, "TP-001 addım: məlumatlar toplandı")
    # işçi tapşırığa bağlı təsdiq sorğusu göndərir — rəhbərə gedir
    n = _icra(i, "TP-001 üzrə təsdiq sorğusu: hesabatı PDF göndərim?")
    assert "Təsdiq sorğusu S-001 göndərildi: Rəşad Əliyev (TP-001 üzrə)" in n["metn"]
    assert TP.bildirisler("resad@asf.gov.az")["sorgu"] == 1
    # rəhbər cavab verir
    n = _icra(r, "S-001 təsdiq")
    assert n["metn"].startswith("Oldu: S-001 — təsdiqləndi")
    n = _icra(i, "Sorğular")
    assert n["ui"]["bloklar"][0]["setirler"][0][6] == "təsdiqləndi"
    _icra(i, "TP-001 tamamlandı: hesabat göndərildi")
    # rəhbər izləyir: addımların tarixçəsi
    n = _icra(r, "TP-001")
    addimlar = [a[2] for a in n["ui"]["bloklar"][1]["setirler"]]
    assert addimlar == ["Tapşırıq verildi", "İcraya başlanıldı", "məlumatlar toplandı", "Təsdiq sorğusu S-001: hesabatı PDF göndərim?",
                        "S-001 — təsdiqləndi", "Tamamlandı: hesabat göndərildi"]
    assert "tamamlandı" in n["metn"]


def test_ayrica_melumat_sorgusu_ve_cavab():
    r, i = REHBER(), ISCI()
    _icra(r, "aynur@asf.gov.az tapşırıq: Müqaviləni yoxla, müddət 3 gün")
    n = _icra(i, "Məlumat sorğusu: məzuniyyət qrafiki nə vaxt təsdiqlənəcək?")    # tapşırığa bağlı deyil
    assert "S-001 göndərildi: Rəşad Əliyev —" in n["metn"]                         # rəhbər: ona tapşırıq verən
    _icra(r, "S-001 cavab yaz")
    assert r.tp_gozle == {"nov": "cavab", "no": "S-001"}
    n = _icra(r, "Gələn həftə")                                                   # növbəti mesaj cavabdır
    assert "S-001 — cavablandı: «Gələn həftə»" in n["metn"]


def test_huquqlar():
    r, i = REHBER(), ISCI()
    assert "Rəhbər" in _icra(i, "Rəşada tapşırıq: test")["metn"]                    # işçi tapşırıq verə bilmir
    _icra(r, "Aynura tapşırıq: Hesabat, müddət sabah")
    assert "yalnız onu verən rəhbər" in _icra(i, "TP-001 ləğv et")["metn"]
    assert TP.tapsiriq_muddet(i, "TP-001", "20.11")["metn"].startswith("Müddəti yalnız")
    yad = S("aysel@asf.gov.az", "Aysel Quliyeva")
    assert "tapılmadı" in _icra(yad, "TP-001")["metn"]                             # başqasının tapşırığını görmür
    assert "tapılmadı" in TP.sorgu_cavab(yad, "S-001", "tesdiq")["metn"]


def test_gecikme(monkeypatch):
    r, i = REHBER(), ISCI()
    _icra(r, "Aynura tapşırıq: Hesabat, müddət 12.11.2026")
    monkeypatch.setattr(QB, "_indi", lambda: dt.datetime(2026, 11, 13, 9, 0))
    n = _icra(r, "Tapşırıqlar")
    assert "1 gecikir" in n["metn"] and "gecikir" in n["ui"]["bloklar"][0]["setirler"][0][4]
    assert TP.bildirisler("aynur@asf.gov.az")["gecikir"] == 1
    _icra(r, "TP-001 müddət 20.11.2026")
    assert TP.bildirisler("aynur@asf.gov.az")["gecikir"] == 0
