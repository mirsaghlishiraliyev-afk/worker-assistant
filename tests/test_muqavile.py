"""Müqavilə nömrəsi olmayan mədaxillər: müqavilə faylı ilə FİN üzrə tutuşdurma, mühasibin qərarı, inteqrasiyaya yazılması."""
import pathlib
import sys
import tempfile

import openpyxl

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
from assistant.modullar import bank as B  # noqa: E402

DATA = pathlib.Path(__file__).parent / "data"
MUQ = DATA / "muqavile_test.xlsx"
KAPITAL = DATA / "bank_kapital_test.xlsx"
QOV = pathlib.Path(tempfile.mkdtemp())


class K:
    def __init__(self):
        B.QAYDA_FAYLI, B.HESABAT_QOVLUGU = QOV / f"q{id(self)}.json", QOV
        self.bank = B.oxu(KAPITAL)
        self.muqavileler = {"ad": MUQ.name, "siyahi": B.muqavile_oxu(MUQ)}


def _setir(k, n):
    return next(r for r in k.bank["setirler"] if r["setir"] == n)


def test_fayl_taninir_ve_oxunur():
    assert B.muqavile_fayli_mi(MUQ)
    assert not B.muqavile_fayli_mi(KAPITAL) and not B.sablon_mi(MUQ)              # çıxarış və şablonla qarışmır
    ms = B.muqavile_oxu(MUQ)
    assert len(ms) == 8 and sum(m["uygun"] for m in ms) == 7                          # «Bitib» statuslu uyğun deyil
    assert ms[0] == {"no": "BS-2026000101-PS", "fin": "1AB2CD3", "ad": "NUMUNOVA LALE KAMIL QIZI", "aqrar": 500.0,
                     "status": "Ödəniş gözləyir", "uygun": True}


def test_tehlil_neticeleri():
    k = K()
    ms = k.muqavileler["siyahi"]
    n, v = B.muqavile_tehlil(_setir(k, 9), B._namizedler(_setir(k, 9), ms, set()))     # 911 = 500 + 411
    assert n.startswith("uyğundur") and v[0]["nomreler"] == ["BS-2026000101-PS", "BS-2026000102-PS"]
    assert all("BS-2026000104-PS" not in x["nomreler"] for x in v)                    # 911-lik, amma statusu «Bitib»
    n, v = B.muqavile_tehlil(_setir(k, 10), B._namizedler(_setir(k, 10), ms, set()))   # 20 = 20
    assert n == "üst-üstə düşür" and v[0]["nomreler"] == ["HS-2026000201-QS"]
    n, v = B.muqavile_tehlil(_setir(k, 14), B._namizedler(_setir(k, 14), ms, set()))   # 7060 ~ 7000
    assert n == "yaxındır" and v[0]["nomreler"] == ["BS-2026000301-AS"] and v[0]["ferq"] == 60.0
    r15 = _setir(k, 15)                                                               # FİN yoxdur — ad ilə tapılır
    n, v = B.muqavile_tehlil(r15, B._namizedler(r15, ms, set()))
    assert n == "uyğun məbləğ yoxdur" and v[0]["nomreler"] == ["BS-2026000401-PS"]


def test_muhasibin_qerari_inteqrasiyaya_yazilir():
    k = K()
    B.bank_hesabat(k)
    n = B.bank_hesabat(k, ehate="daxilolma")
    gorulen = []
    while k.bank.get("muq_gozle") is not None:                                         # mədaxillər bir-bir soruşulur
        s = k.bank["muq_gozle"]
        gorulen.append(s)
        if s == 9:
            assert "uyğundur" in n["metn"] and n["ui"]["variantlar"][0]["mesaj"] == "9-ci sətir müqaviləsi BS-2026000101-PS, BS-2026000102-PS"
            plan = B.qayda(k, n["ui"]["variantlar"][0]["mesaj"])                     # mühasib təklifi qəbul edir
        elif s == 10:
            plan = B.qayda(k, "HS-2026000201-QS")                                     # nömrəni özü yazır
        elif s == 14:
            assert "yaxındır" in n["metn"]
            plan = B.qayda(k, "14-cü sətir müqaviləsiz qalsın")                       # mühasib rədd edir
        else:
            plan = B.qayda(k, "Müqaviləsiz qalsın")
        n = getattr(B, plan[0])(k, **plan[1])
    assert gorulen == [9, 10, 14, 15]
    while k.bank.get("madde_gozle"):                                                    # maddələr
        n = B.bank_madde(k, k.bank["madde_gozle"]["kod"], "M")
    assert "2 mədaxilə müqavilə nömrəsi yazıldı (3 sətir)" in n["metn"]
    rows = list(openpyxl.load_workbook(QOV / n["ui"]["fayllar"][0]["ad"])["Sheet1"].iter_rows(values_only=True))[1:]
    nom = {r[2]: r[4] for r in rows if r[2]}
    assert nom == {"BS-2026000101-PS": 500.0, "BS-2026000102-PS": 411.0, "HS-2026000201-QS": 20.0, "HS-2026000001-QS": 2700.0}
    assert len(rows) == 10                                                              # 9 daxilolma, 911 iki müqaviləyə bölündü


def test_qalanlari_kec_ve_yalniz_sigorta():
    k = K()
    B.bank_hesabat(k)
    n = B.bank_hesabat(k, ehate="sigorta")
    assert k.bank["muq_gozle"] == 9                                                    # yalnız sığorta haqqı soruşulur
    n = B.bank_muqavile(k, **B.qayda(k, "Qalan müqavilələri keç")[1])
    assert k.bank.get("muq_gozle") is None and "qalan mədaxillər müqaviləsiz" in n["metn"]
