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
        else:
            plan = B.qayda(k, "Müqaviləsiz qalsın")
        n = getattr(B, plan[0])(k, **plan[1])
    assert gorulen == [9, 15]                                                     # 14 — qaytarılma, 10 — təlim: müqaviləyə aid deyil
    while k.bank.get("madde_gozle"):                                                    # maddələr
        n = B.bank_madde(k, k.bank["madde_gozle"]["kod"], "M")
    assert "1 mədaxilə müqavilə nömrəsi yazıldı (2 sətir)" in n["metn"]
    rows = list(openpyxl.load_workbook(QOV / n["ui"]["fayllar"][0]["ad"])["Sheet1"].iter_rows(values_only=True))[1:]
    nom = {r[2]: r[4] for r in rows if r[2]}
    assert nom == {"BS-2026000101-PS": 500.0, "BS-2026000102-PS": 411.0, "HS-2026000001-QS": 2700.0}
    assert len(rows) == 10                                                              # 9 daxilolma, 911 iki müqaviləyə bölündü


def test_qalanlari_kec_ve_yalniz_sigorta():
    k = K()
    B.bank_hesabat(k)
    n = B.bank_hesabat(k, ehate="sigorta")
    assert k.bank["muq_gozle"] == 9                                                    # yalnız sığorta haqqı soruşulur
    n = B.bank_muqavile(k, **B.qayda(k, "Qalan müqavilələri keç")[1])
    assert k.bank.get("muq_gozle") is None and "qalan mədaxillər müqaviləsiz" in n["metn"]


def test_hamisina_butun_muqavileler():
    """Mühasib «hamısına» yazır: mədaxil bu FİN-in bütün müqavilələrinə; cəmdən artıq qalan müqaviləsiz ayrıca sətir."""
    k = K()
    B.bank_hesabat(k)
    B.bank_hesabat(k, ehate="daxilolma")
    assert k.bank["muq_gozle"] == 9
    assert any(v["etiket"] == "Hamısına (3 müqavilə; cəmi 1 031,00)" for v in B._muq_suali(
        k.bank, next(r for r in k.bank["setirler"] if r["setir"] == 9),
        B._namizedler(next(r for r in k.bank["setirler"] if r["setir"] == 9), k.muqavileler["siyahi"], set()), 1, 4)["ui"]["variantlar"])
    plan = B.qayda(k, "hamısına")
    assert plan == ("bank_muqavile", {"setir": 9, "muqavile": "hamisi"})
    n = B.bank_muqavile(k, **plan[1])                                           # 500 + 411 + 120 = 1031 > 911: sıra ilə paylanır
    assert k.bank["muq_qerar"]["9"] == [("BS-2026000101-PS", 500.0), ("BS-2026000102-PS", 411.0)]
    assert "pul müqavilələrə sıra ilə paylandı, 1 müqaviləyə çatmadı" in n["metn"]
    assert k.bank["muq_gozle"] == 15
    n = B.bank_muqavile(k, **B.qayda(k, "nə?")[1])                                # başa düşülməyən cavab: sual təkrarlanır
    assert n["metn"].startswith("Cavabı başa düşmədim") and k.bank["muq_gozle"] == 15
    assert B.qayda(k, "Bütün daxilolmalar")[1]["muqavile"] == "hamisi"           # cədvələ keçmir


def test_hamisina_artiq_odenis():
    k = K()
    for m in k.muqavileler["siyahi"]:
        if m["fin"] == "1AB2CD3":
            m["aqrar"] = 300.0                                                      # 3 × 300 = 900 < 911
    B.bank_hesabat(k)
    B.bank_hesabat(k, ehate="daxilolma")
    n = B.bank_muqavile(k, 9, "hamisi")
    assert k.bank["muq_qerar"]["9"] == [("BS-2026000101-PS", 300.0), ("BS-2026000102-PS", 300.0), ("BS-2026000103-PS", 300.0), ("", 11.0)]
    assert "artıq qalan 11,00 müqaviləsiz ayrıca sətirdir" in n["metn"]


def test_qaytarilma_muqavileye_aid_deyil():
    """Vəsaitin qaytarılması (və depozit, faiz, təlim...) müqavilə ilə tutuşdurulmur, FİN uyğun gəlsə belə."""
    k = K()
    B.bank_hesabat(k)
    B.bank_hesabat(k, ehate="daxilolma")
    gorulen = []
    while k.bank.get("muq_gozle") is not None:
        gorulen.append(k.bank["muq_gozle"])
        B.bank_muqavile(k, k.bank["muq_gozle"], "yox")
    assert 14 not in gorulen and 10 not in gorulen                                # qaytarılma və təlim


def test_teyinata_gore_hs_ve_bs():
    """Təyinatda «heyvan» — yalnız HS- müqavilələr; «bitki»/«ting» — yalnız BS-; heç biri yoxdursa — hamısı."""
    ms = [{"no": "HS-2026000001-QS", "fin": "1AB2CD3", "ad": "", "aqrar": 50.0, "status": "Aktiv", "uygun": True},
          {"no": "BS-2026000002-PS", "fin": "1AB2CD3", "ad": "", "aqrar": 50.0, "status": "Aktiv", "uygun": True}]
    r = {"fin": "1AB2CD3", "mebleg": 50.0, "odeyici": "", "teyinat_xam": ""}
    nomre = lambda t: [m["no"] for m in B._namizedler(dict(r, teyinat_xam=t), ms, set())]  # noqa: E731
    assert nomre("Sığorta ödənişi DAMAZLIQ HEYVANLARIN SIGORTA ODENISI") == ["HS-2026000001-QS"]
    assert nomre("Bitkiçilik sığortası") == ["BS-2026000002-PS"]
    assert nomre("alma ting sığortası") == ["BS-2026000002-PS"]
    assert nomre("SIĞORTA") == ["HS-2026000001-QS", "BS-2026000002-PS"]
    k = K()
    r9 = next(x for x in k.bank["setirler"] if x["setir"] == 9)
    r9["teyinat_xam"] = "heyvan sigortasi"                                        # 1AB2CD3-ün müqavilələri BS-dir
    assert B._namizedler(r9, k.muqavileler["siyahi"], set()) == []
    n = B.bank_muqavile(k, 9, "BS-2026000101-PS")
    assert "heyvandarlıq" not in n["metn"] and "bitkiçilik (BS-) müqaviləsidir — yoxlayın" in n["metn"]
