"""Bank daxilolmaları: çıxarışın tanınması, təsnifat, öyrənmə, xəbərdarlıqlar, 1C faylı."""
import pathlib
import sys
import tempfile

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
from assistant.modullar import bank as B  # noqa: E402
from assistant.modullar import excel as EX  # noqa: E402

FAYL = pathlib.Path(__file__).parent / "data" / "bank_cixarisi_test.xlsx"
QOV = pathlib.Path(tempfile.mkdtemp())


class S:
    def __init__(self, ad="q.json"):
        B.QAYDA_FAYLI = QOV / ad
        B.HESABAT_QOVLUGU = QOV
        self.bank = B.oxu(FAYL)


def _madde_deyilib():  # mühasib bütün kateqoriyaların maddəsini deyib
    B._qaydalar_yaz({**B._qaydalar(), "madde": {k["kod"]: k["teklif"] for k in B.KATEQORIYALAR}})


def _hazirla(s, ehate="sigorta", **kw):  # 1C faylı: əvvəl «hansı sətirlər?», sonra hər kateqoriyanın maddəsi soruşulur
    n = B.bank_hesabat(s, **kw)
    while n["ui"]["tip"] == "sual":
        if s.bank.get("ehate_gozle") is not None:
            n = B.bank_hesabat(s, ehate=ehate, **kw)
            continue
        kod = s.bank["madde_gozle"]["kod"]
        n = B.bank_madde(s, kod, (B._teklifler(kod) or [f"M {kod}"])[0])
    return n


def _kat(s, setir):
    return next(r for r in s.bank["setirler"] if r["setir"] == setir)["kateqoriya"]


def test_tanima_ve_oxuma():
    s = S()
    assert s.bank["basliq_setri"] == 4
    assert len(s.bank["setirler"]) == 12                      # debet və «Cəmi» sətri atılır
    assert round(sum(r["mebleg"] for r in s.bank["setirler"]), 2) == 27608.45  # «3 400,00» mətn kimi yazılıb
    assert B.oxu(pathlib.Path(__file__).parent / "data" / "ixtiyari_excel.xlsx") is None


def test_tesnifat():
    s = S()
    assert [_kat(s, i) for i in (5, 6, 7)] == ["sigorta_haqqi"] * 3
    assert _kat(s, 9) == "subroqasiya"
    assert _kat(s, 10) == "emek_haqqi" and _kat(s, 11) == "ezamiyye"
    assert _kat(s, 12) == "yoxlanmali"                         # xəzinədarlıq, amma təyinat aydın deyil
    assert _kat(s, 13) == "telim"
    assert _kat(s, 8) == "yoxlanmali"                          # «Hesaba mədaxil»
    r = next(r for r in s.bank["setirler"] if r["setir"] == 5)
    assert r["muqavile"] == "HS-2026000123-QS"


def test_xeberdarliqlar():
    s = S()
    nov = {(a, b) for a, b, _ in s.bank["xeberdarliqlar"]}
    assert (15, "Başqa ay") in nov
    assert (16, "Təkrar ola bilər") in nov


def test_oyrenme_ve_madde():
    s = S("oyren.json")
    plan = B.qayda(s, "8-ci sətir sığorta haqqıdır")
    assert plan == ("bank_teyin", {"setir": 8, "kateqoriya": "sigorta_haqqi"})
    n = B.bank_teyin(s, **plan[1])
    assert "daha 1" in n["metn"]                                # eyni VÖEN-li 17-ci sətir də keçdi
    assert _kat(s, 17) == "sigorta_haqqi"
    assert _kat(S("oyren.json"), 8) == "sigorta_haqqi"         # yeni çıxarışda da tanınır
    plan = B.qayda(s, 'Subroqasiya maddəsi "Subroqasiya gəlirləri" olsun')
    assert plan == ("bank_madde", {"kateqoriya": "subroqasiya", "madde": "Subroqasiya gəlirləri"})
    B.bank_madde(s, **plan[1])
    assert B._madde("subroqasiya") == "Subroqasiya gəlirləri"


def test_1c_fayli():
    import openpyxl
    s = S("hesabat.json")
    _madde_deyilib()
    n = _hazirla(s)
    wb = openpyxl.load_workbook(QOV / n["ui"]["fayllar"][1]["ad"])   # yoxlama faylı
    ws = wb["1C - Sığorta haqları"]
    assert [c.value for c in ws[1]][:len(B.SUTUNLAR_1C)] == B.SUTUNLAR_1C
    assert ws.max_row - 1 == 2                                  # yalnız HS-/BS- müqavilə nömrəsi olanlar (5 və 7-ci sətir)
    assert wb["Sığorta haqları (müqaviləsiz)"].max_row - 1 == 2
    assert any(c.value == 16 for c in wb["Yoxlanmalı"]["A"])
    assert "Yoxlanmalı" in wb.sheetnames and "Xəbərdarlıqlar" in wb.sheetnames


def test_lokal_sorgular():
    s = S()
    assert B.qayda(s, "Yoxlanmalıları göstər") == ("bank_siyahi", {"kateqoriya": "yoxlanmali"})
    assert B.qayda(s, "Subroqasiya daxilolmaları") == ("bank_siyahi", {"kateqoriya": "subroqasiya"})
    assert B.qayda(s, "1C üçün Excel hazırla") == ("bank_hesabat", {})
    assert B.qayda(s, "Xülasə") == ("bank_xulase", {})
    assert EX is not None


def test_hazir_sorgular():
    from assistant import teklif as TK
    s = S("teklif.json")
    s.excel, s.m, s.son_niyyet = {"ad": "x", "vereqler": {}}, None, None
    S.hazir = property(lambda self: self.m is not None)
    t = TK.hazirla(s)
    assert t[0] == "Sualları ver (4)" and "1C üçün Excel hazırla" in t      # 3 tanınmayan + 1 «Orta» (subroqasiya)
    assert all(B.qayda(s, x) for x in t)                       # hər təklif başa düşülür


KAPITAL = pathlib.Path(__file__).parent / "data" / "bank_kapital_test.xlsx"


def test_kapital_formati():
    B.QAYDA_FAYLI, B.HESABAT_QOVLUGU = QOV / "kapital.json", QOV
    b = B.oxu(KAPITAL)
    assert b and b["meta"]["hesab"] == "AZ00TEST0000000000000000000"
    ss = {r["setir"]: r for r in b["setirler"]}
    assert len(ss) == 9                                            # mənfi (məxaric) sətir atılır
    k = {i: r["kateqoriya"] for i, r in ss.items()}
    assert k[8] == "telim" and k[10] == "telim"                    # «AQRAR SIĞORTA FONDU» təyinatda sığorta sayılmır
    assert k[9] == "sigorta_haqqi" and k[11] == "sigorta_haqqi"
    assert k[12] == "ezamiyye"                                     # xəzinədarlıq hesabı (CTRE)
    assert k[14] == "qaytarilma" and k[16] == "faiz" and k[17] == "depozit" and k[15] == "yoxlanmali"
    assert ss[9]["odeyici"] == "NUMUNOVA LALE KAMIL QIZI" and ss[9]["fin"] == "1AB2CD3" and ss[9]["teyinat"] == "SIĞORTA"
    assert ss[8]["odeyici"] == "TESTOV AYDIN VAQIF OGLU" and ss[8]["voen"] == ""   # ödəyici bank deyil, şəxsdir
    assert ss[11]["odeyici"] == "Testov Elnur" and ss[11]["muqavile"] == "HS-2026000001-QS"
    assert ss[11]["t_fin_voen"] == "1234567890" and ss[11]["odeyici_hesab"] == "AZ00RBTA00000000000000000099"


def test_1c_inteqrasiya_sablonu():
    import openpyxl
    B.QAYDA_FAYLI, B.HESABAT_QOVLUGU = QOV / "kapital2.json", QOV

    class K:
        bank = B.oxu(KAPITAL)
    _madde_deyilib()
    n = _hazirla(K())
    f1c = n["ui"]["fayllar"][0]["ad"]
    wb = openpyxl.load_workbook(QOV / f1c)
    assert wb.sheetnames == ["Sheet1"]
    rows = list(wb["Sheet1"].iter_rows(values_only=True))
    assert list(rows[0]) == ["TARİX", "Sığortalı Fin/Voen", "Sığorta müqaviləsinin nömrəsi", "HESAB NÖMRƏSİ", "MƏBLƏĞ", "Pul vəsaitinin maddəsi"]
    assert len(rows) == 2                                       # yalnız HS- nömrəli sətir
    assert rows[1][1:5] == ("1234567890", "HS-2026000001-QS", "AZ00TEST0000000000000000000", 2700)   # Fondun hesabı
    assert rows[1][5] == "Sığorta haqları"


def test_muqavilesizler_istege_gore():
    import openpyxl
    B.QAYDA_FAYLI, B.HESABAT_QOVLUGU = QOV / "kapital3.json", QOV

    class K:
        bank = B.oxu(KAPITAL)
    plan = B.qayda(K(), "Müqaviləsizləri də şablona sal")
    assert plan == ("bank_hesabat", {"muqavilesiz": True})
    _madde_deyilib()
    n = _hazirla(K(), **plan[1])
    rows = list(openpyxl.load_workbook(QOV / n["ui"]["fayllar"][0]["ad"])["Sheet1"].iter_rows(values_only=True))
    assert len(rows) == 3                                       # HS- nömrəli + müqaviləsiz sığorta haqqı
    assert {r[2] for r in rows[1:]} == {"HS-2026000001-QS", None}
    assert B.qayda(K(), "Müqaviləsizləri salma")[1] == {"muqavilesiz": False}


def test_muhasibe_sual_ve_cavab():
    s = S("sual.json")
    plan = B.qayda(s, "Sualları ver")
    assert plan == ("bank_sual", {})
    n = B.bank_sual(s)
    assert n["metn"].startswith("Sual 1/4: 8-ci sətir") and n["ui"]["tip"] == "sual"
    assert any(v["mesaj"] == "8-ci sətir — Sığorta haqları olsun" for v in n["ui"]["variantlar"])
    # düymə ilə cavab: təyin edir və növbəti sualı verir
    n = B.bank_teyin(s, **B.qayda(s, "8-ci sətir — Sığorta haqları olsun")[1])
    assert "Oldu" in n["metn"] and "Sual" in n["metn"] and n["ui"]["setir"] != 8
    # sərbəst cavab aktiv suala aiddir
    aktiv = s.bank["aktiv_sual"]
    assert B.qayda(s, "bilmirəm, sonra baxaram") == ("bank_kec", {"setir": aktiv})
    n = B.bank_kec(s, aktiv)
    aktiv = s.bank["aktiv_sual"]
    assert B.qayda(s, "bu digər daxilolmadır") == ("bank_teyin", {"setir": aktiv, "kateqoriya": "diger"})
    # «Orta» təxmin üçün «bəli» təxmini təsdiqləyir
    while s.bank["aktiv_sual"] is not None and next(x for x in s.bank["setirler"] if x["setir"] == s.bank["aktiv_sual"])["kateqoriya"] == "yoxlanmali":
        B.bank_teyin(s, s.bank["aktiv_sual"], "diger")
    if s.bank["aktiv_sual"] is not None:
        r = next(x for x in s.bank["setirler"] if x["setir"] == s.bank["aktiv_sual"])
        assert B.qayda(s, "Bəli") == ("bank_teyin", {"setir": r["setir"], "kateqoriya": r["kateqoriya"]})


def test_mushasibin_sablonunu_doldurur():
    import datetime as dt
    import openpyxl
    B.QAYDA_FAYLI, B.HESABAT_QOVLUGU = QOV / "sablon.json", QOV
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(B.SUTUNLAR_1C)
    ws.append([dt.datetime(2026, 9, 30), "1111111111", "HS-2026000000-QS", "AZ00", 10.0, "Sığorta haqları"])
    yol = QOV / "Sığorta haqqı inteqrasiya faylı.xlsx"
    wb.save(yol)
    assert B.sablon_mi(yol)[2]["MƏBLƏĞ"] == 5
    assert B.sablon_mi(KAPITAL) is None and B.sablon_mi(FAYL) is None

    class K:
        bank = B.oxu(KAPITAL)
        sablon = {"yol": yol, "ad": yol.name}
    _madde_deyilib()
    n = _hazirla(K())
    assert "«Sığorta haqqı inteqrasiya faylı.xlsx» şablonunuz dolduruldu: 1 yeni sətir" in n["metn"]
    rows = list(openpyxl.load_workbook(QOV / n["ui"]["fayllar"][0]["ad"])["Sheet1"].iter_rows(values_only=True))
    assert len(rows) == 3 and rows[1][2] == "HS-2026000000-QS" and rows[2][2] == "HS-2026000001-QS"   # köhnə sətir qalır


def test_madde_muhasibden_sorusulur():
    import openpyxl
    B.QAYDA_FAYLI, B.HESABAT_QOVLUGU = QOV / "madde.json", QOV

    class K:
        bank = B.oxu(KAPITAL)
    k = K()
    n = B.bank_hesabat(k)
    assert "hansı sətirlər düşsün" in n["metn"] and n["ui"]["tip"] == "sual"       # əvvəl: inteqrasiyaya hansı sətirlər
    assert [v["mesaj"] for v in n["ui"]["variantlar"]][:3] == [
        "İnteqrasiyaya yalnız sığorta haqları", "İnteqrasiyaya bütün daxilolmalar", "İnteqrasiyaya bütün sətirlər"]
    plan = B.qayda(k, "İnteqrasiyaya yalnız sığorta haqları")
    assert plan == ("bank_hesabat", {"ehate": "sigorta", "muqavilesiz": False})
    n = B.bank_hesabat(k, ehate="sigorta", muqavilesiz=True)
    assert "Pul vəsaitinin maddəsi" in n["metn"] and "bütün 2 sətrinə" in n["metn"]
    plan = B.qayda(k, "Sığorta haqqı daxilolmaları")                               # mühasib öz adını əl ilə yazır
    assert plan == ("bank_madde", {"kateqoriya": "sigorta_haqqi", "madde": "Sığorta haqqı daxilolmaları"})
    n = B.bank_madde(k, **plan[1])
    assert "fayllar" in n["ui"]                                                     # cavabdan sonra fayl dərhal hazırlanır
    rows = list(openpyxl.load_workbook(QOV / n["ui"]["fayllar"][0]["ad"])["Sheet1"].iter_rows(values_only=True))
    assert len(rows) == 3 and all(r[5] == "Sığorta haqqı daxilolmaları" for r in rows[1:])   # bütün sətirlərə
    assert rows[1][3] == "AZ00TEST0000000000000000000"
    n = B.bank_hesabat(k)                                                           # növbəti doldurmada yenə soruşur
    assert "hansı sətirlər" in n["metn"]
    n = B.bank_hesabat(k, ehate="sigorta")
    assert not any("Sığorta haqqı daxilolmaları" in v["etiket"] for v in n["ui"]["variantlar"])   # əvvəlki seçim təklif edilmir


def test_her_kateqoriyanin_maddesi_sorusulur():
    """«Bütün daxilolmalar»: hər kateqoriyanın maddəsi soruşulur, hər doldurmada yenidən (yaddaşdan götürülmür)."""
    import openpyxl
    B.QAYDA_FAYLI, B.HESABAT_QOVLUGU = QOV / "madde_hamisi.json", QOV

    class K:
        bank = B.oxu(KAPITAL)
    k = K()
    B.bank_hesabat(k)
    n = B.bank_hesabat(k, **B.qayda(k, "İnteqrasiyaya bütün daxilolmalar")[1])
    sorusulan = []
    while n["ui"]["tip"] == "sual":
        kod = k.bank["madde_gozle"]["kod"]
        sorusulan.append(kod)
        cavab = "Faiz maddəsi boş qalsın" if kod == "faiz" else f"«1C {kod}»"
        plan = B.qayda(k, cavab)
        assert plan[1]["kateqoriya"] == kod
        n = T_cagir(k, plan)
    assert sorusulan == ["sigorta_haqqi", "ezamiyye", "telim", "qaytarilma", "depozit", "faiz", "yoxlanmali"]
    rows = list(openpyxl.load_workbook(QOV / n["ui"]["fayllar"][0]["ad"])["Sheet1"].iter_rows(values_only=True))[1:]
    assert len(rows) == 6                                     # 9 daxilolma − faiz (maddəsi verilmədi) − 2 FİN/VÖEN-siz (8, 15)
    assert {r[5] for r in rows} == {"1C sigorta_haqqi", "1C ezamiyye", "1C telim", "1C qaytarilma", "1C depozit"}
    assert all(r[1] for r in rows)                                                # hamısında Sığortalı Fin/Voen var
    assert "Depozit faizləri" in n["metn"] and "salınmadı" in n["metn"]
    B.bank_hesabat(k)                                                               # növbəti doldurma: yenə hər şey soruşulur
    n = B.bank_hesabat(k, ehate="daxilolma")
    assert k.bank["madde_gozle"]["kod"] == "sigorta_haqqi"
    assert not any("1C" in v["etiket"] for v in n["ui"]["variantlar"])           # əvvəlki ad təklif kimi göstərilmir


def test_mexaric_de_inteqrasiyaya():
    import openpyxl
    B.QAYDA_FAYLI, B.HESABAT_QOVLUGU = QOV / "mexaric.json", QOV

    class K:
        bank = B.oxu(KAPITAL)
    k = K()
    assert len(k.bank["mexaricler"]) == 1 and k.bank["mexaricler"][0]["mebleg"] < 0   # məxaric ayrıca oxunur
    n = _hazirla(k, ehate="hamisi")
    rows = list(openpyxl.load_workbook(QOV / n["ui"]["fayllar"][0]["ad"])["Sheet1"].iter_rows(values_only=True))[1:]
    assert len(rows) == 7 and all(r[4] > 0 for r in rows)   # 7 FİN-li daxilolma; məxaricin FİN/VÖEN-i yoxdur — düşmür
    wb = openpyxl.load_workbook(QOV / n["ui"]["fayllar"][1]["ad"])
    finsiz = list(wb["FİN-VÖEN yoxdur"].iter_rows(min_row=2, values_only=True))
    assert [r[0] for r in finsiz] == [8, 15, 13] and finsiz[-1][3] < 0               # yoxlama faylında ayrıca (məxaric də)
    assert "Sığortalı Fin/Voen olmayan" in n["metn"]


def T_cagir(s, plan):
    return getattr(B, plan[0])(s, **plan[1])


def test_hamisina_eyni_madde():
    """Mühasibin ssenarisi: sığortaya «F-011», sonra «Hamısına F-011» — qalan kateqoriyalar bir daha soruşulmur."""
    import openpyxl
    B.QAYDA_FAYLI, B.HESABAT_QOVLUGU = QOV / "hamisina.json", QOV

    class K:
        bank = B.oxu(KAPITAL)
    k = K()
    B.bank_hesabat(k)
    n = B.bank_hesabat(k, ehate="daxilolma")
    assert k.bank["madde_gozle"]["kod"] == "sigorta_haqqi"
    n = B.bank_madde(k, **B.qayda(k, "F-011")[1])                                   # dırnaqsız ad
    assert n["ui"]["tip"] == "sual" and k.bank["madde_gozle"]["kod"] != "sigorta_haqqi"
    assert any(v["etiket"] == "Bütün qalanlara «F-011»" and v["mesaj"] == "Hamısına F-011" for v in n["ui"]["variantlar"])
    plan = B.qayda(k, "Hamısına F-011")
    assert plan == ("bank_madde_hamisi", {"madde": "F-011"})
    n = B.bank_madde_hamisi(k, **plan[1])
    assert "fayllar" in n["ui"]                                                     # bir daha soruşmadı: fayl hazırdır
    wb = openpyxl.load_workbook(QOV / n["ui"]["fayllar"][1]["ad"])
    xu = {r[0]: r[4] for r in wb["Xülasə"].iter_rows(min_row=6, values_only=True) if r[0] and r[1]}
    assert all(v == "F-011" for a, v in xu.items() if a not in ("Cəmi",)), xu
    assert B._madde("telim") == "F-011"                                             # yadda qalıb, «Hamısına F-011» yox


def test_emr_madde_adi_sayilmir():
    B.QAYDA_FAYLI, B.HESABAT_QOVLUGU = QOV / "emr.json", QOV

    class K:
        bank = B.oxu(KAPITAL)
    k = K()
    B.bank_hesabat(k)
    assert B.qayda(k, "Şablonu doldur")[0] == "bank_hesabat"                         # əmrdir — sual yenidən verilir
    assert B.qayda(k, "1C üçün Excel hazırla")[0] == "bank_hesabat"
    assert "maddənin adına bənzəmir" in B.bank_madde(k, "sigorta_haqqi", "Şablonu doldur")["metn"]
    assert B._madde("sigorta_haqqi") == ""


def test_sehv_saxlanmis_adlar_temizlenir():
    """Əvvəlki versiyada səhvən saxlanmış adlar: «Hamısına F-011» -> «F-011», «Şablonu doldur» atılır."""
    B.QAYDA_FAYLI = QOV / "kohne.json"
    B._qaydalar_yaz({"odeyici": {}, "madde": {"subroqasiya": "Hamısına F-011", "telim": "Şablonu doldur"},
                     "madde_tarixce": {"sigorta_haqqi": ["Şablonu doldur", "F-011"], "subroqasiya": ["Hamısına F-011"]}})
    assert B._madde("subroqasiya") == "F-011" and B._madde("telim") == ""
    assert "Şablonu doldur" not in B._teklifler("sigorta_haqqi")
