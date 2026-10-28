"""Mühərriki test məlumatı üzərində işlədib cavab açarı ilə tutuşdurur.
İstifadə:  pytest   və ya   python tests/test_cavab_acari.py
"""
import datetime as dt
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
from uzlasma.cli import isle  # noqa: E402

DATA = pathlib.Path(__file__).parent / "data" / "test_meyar_data.xlsx"

# Gözlənilən nəticələr: (növ, istinadda olmalı mətn) və ya ("BAĞLANTI", qaimə, qayda)
GOZLENEN_ISTISNALAR = [
    ("Satış qaiməsi ödənilməyib", "AA1002"),
    ("ƏDV depozitə gəlməyib", "AA1005"),
    ("Qaiməsiz mədaxil", "18.09.2026"),
    ("Dublikat qaimə", "AA1004"),
    ("Müqaviləsiz kontragent", "AA1012"),
    ("Müqaviləsiz kontragent", "BB5008"),
    ("Bitmiş müqavilə üzrə qaimə", "AA1006"),
    ("Müqavilə məbləği aşılıb", "AA1009"),
    ("ƏDV hesablanması səhvdir", "AA1008"),
    ("Alış ƏDV-si depozitdən ödənilməyib", "BB5005"),
    ("Ləğv edilmiş qaimə üzrə ödəniş", "BB5003"),
    ("Dövr fərqi", "AA1011"),
    ("Müddəti bitmək üzrə müqavilə", "M-2026/12"),
    ("Qismən ödəniş", "AA1003"),
    ("Qeyri-müəyyən uyğunluq", "AA1015"),
    ("Ödəniş sətri təkrar", "Bank"),
]
GOZLENEN_BAGLANTILAR = [
    ("AA1001", "Təyinatda qaimə nömrəsi"),       # 2 qəpik fərq, dözüm daxilində
    ("AA1007", "VÖEN + yeganə eyni məbləğ"),
    ("AA1013", "VÖEN + qaimələrin cəmi"),
    ("AA1014", "VÖEN + qaimələrin cəmi"),
]
OLMAMALI = [("Qismən ödəniş", "AA1001"), ("Uyğunlaşdırılmayıb", "AA1007")]


def main(fayl=DATA):
    m, ist, _, _ = isle(fayl, dt.date(2026, 9, 30))
    ok = True
    for nov, istinad in GOZLENEN_ISTISNALAR:
        tap = any(i["nov"] == nov and istinad in i["istinad"] for i in ist)
        print(("✔" if tap else "✘"), "istisna  ", nov, istinad)
        ok &= tap
    for qid, qayda in GOZLENEN_BAGLANTILAR:
        tap = any(b["qaime_id"] == qid and b["qayda"] == qayda and b["axin"] == "bank" for b in m.baglantilar)
        print(("✔" if tap else "✘"), "bağlantı ", qid, qayda)
        ok &= tap
    for nov, istinad in OLMAMALI:
        yanlis = any(i["nov"] == nov and istinad in i["istinad"] for i in ist)
        print(("✘" if yanlis else "✔"), "olmamalı ", nov, istinad)
        ok &= not yanlis
    print("\nNƏTİCƏ:", "HAMISI KEÇDİ" if ok else "XƏTA VAR")
    return 0 if ok else 1


def test_cavab_acari():
    assert main() == 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else DATA))
