"""Bütün şirkətə xas ayarlar buradadır. Kodu dəyişmədən yeni şirkətə uyğunlaşdırmaq üçün yalnız bu faylı redaktə edin."""

import os

SIRKET_ADI = os.environ.get("WA_SIRKET", "Test MMC")
VALYUTA = "AZN"
EDV_DERECESI = 0.18
DOZUM = 0.10                 # 10 qəpiyə qədər fərq uyğun sayılır (mühasibin cavabı)
KOMBINASIYA_MAKS = 4         # bir ödənişin örtə biləcəyi maksimum qaimə sayı (kombinasiya axtarışında)
BORC_GECIKME_GUN = 30        # bu qədər gündən köhnə açıq borc "gecikmiş" sayılır
MUQAVILE_XEBERDARLIQ_GUN = 45  # müqavilənin bitməsinə bu qədər gün qalırsa, bildiriş
CAVAB_MUDDETI_IS_GUNU = 5    # borc məktubunda qarşı tərəfə verilən cavab müddəti

# Bank istiqaməti -> qaimə növü
ISTIQAMET = {"Mədaxil": "Satış", "Məxaric": "Alış"}

# Qaimə nömrəsini təyinat mətnində tapmaq üçün nümunə (məs. AA1005, AA 1005, AA-1005)
QAIME_REGEX = r"\b([A-Z]{2})\s?-?(\d{3,})\b"

# Mənbə faylların sütun xəritəsi: {standart ad: fayldakı ad}
# Real e-taxes və bank eksportları gələndə yalnız bu hissə dəyişir.
MENBELER = {
    "satis": {
        "vereq": "E-qaimə_Satış",
        "sutunlar": {"seriya": "Seriya", "nomre": "Nömrə", "tarix": "Tarix", "voen": "Alıcı VÖEN",
                     "ad": "Alıcı adı", "netto": "Məbləğ ƏDV-siz (AZN)", "edv": "ƏDV (AZN)", "status": "Status"},
    },
    "alis": {
        "vereq": "E-qaimə_Alış",
        "sutunlar": {"seriya": "Seriya", "nomre": "Nömrə", "tarix": "Tarix", "voen": "Satıcı VÖEN",
                     "ad": "Satıcı adı", "netto": "Məbləğ ƏDV-siz (AZN)", "edv": "ƏDV (AZN)", "status": "Status"},
    },
    "bank": {
        "vereq": "Bank_çıxarışı",
        "sutunlar": {"tarix": "Tarix", "istiqamet": "İstiqamət", "voen": "Kontragent VÖEN",
                     "ad": "Kontragent adı", "mebleg": "Məbləğ (AZN)", "metn": "Təyinat"},
    },
    "depozit": {
        "vereq": "ƏDV_depozit",
        "sutunlar": {"tarix": "Tarix", "istiqamet": "İstiqamət", "voen": "Kontragent VÖEN",
                     "ad": "Kontragent adı", "mebleg": "ƏDV məbləği (AZN)", "metn": "Qaimə"},
    },
    "muqavile": {
        "vereq": "Müqavilələr",
        "sutunlar": {"nomre": "Müqavilə №", "voen": "Kontragent VÖEN", "ad": "Kontragent adı",
                     "baslama": "Başlama", "bitme": "Bitmə", "mebleg": "Məbləğ ƏDV-siz (AZN)",
                     "evvelki": "Yanvar–Avqust qaimələri (AZN)", "nov": "Növ"},
    },
    "kontragent": {
        "vereq": "Kontragentlər",
        "sutunlar": {"voen": "VÖEN", "ad": "Ad", "email": "E-poçt"},
    },
}

LEGV_STATUSLARI = {"Ləğv edilib"}


def az(x):
    """1234.5 -> '1 234,50' (Azərbaycan formatı)"""
    return f"{x:,.2f}".replace(",", " ").replace(".", ",")
