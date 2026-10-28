"""Test üçün uydurma bank çıxarışı (Kapital Bank formatına bənzər). Adlar və nömrələr uydurmadır."""
import datetime as dt
import pathlib

from openpyxl import Workbook

YOL = pathlib.Path(__file__).resolve().parents[1] / "tests" / "data" / "bank_cixarisi_test.xlsx"
D = dt.datetime
SETIRLER = [  # tarix, sənəd, qarşı tərəf, VÖEN, hesab, debet, kredit, təyinat
    (D(2026, 9, 1), "101", "Yaşıl Tarla MMC", "1400000011", "AZ00TEST0001", None, 1250.00, "Müqavilə № HS-2026000123-QS üzrə sığorta haqqı"),
    (D(2026, 9, 2), "102", "Fermer Nümunə Kəndli", "2400000022", "AZ00TEST0002", None, 86.40, "Sığorta haqqı, polis № PL-77881"),
    (D(2026, 9, 3), "103", "Aqro Test ASC", "1500000033", "AZ00TEST0003", None, "3 400,00", "SH ödənişi BS-2026000150"),
    (D(2026, 9, 4), "104", "Bağ Test MMC", "1600000044", "AZ00TEST0004", None, 515.75, "Hesaba mədaxil"),
    (D(2026, 9, 5), "105", "Sürücü Test", "", "AZ00TEST0005", None, 700.00, "Subroqasiya qaydasında zərərin əvəzi, iş 45/2026"),
    (D(2026, 9, 8), "106", "Dövlət Xəzinədarlıq Agentliyi", "1700000055", "AZ00TEST0006", None, 18500.00, "Sentyabr ayı üzrə əmək haqqı"),
    (D(2026, 9, 8), "107", "Dövlət Xəzinədarlıq Agentliyi", "1700000055", "AZ00TEST0006", None, 640.00, "Ezamiyyə xərcləri"),
    (D(2026, 9, 9), "108", "Dövlət Xəzinədarlıq Agentliyi", "1700000055", "AZ00TEST0006", None, 2000.00, "Maliyyələşmə"),
    (D(2026, 9, 10), "109", "Namizəd Test", "", "AZ00TEST0007", None, 120.00, "Ekspert namizədi üçün təlim haqqı"),
    (D(2026, 9, 11), "110", "Ofis Təchizat MMC", "1800000066", "AZ00TEST0008", 300.00, None, "Dəftərxana ləvazimatı"),
    (D(2026, 8, 29), "111", "Yaşıl Tarla MMC", "1400000011", "AZ00TEST0001", None, 99.90, "Sığorta haqqı"),
    (D(2026, 9, 2), "102", "Fermer Nümunə Kəndli", "2400000022", "AZ00TEST0002", None, 86.40, "Sığorta haqqı, polis № PL-77881"),
    (D(2026, 9, 12), "112", "Bağ Test MMC", "1600000044", "AZ00TEST0004", None, 210.00, "Ödəniş"),
]


def yarat(yol=YOL):
    wb = Workbook()
    ws = wb.active
    ws.title = "Çıxarış"
    ws.append(["Kapital Bank ASC — hesabdan çıxarış (TEST)"])
    ws.append(["Hesab: AZ00TEST0000000000000000", None, None, "Dövr: 01.09.2026 – 30.09.2026"])
    ws.append([])
    ws.append(["Tarix", "Sənəd №", "Qarşı tərəfin adı", "VÖEN", "Qarşı tərəfin hesabı", "Debet", "Kredit", "Təyinat"])
    for r in SETIRLER:
        ws.append(list(r))
    ws.append([None, None, None, None, None, 300.00, 27608.45, "Cəmi dövriyyə"])
    yol.parent.mkdir(parents=True, exist_ok=True)
    wb.save(yol)
    return yol




# Kapital Bankın real formatına bənzər çıxarış: «Əks hesabın nömrəsi, adı və vöen-i» bir xanada, işarəli «Məbləğ».
# Bütün adlar, FİN-lər və hesablar UYDURMADIR.
KAPITAL_YOL = YOL.parent / "bank_kapital_test.xlsx"
KAPITAL = [
    (D(2026, 10, 1), "AZ00AIIB0000000000000000001 / KAPITAL BANK ASC / 1400000001", 20.0, "TESTOV AYDIN VAQIF OGLU - telim odenisi,Million100000001"),
    (D(2026, 10, 1), "AZ00AIIB0000000000000000002 / KAPITAL / 1400000002", 911.0,
     "SV:1AB2CD3,Teyinat:   SIĞORTA,Fin:1AB2CD3,ASA:NUMUNOVA LALE KAMIL QIZI,BT:,DT:1970-01-01,K100"),
    (D(2026, 10, 1), "AZ00AIIB0000000000000000003 / KAPITAL BANK ASC / 1400000003", 20.0, "Təhsil haqqı,AQRAR SIĞORTA FONDU,9ZZ8YY7,KassamAz100000002"),
    (D(2026, 10, 2), "AZ00AIIB0000000000000000004 / KAPITAL BANK ASC / 1400000004", 2700.0,
     "TYPE=AniPay;TX_ID=TEST1;BIC=TEST;rrn=TEST;DTR_IBAN=AZ00RBTA00000000000000000099;DTR_NAME=Testov Elnur;TAX=1234567890;"
     "HS-2026000001-QS 15.10.2026-CI HF-YA ESASEN KEND TESERRUFATI HEYVANLARININ SIGORTASI;;"),
    (D(2026, 10, 6), "AZ00CTRE0000000000000000005 / Aqrar Sigorta Fondu / 1400000005", 894.0, "Ezamiyye xerci         100000003"),
    (D(2026, 10, 6), "AZ00AIIB0000000000000000006 / KAPITAL / 1400000006", -297.0, "100000004 TESTOV AYDIN VAQIF OGLU Ezamiyya"),
    (D(2026, 10, 13), "AZ00AIIB0000000000000000007 / KAPITAL / 1400000007", 7060.0,
     "SV:2BC3DE4,Teyinat:  FOND VƏSAITIN QAYTARILMASI,Fin:2BC3DE4,ASA:TESTOVA GULNAR ALI QIZI,BT:,DT:1980-01-01,K100"),
    (D(2026, 10, 23), "AZ00AIIB0000000000000000008 / KAPITAL BANK ASC / 1400000008", 110.0, "TESTOV RAMIZ ELI OGLU - alma ting,Million100000005"),
    (D(2026, 10, 26), "AZ00AIIB0000000000000000009 / AQRAR SIĞORTA FONDU / 1400000009", 9600.0, "256/CDPST/0001 müqavilə üzrə faizlərin ödənilməsi"),
    (D(2026, 10, 26), "AZ00AIIB0000000000000000009 / AQRAR SIĞORTA FONDU / 1400000009", 600000.0, "Müqavilə üzrə depozit məbləğinin ödənilməsi 256/CDPST/0001"),
]


def yarat_kapital(yol=KAPITAL_YOL):
    wb = Workbook()
    ws = wb.active
    ws.title = "Hesabdan Çıxarış"
    ws.append(["Hesabdan çıxarış / Account statement"])
    ws.append([])
    ws.append(["Müştərinin Adı / Client's name", "TEST FONDU", "Çıxarışın tarixi / Statement date", "28/10/2026 12:58"])
    ws.append(["Hesab nömrəsi / Account number", "AZ00TEST0000000000000000000", "Hesabın valyutası / Account currency", "AZN"])
    ws.append(["Hesabın tipi / Account type", "Cari hesab", "Dövr üçün / For period", "01/10/2026 - 28/10/2026"])
    ws.append([])
    ws.append(["№", "Tarix", "Sənədin nömrəsi", "Əks hesabın nömrəsi, adı və vöen-i", "Məbləğ", "Valyuta", "AZN dəyəri", "Balans", "Təyinat"])
    bal = 1000.0
    for i, (tarix, eks, mebleg, teyinat) in enumerate(KAPITAL, 1):
        bal += mebleg
        ws.append([float(i), tarix, f"{i:011d} / 0000000000", eks, mebleg, "AZN", mebleg, round(bal, 2), teyinat])
    wb.save(yol)
    return yol


if __name__ == "__main__":
    print(yarat())
    print(yarat_kapital())
