"""İstifadəçi profili: işçinin məşğul olduğu peşələr (hər istifadəçinin öz faylında, WA_DATA/profil/).

Girişdən sonra peşələr seçilməyibsə, interfeys soruşur (yalnız siyahıdan seçilir); cavab yadda qalır və istənilən
kompüterdən keçərlidir. Hər peşənin öz hazır həlləri var: boş ekranda istifadəçiyə peşəsinə görə həllər təklif olunur.
Həllər yalnız sistemin bu gün həqiqətən bacardığı sorğulardır (modulu hələ qurulmayan peşələrə Excel təhlili ilə).
"""
import hashlib
import json
import threading
import time

from .state import DATA

QOVLUQ = DATA / "profil"
_KILID = threading.Lock()

# Peşə -> uyğun modullar və hazır həllər. «s» — sistemə gedən sorğu (lokal rejimdə başa düşülən), «b» — başlıq, «a» — izah.
PESELER = [
    {"kod": "muhasib", "ad": "Mühasib", "tesvir": "bank daxilolmaları, uzlaşma, 1C", "modullar": ["bank", "muhasibat", "excel"],
     "heller": [
         {"s": "1C üçün Excel hazırla", "b": "Bank çıxarışından 1C faylı", "a": "sığorta haqları inteqrasiya faylı"},
         {"s": "Yoxlanmalıları göstər", "b": "Tanınmayan daxilolmalar", "a": "kateqoriyanı siz təyin edin"},
         {"s": "Vəziyyət necədir?", "b": "Uzlaşmanın vəziyyəti", "a": "bağlanma faizi, istisnalar, borclar"},
         {"s": "Debitor borcları göstər", "b": "Debitor borclar", "a": "bizə kim nə qədər borcludur"}]},
    {"kod": "maliyye", "ad": "Maliyyə / iqtisadçı", "tesvir": "hesabatlar, Excel təhlili", "modullar": ["excel", "muhasibat"],
     "heller": [
         {"s": "Şirkətin ümumi xərci nə qədərdir?", "b": "Ümumi xərc", "a": "uçot yazılışlarından hesablanır"},
         {"s": "Fayl haqqında ümumi məlumat ver", "b": "Hesabatın strukturu", "a": "vərəqlər, sütunlar, sətir sayı"},
         {"s": "Fayldakı problemləri tap", "b": "Hesabatda səhvlər", "a": "təkrarlar, boş xanalar, qeyri-adi dəyərlər"},
         {"s": "Vəziyyət necədir?", "b": "Uzlaşmanın vəziyyəti", "a": "bağlanma faizi, istisnalar"}]},
    {"kod": "huquq", "ad": "Hüquqşünas", "tesvir": "müqavilələr, müddətlər", "modullar": ["muqavile", "excel"],
     "heller": [
         {"s": "Fayl haqqında ümumi məlumat ver", "b": "Müqavilə reyestrinin strukturu", "a": "Excel reyestrini yükləyin"},
         {"s": "Fayldakı problemləri tap", "b": "Reyestrdə səhvlər", "a": "təkrar nömrələr, boş xanalar"}]},
    {"kod": "karguzar", "ad": "Kargüzar", "tesvir": "məktublar, sənəd dövriyyəsi", "modullar": ["senedler", "excel"],
     "heller": [
         {"s": "Fayl haqqında ümumi məlumat ver", "b": "Qeydiyyat jurnalının strukturu", "a": "Excel jurnalını yükləyin"},
         {"s": "Fayldakı problemləri tap", "b": "Jurnalda səhvlər", "a": "təkrar qeydiyyat, boş xanalar"}]},
    {"kod": "kadr", "ad": "Kadr mütəxəssisi", "tesvir": "məzuniyyət, əmrlər, tabel", "modullar": ["kadr", "excel"],
     "heller": [
         {"s": "Fayl haqqında ümumi məlumat ver", "b": "Tabelin / kadr cədvəlinin strukturu", "a": "Excel faylını yükləyin"},
         {"s": "Fayldakı problemləri tap", "b": "Tabeldə səhvlər", "a": "təkrarlar, boş xanalar, qeyri-adi dəyərlər"}]},
    # Tibb: modul hələ yoxdur — həllər Excel təhlili ilə (xəstə qeydiyyatı, qəbul jurnalı, klinikanın gəlir cədvəli)
    {"kod": "hekim", "ad": "Həkim", "tesvir": "xəstə qeydiyyatı, qəbullar", "modullar": ["excel"],
     "heller": [
         {"s": "Fayl haqqında ümumi məlumat ver", "b": "Xəstə qeydiyyatının strukturu", "a": "vərəqlər, sütunlar, qeyd sayı"},
         {"s": "Sütunları göstər", "b": "Göstəricilər üzrə statistika", "a": "ən çox rast gələn dəyərlər, min / maks / orta"},
         {"s": "Fayldakı problemləri tap", "b": "Qeydiyyatda səhvlər", "a": "təkrar qeydlər, boş xanalar, şübhəli tarixlər"}]},
    {"kod": "dis_hekimi", "ad": "Diş həkimi", "tesvir": "qəbullar, WhatsApp xatırlatmaları, stoka tələb", "modullar": ["qebul", "teleb", "excel"],
     "heller": [
         {"s": "Stoka tələb yaz", "b": "Stoka tələb", "a": "nə lazımdır — rəhbər təsdiq edir"},
         {"s": "Təsdiq mesajlarını hazırla", "b": "Qəbula yazılma mesajı", "a": "pasiyentə WhatsApp: tarix, həkim, problem"},
         {"s": "Xatırlatmaları hazırla", "b": "Qəbula yaxın xatırlatma", "a": "pasiyentə və həkimə WhatsApp"},
         {"s": "Stok tələblərini göstər", "b": "Tələblərim", "a": "gözləyir / təsdiqləndi / sifariş edildi"},
         {"s": "Fayl haqqında ümumi məlumat ver", "b": "Qəbul jurnalının strukturu", "a": "vərəqlər, sütunlar, qəbul sayı"},
         {"s": "Məbləğ cəmi xidmət üzrə", "b": "Xidmətlər üzrə məbləğ", "a": "plomb, kanal, implant… hər birindən nə qədər"},
         {"s": "Məbləğ cəmi xəstə üzrə", "b": "Xəstələr üzrə məbləğ", "a": "hansı xəstəyə nə qədər xidmət göstərilib"},
         {"s": "Fayldakı problemləri tap", "b": "Jurnalda səhvlər", "a": "təkrar qəbul, boş xanalar, qeyri-adi məbləğlər"}]},
    {"kod": "klinika_rehber", "ad": "Diş klinikasının rəhbəri", "tesvir": "qəbullar, stok, həkimlər, gəlir", "modullar": ["qebul", "stok", "teleb", "excel"],
     "heller": [
         {"s": "Stok tələblərini göstər", "b": "Həkimlərin stok tələbləri", "a": "təsdiq və ya rədd edin"},
         {"s": "Qəbulları göstər", "b": "Qəbul jurnalı", "a": "kim, nə vaxt, hansı həkimdə; telefonu olmayanlar"},
         {"s": "Azalan məhsulları göstər", "b": "Stokda azalan məhsullar", "a": "nə bitib, nə sifariş edilməlidir"},
         {"s": "Təchizatçılara sifariş hazırla", "b": "Təchizatçıya sifariş", "a": "WhatsApp: hər təchizatçıya öz məhsulları"},
         {"s": "Xatırlatmaları hazırla", "b": "Qəbula yaxın xatırlatma", "a": "pasiyentə və həkimə WhatsApp"},
         {"s": "Məbləğ cəmi həkim üzrə", "b": "Həkimlər üzrə gəlir", "a": "hər həkimin göstərdiyi xidmətlərin məbləği"},
         {"s": "Ödənilib cəmi həkim üzrə", "b": "Həkimlər üzrə ödənilən", "a": "xəstələrin faktiki ödədiyi məbləğ"},
         {"s": "Məbləğ cəmi xidmət üzrə", "b": "Xidmətlər üzrə gəlir", "a": "ən çox gəlir gətirən xidmətlər"},
         {"s": "Ən böyük 10 məbləğ", "b": "Ən böyük 10 qəbul", "a": "ən yüksək məbləğli qəbullar"},
         {"s": "Fayldakı problemləri tap", "b": "Hesabatda səhvlər", "a": "təkrar sətirlər, boş həkim / xidmət xanaları"}]},
    {"kod": "rehber", "ad": "Rəhbər", "tesvir": "xülasələr, nəzarət", "modullar": ["bank", "muhasibat", "excel"],
     "heller": [
         {"s": "Vəziyyət necədir?", "b": "Ümumi vəziyyət", "a": "bağlanma faizi, istisnalar, borclar"},
         {"s": "Bank daxilolmalarının xülasəsi", "b": "Daxilolmalar", "a": "kateqoriyalar üzrə"},
         {"s": "Debitor borcları göstər", "b": "Debitor borclar", "a": "bizə kim nə qədər borcludur"},
         {"s": "Şirkətin ümumi xərci nə qədərdir?", "b": "Ümumi xərc", "a": "uçot yazılışlarından"}]},
]
KODLAR = {p["kod"] for p in PESELER}


def _fayl(login):
    ad = hashlib.sha256((login or "lokal").encode()).hexdigest()[:20]  # fayl adında e-poçt görünməsin
    return QOVLUQ / f"{ad}.json"


def al(login):
    """Seçilmiş peşələr: {"peseler": [...]}; hələ seçilməyibsə None."""
    with _KILID:
        try:
            d = json.loads(_fayl(login).read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return None
    peseler = [k for k in d.get("peseler", []) if k in KODLAR]
    return {"peseler": peseler} if peseler else None  # siyahıdan heç nə seçilməyibsə — yenidən soruşulur


def saxla(login, peseler):
    """Yalnız siyahıdakı peşələr qəbul olunur (istifadəçi özü peşə yazmır)."""
    peseler = [k for k in dict.fromkeys(peseler or []) if k in KODLAR]
    if not peseler:
        raise ValueError("Siyahıdan ən azı bir peşə seçin.")
    d = {"login": login, "peseler": peseler, "vaxt": time.time()}
    with _KILID:
        QOVLUQ.mkdir(parents=True, exist_ok=True)
        f = _fayl(login)
        tmp = f.with_suffix(".tmp")
        tmp.write_text(json.dumps(d, ensure_ascii=False), encoding="utf-8")
        tmp.replace(f)
    return {"peseler": peseler}


def modullar(profil):
    """Profilə uyğun modul kodları (seçim sırası ilə)."""
    out = []
    for k in (profil or {}).get("peseler", []):
        for m in next(p for p in PESELER if p["kod"] == k)["modullar"]:
            if m not in out:
                out.append(m)
    return out


def heller(profil):
    """İstifadəçiyə peşəsinə görə hazır həllər (bir neçə peşə: hər birindən, təkrarsız)."""
    out, gorulen = [], set()
    for k in (profil or {}).get("peseler", []):
        p = next(x for x in PESELER if x["kod"] == k)
        for h in p["heller"]:
            if h["s"] not in gorulen:
                gorulen.add(h["s"])
                out.append({**h, "pese": p["ad"]})
    return out
