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
    # «ust»: alt peşə — üst peşənin bütün imkanları da ona aiddir (diş həkimi həm də həkimdir)
    {"kod": "dis_hekimi", "ust": "hekim", "ad": "Diş həkimi", "tesvir": "qəbullar, WhatsApp xatırlatmaları, stoka tələb", "modullar": ["qebul", "teleb", "excel"],
     "heller": [
         {"s": "Stoka tələb yaz", "b": "Stoka tələb", "a": "nə lazımdır — rəhbər təsdiq edir"},
         {"s": "Təsdiq mesajlarını hazırla", "b": "Qəbula yazılma mesajı", "a": "pasiyentə WhatsApp: tarix, həkim, problem"},
         {"s": "Xatırlatmaları hazırla", "b": "Qəbula yaxın xatırlatma", "a": "pasiyentə və həkimə WhatsApp"},
         {"s": "Stok tələblərini göstər", "b": "Tələblərim", "a": "gözləyir / təsdiqləndi / sifariş edildi"},
         {"s": "Fayl haqqında ümumi məlumat ver", "b": "Qəbul jurnalının strukturu", "a": "vərəqlər, sütunlar, qəbul sayı"},
         {"s": "Məbləğ cəmi xidmət üzrə", "b": "Xidmətlər üzrə məbləğ", "a": "plomb, kanal, implant… hər birindən nə qədər"},
         {"s": "Məbləğ cəmi xəstə üzrə", "b": "Xəstələr üzrə məbləğ", "a": "hansı xəstəyə nə qədər xidmət göstərilib"},
         {"s": "Fayldakı problemləri tap", "b": "Jurnalda səhvlər", "a": "təkrar qəbul, boş xanalar, qeyri-adi məbləğlər"}]},
    {"kod": "klinika_rehber", "ust": "rehber", "ad": "Diş klinikasının rəhbəri", "tesvir": "qəbullar, stok, həkimlər, gəlir", "modullar": ["qebul", "stok", "teleb", "excel"],
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


def effektiv(profil):
    """Seçilmiş peşələr + onların üst peşələri (alt peşə əvvəl: «Diş həkimi», sonra «Həkim»)."""
    out = []
    for k in (profil or {}).get("peseler", []):
        while k and k not in out:
            out.append(k)
            k = next(p for p in PESELER if p["kod"] == k).get("ust")
    return out


def modullar(profil):
    """Profilə uyğun modul kodları (seçim sırası ilə, üst peşələr daxil)."""
    out = []
    for k in effektiv(profil):
        for m in next(p for p in PESELER if p["kod"] == k)["modullar"]:
            if m not in out:
                out.append(m)
    return out


def modul_siyahisi(hamisi, profil):
    """Yan paneldəki modullar: yalnız işləyənlər; peşə seçilibsə — yalnız peşəyə (üst peşələr daxil) aid olanlar,
    peşənin sırası ilə. Peşə seçilməyibsə — bütün işləyən modullar."""
    aktiv = [m for m in hamisi if m.get("status") == "aktiv"]
    kodlar = modullar(profil)
    if not kodlar:
        return aktiv
    umumi = [m for m in aktiv if m.get("umumi")]  # hamıya aid (məs. Tapşırıqlar)
    return umumi + sorted([m for m in aktiv if m["kod"] in kodlar and not m.get("umumi")], key=lambda m: kodlar.index(m["kod"]))


def heller(profil):
    """İstifadəçiyə peşəsinə görə hazır həllər (bir neçə peşə: hər birindən, təkrarsız)."""
    out, gorulen = [], set()
    for k in effektiv(profil):
        p = next(x for x in PESELER if x["kod"] == k)
        for h in p["heller"]:
            if h["s"] not in gorulen:
                gorulen.add(h["s"])
                out.append({**h, "pese": p["ad"]})
    return out


# ======================= peşəyə görə istifadə təlimatı =======================
# Peşə seçiləndə (və «Təlimat» yazılanda) göstərilir. Hər addımın nümunə sorğusu («s») düymədir — basanda işləyir.
# Yalnız sistemin bu gün həqiqətən bacardığı şeylər yazılır (test yoxlayır: hər «s» işləyən sorğudur).
_EXCEL = [
    {"b": "Excel faylını yükləyin", "m": "📎 düyməsi ilə faylı əlavə edin. Sistem vərəqləri, sütunları və sətir sayını göstərir.",
     "s": "Fayl haqqında ümumi məlumat ver"},
    {"b": "Səhvləri tapın", "m": "Təkrar sətirlər, boş xanalar, şübhəli tarixlər və qeyri-adi dəyərlər bir siyahıda.",
     "s": "Fayldakı problemləri tap"},
    {"b": "Hesablayın və axtarın", "m": "Yazın: «<rəqəm sütunu> cəmi <sütun> üzrə», «Ən böyük 10 <sütun>», «'mətn' axtar», «Sütunları göstər»."},
]
TELIMATLAR = {
    "muhasib": {
        "giris": "Sistem bank çıxarışındakı daxilolmaları təsnif edir, 1C üçün inteqrasiya faylı hazırlayır və uzlaşma aparır.",
        "fayl": "Kapital Bank Excel çıxarışı, «Sığorta haqqı inteqrasiya faylı» şablonu, uzlaşma faylı",
        "addimlar": [
            {"b": "Bank çıxarışını yükləyin", "m": "Daxilolmalar kateqoriyalara bölünür: sığorta haqqı, subroqasiya, əmək haqqı, ezamiyyə, təlim…",
             "s": "Bank daxilolmalarının xülasəsi"},
            {"b": "Suallara cavab verin", "m": "Əmin olmadığı daxilolmaları sistem bir-bir soruşur. Cavabınız yadda qalır — eyni ödəyici növbəti "
                                             "dəfə avtomatik tanınır. Tanınmayanlar «Yoxlanmalı»dadır.", "s": "Yoxlanmalıları göstər"},
            {"b": "1C faylını hazırlayın", "m": "Sistem «Pul vəsaitinin maddəsi»ni soruşur — təkliflərdən seçin və ya 1C-dəki adı yazın. "
                                               "Öz şablonunuzu yükləsəniz, onu doldurur.", "s": "1C üçün Excel hazırla"},
            {"b": "Uzlaşma", "m": "Uzlaşma faylını yükləyin: bağlantılar, istisnalar, debitor borcları.", "s": "Vəziyyət necədir?"}]},
    "maliyye": {
        "giris": "Sistem Excel hesabatlarını təhlil edir: cəmlər, qruplaşdırma, ən böyüklər, səhvlər.",
        "fayl": "istənilən Excel hesabatı, uçot jurnalı",
        "addimlar": _EXCEL[:1] + [{"b": "Ümumi xərc", "m": "Uçot yazılışlarından (Dt 7xx) hesablanır; hansı hesabların daxil olduğunu "
                                                         "özünüz dəyişə bilərsiniz.", "s": "Şirkətin ümumi xərci nə qədərdir?"}] + _EXCEL[1:]},
    "huquq": {
        "giris": "Müqavilələr modulu (müddət bildirişləri) hazırlanır. Hələlik müqavilə reyestrini Excel kimi təhlil edə bilərsiniz.",
        "fayl": "müqavilə reyestri (Excel)", "addimlar": _EXCEL},
    "karguzar": {
        "giris": "Sənəd dövriyyəsi modulu hazırlanır. Hələlik məktub qeydiyyatı jurnalını Excel kimi təhlil edə bilərsiniz.",
        "fayl": "qeydiyyat jurnalı (Excel)", "addimlar": _EXCEL},
    "kadr": {
        "giris": "Kadrlar modulu hazırlanır. Hələlik tabeli və kadr cədvəlini Excel kimi təhlil edə bilərsiniz.",
        "fayl": "tabel, kadr cədvəli (Excel)", "addimlar": _EXCEL},
    "rehber": {
        "giris": "Sistem ümumi vəziyyəti, daxilolmaları, borcları və xərcləri qısa xülasə kimi göstərir.",
        "fayl": "uzlaşma faylı, bank çıxarışı, uçot jurnalı",
        "addimlar": [
            {"b": "Ümumi vəziyyət", "m": "Uzlaşma faylı yüklənibsə: bağlanma faizi, istisnalar, borclar.", "s": "Vəziyyət necədir?"},
            {"b": "Daxilolmalar", "m": "Bank çıxarışı yüklənibsə: kateqoriyalar üzrə məbləğlər.", "s": "Bank daxilolmalarının xülasəsi"},
            {"b": "Borclar və xərclər", "m": "Kim bizə nə qədər borcludur; şirkətin ümumi xərci.", "s": "Debitor borcları göstər"}]},
    "hekim": {
        "giris": "Sistem xəstə qeydiyyatını və qəbul cədvəlini Excel kimi təhlil edir. Xəstə məlumatı həssasdır — faylı yalnız "
                 "klinikanın öz serverində işlədin.",
        "fayl": "xəstə qeydiyyatı, qəbul cədvəli (Excel)",
        "addimlar": _EXCEL[:1] + [{"b": "Statistika", "m": "Hər sütun üzrə: ən çox rast gələn dəyərlər (məs. diaqnozlar), tarix aralığı, "
                                                         "min / maks / orta.", "s": "Sütunları göstər"}] + _EXCEL[1:2]},
    "dis_hekimi": {
        "giris": "Sistem stoka tələblərinizi rəhbərə çatdırır və qəbula yazılan xəstələrə WhatsApp mesajları hazırlayır.",
        "fayl": "qəbul jurnalı (Xəstə, Telefon, Qəbul tarixi, Saat, Həkim, Problem, Yazılma tarixi)",
        "addimlar": [
            {"b": "Stoka tələb yazın", "m": "Fayl lazım deyil. Yazın: «Stoka tələb: Kompozit plomb A2 — 5 ədəd, təcili». Rəhbər təsdiq "
                                           "edəndə «Tələblərim»də görünəcək.", "s": "Stoka tələb yaz"},
            {"b": "Qəbul jurnalını yükləyin", "m": "Sistem xəstələri, qəbul vaxtını və telefonu olmayanları göstərir.", "s": "Qəbulları göstər"},
            {"b": "Xəstəyə təsdiq mesajı", "m": "Nə vaxt yazıldığı, qəbul vaxtı, həkim və problem. Xəstədən gecikəcəksə sizi "
                                               "məlumatlandırmasını xahiş edir. «WhatsApp-da aç»ı basıb göndərin.",
             "s": "Təsdiq mesajlarını hazırla"},
            {"b": "Qəbula yaxın xatırlatma", "m": "Həm xəstəyə, həm sizə: işiniz yubanarsa, pasiyenti məlumatlandırın (nömrəsi mesajdadır). "
                                                 "Nömrəniz jurnalda yoxdursa yazın: «Dr. Ad Soyad telefonu 050 …».",
             "s": "Xatırlatmaları hazırla"}]},
    "klinika_rehber": {
        "giris": "Sistem həkimlərin stok tələblərini sizə gətirir, azalan məhsulları göstərir, təchizatçılara sifariş və xəstələrə "
                 "xatırlatma hazırlayır. Heç nə sizin təsdiqiniz olmadan göndərilmir.",
        "fayl": "stok faylı (Məhsul, Qalıq, Minimum qalıq, Təchizatçı, Təchizatçının telefonu), qəbul jurnalı",
        "addimlar": [
            {"b": "Nömrənizi yazın", "m": "Stok bildirişləri və həkimlərin tələbləri WhatsApp ilə sizə gəlsin: «Rəhbərin telefonu 050 123 45 67»."},
            {"b": "Həkimlərin tələbləri", "m": "Təsdiq və ya rədd yalnız sizdədir. Təsdiqlənənlər sifarişə əlavə olunur.",
             "s": "Stok tələblərini göstər"},
            {"b": "Stoku yoxlayın", "m": "Stok faylını yükləyin: bitən və minimuma çatan məhsullar, sifariş təklifi ilə.",
             "s": "Azalan məhsulları göstər"},
            {"b": "Təchizatçılara sifariş", "m": "Hər təchizatçıya öz məhsulları ilə WhatsApp mesajı.", "s": "Təchizatçılara sifariş hazırla"},
            {"b": "Gəlir təhlili", "m": "Qəbul / gəlir cədvəli üzrə həkimlər və xidmətlər.", "s": "Məbləğ cəmi həkim üzrə"}]},
}


TAPSIRIQ_TELIMATI = {
    "ad": "Tapşırıqlar", "giris": "Rəhbərin verdiyi tapşırıqlar sizə gəlir; icranı yazırsınız, rəhbərə sorğu göndərirsiniz.",
    "fayl": "lazım deyil",
    "addimlar": [
        {"b": "Tapşırıqlarınıza baxın", "m": "Yeni tapşırıq gələndə boş ekranda «Bildiriş» kartı görünür.", "s": "Tapşırıqlarım"},
        {"b": "İcranı yazın", "m": "«TP-001 icraya başladım», «TP-001 addım: …», «TP-001 tamamlandı». Rəhbər hər addımı görür."},
        {"b": "Rəhbərə sorğu", "m": "«Təsdiq sorğusu: …» və ya «Məlumat sorğusu: …»; tapşırığa bağlı: «TP-001 üzrə təsdiq sorğusu: …».",
         "s": "Sorğular"}]}
TAPSIRIQ_TELIMATI_REHBER = {
    "ad": "Tapşırıqlar", "giris": "İşçiyə tapşırıq və icra müddəti verir, icranı izləyir, sorğulara cavab verirsiniz.",
    "fayl": "lazım deyil",
    "addimlar": [
        {"b": "Tapşırıq verin", "m": "İşçinin adı və ya e-poçtu ilə: «Aynur Məmmədovaya tapşırıq: Hesabatı hazırla, müddət 15.11» "
                                    "(müddət: tarix, «3 gün», «sabah»)."},
        {"b": "İcranı izləyin", "m": "Verdiyiniz tapşırıqlar — gecikənlər yuxarıda; nömrəsini yazsanız, bütün addımlar görünür.",
         "s": "Tapşırıqlarım"},
        {"b": "Sorğulara cavab verin", "m": "«S-001 təsdiq», «S-001 rədd», «S-001 cavab: …». Müddəti dəyişmək: «TP-001 müddət 20.11».",
         "s": "Sorğular"}]}


def telimat(profil):
    """Seçilmiş peşələr üçün təlimat (interfeysdə «telimat» bloku kimi göstərilir)."""
    kodlar = effektiv(profil)
    if not kodlar:
        return None
    peseler = [{"ad": next(p["ad"] for p in PESELER if p["kod"] == k), **TELIMATLAR[k]} for k in kodlar if k in TELIMATLAR]
    peseler.append(TAPSIRIQ_TELIMATI_REHBER if "rehber" in kodlar else TAPSIRIQ_TELIMATI)  # hamıya aid modul
    return {"tip": "telimat", "peseler": peseler,
            "son": "Bu təlimatı istənilən vaxt yenidən açmaq üçün «Təlimat» yazın. Peşəni yan paneldəki «Peşələrim»dən dəyişə bilərsiniz."}
