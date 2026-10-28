# Worker's Assistant

İşçilərin əl ilə gördüyü təkrarlanan işləri avtomatlaşdıran AI köməkçi. İşçi ChatGPT tərzində söhbət edir, Worker's Assistant isə uyğun modulun alətlərini çağırıb işi görür.

Sistem yalnız **təklif edir**, yekun qərarı işçi verir.

## Modullar

| Modul | Status | Nə edir |
|---|---|---|
| Bank daxilolmaları | Aktiv | Bank çıxarışındakı daxilolmaları təsnif edir (sığorta haqqı, subroqasiya, əmək haqqı, ezamiyyə, təlim), maddəni təklif edir, tarix və təkrar xətalarını göstərir, 1C üçün Excel hazırlayır |
| Mühasibat | Aktiv | Bank, ƏDV depozit və e-qaimə uzlaşması, açıq borclar, borc məktubları, Excel hesabatı |
| Excel analizi | Aktiv | İstənilən Excel faylı: struktur, sütun statistikası, qruplaşdırma (cəm/orta/say), sıralama, axtarış, problemlərin tapılması |
| Müqavilələr | Planlaşdırılır | Reyestr, müddət bildirişləri, müqavilə müqayisəsi |
| Sənəd dövriyyəsi | Planlaşdırılır | Məktubların qeydiyyatı, cavab qaralamaları, axtarış |
| Kadrlar | Planlaşdırılır | Məzuniyyət, əmrlər, tabel və arayışlar |

## Quraşdırma (Windows)

```powershell
cd D:\GitHub\worker-assistant
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## İşə salmaq

```powershell
python -m uzlasma tests\data\test_meyar_data.xlsx --dovr-sonu 2026-09-30 --cixis uzlasma_hesabati.xlsx
pytest
```

Test məlumatını yenidən yaratmaq: `python tools\gen_test_data.py`

## Bank daxilolmaları

Kapital Bank çıxarışını (Excel) 📎 ilə yükləyin: sistem başlıq sətrini (Tarix, Kredit/Mədaxil, Təyinat, Qarşı tərəf, VÖEN) özü tapır.

- Hər daxilolma üçün kateqoriya, pul vəsaitinin maddəsi, etibar dərəcəsi və səbəb təklif olunur; tanınmayanlar «Yoxlanmalı»ya düşür.
- Düzəliş: «12-ci sətir subroqasiyadır» — ödəyici (VÖEN və ya ad) `data\bank_qaydalari.json`-da yadda saxlanır, növbəti çıxarışda avtomatik tanınır.
- Maddə adı 1C-dəki kimi olsun: «Subroqasiya maddəsi "..." olsun».
- «1C üçün Excel hazırla»: «1C - Sığorta haqları» yükləmə vərəqi, digər kateqoriyalar, Yoxlanmalı, Xəbərdarlıqlar (başqa ay, tarix oxunmur, təkrar ola bilər). Təkrar ola biləcək sətirlər 1C vərəqinə düşmür.
- Kapital Bankın «Hesabdan çıxarış» formatı dəstəklənir: «Əks hesabın nömrəsi, adı və vöen-i» bir xanada, işarəli «Məbləğ» (mənfi = məxaric). Ödəyici, FİN və qısa təyinat təyinat mətnindən çıxarılır (SV/Fin/ASA, AniPay, «AD - telim odenisi»).
- 1C üçün iki fayl hazırlanır: «Sığorta haqqı inteqrasiya faylı» (mühasibin şablonu: Sheet1 — TARİX, Sığortalı Fin/Voen, Sığorta müqaviləsinin nömrəsi, HESAB NÖMRƏSİ, MƏBLƏĞ, Pul vəsaitinin maddəsi) və yoxlama faylı.
- İnteqrasiya faylına yalnız təyinatında sığorta müqaviləsinin nömrəsi (HS- və ya BS- ilə başlayan) olan sətirlər düşür; Fin/VÖEN və HESAB NÖMRƏSİ (ödəyicinin hesabı) təyinatdan götürülür. Maddə: «Sığorta haqları». Müqaviləsiz sığorta haqları yoxlama faylında ayrıca vərəqdədir.

Test çıxarışları (uydurma): `python tools\gen_bank_test.py` — `bank_cixarisi_test.xlsx` və Kapital formatında `bank_kapital_test.xlsx`. Real çıxarışları repoya əlavə etməyin.

## Worker's Assistant interfeysi

```powershell
python -m assistant
```

Kod dəyişəndə serveri yenidən başlatmaq lazımdır (Ctrl+C, sonra yenidən `python -m assistant`). İnkişaf zamanı avtomatik yenilənmə üçün: `$env:WA_RELOAD="1"; python -m assistant`.

Brauzerdə `http://127.0.0.1:8000` açın. İnterfeys ChatGPT tərzindədir: 📎 ilə Excel faylını əlavə edin, dövrün sonunu seçin və sual verin:
«Vəziyyət necədir?», «Yüksək riskli istisnalar», «Alfa borcu», «AA1005 qaiməsi», «Alfa üçün borc məktubu», «Excel hesabatı».

İstənilən Excel faylı yüklənə bilər. Fayl uzlaşma şablonuna uyğundursa (qaimə, bank, ƏDV depozit vərəqləri), Mühasibat modulu da işə düşür; əks halda fayl Excel analizi modulu ilə təhlil olunur: başlıq sətri avtomatik tapılır, "1 234,56" kimi yazılmış rəqəmlər tanınır. Nümunələr: «Məbləğ cəmi müştəri üzrə», «Ən böyük 10 məbləğ», «'Gəncə' axtar», «Problemləri tap». Test faylı: `tests\data\ixtiyari_excel.xlsx`.

**Böyük fayllar** (məs. 1C bazasının yüzlərlə vərəqli, milyonlarla sətirli ixracı) bir dəfə lokal SQLite bazasına çevrilir: 300+ vərəq və 1,4 mln sətirlik 70 MB fayl təxminən 2–3 dəqiqə çəkir, interfeysdə "Oxunur: 120/323 vərəq" kimi irəliləyiş görünür. Nəticə `data\excel_kes\` qovluğunda saxlanılır; eyni fayl yenidən yüklənəndə bir neçə saniyədə açılır, suallar 1 saniyədən tez cavablanır. Keşi silmək üçün bu qovluğu silin.

AI rəqəmləri özü hesablamır: hər sual mühərrikin alətinə çevrilir (`assistant/tools.py`), AI yalnız nəticəni izah edir.
Bağlantıları mühasib cədvəldəki **✓ / ✗** düymələri ilə təsdiqləyir; qərarlar `data/tesdiqler.json`-da saxlanır və Excel hesabatına düşür.

### Rejimlər (`WA_LLM` mühit dəyişəni)

| Rejim | Necə | Məlumat harada qalır |
|---|---|---|
| `qayda` (default) | AI-sız, açar sözlərlə işləyir | Yalnız bu kompüterdə |
| `openai` | Lokal model, məs. Ollama: `$env:WA_OPENAI_BASE="http://localhost:11434/v1"; $env:WA_MODEL="qwen2.5:14b"` | Yalnız bu kompüterdə |
| `anthropic` | Claude API: `$env:ANTHROPIC_API_KEY="..."; $env:WA_MODEL="..."`, `pip install anthropic` | Alət nəticələri (adlar, məbləğlər) Anthropic serverinə gedir. **Rəhbərliyin yazılı icazəsi lazımdır.** |

```powershell
$env:WA_LLM="openai"; python -m assistant
```

Server yalnız `127.0.0.1`-də işləyir. Şəbəkədə başqalarına açmazdan əvvəl giriş (login) və rollar əlavə olunmalıdır.

`assistant/` strukturu: `app.py` (FastAPI), `modullar/` (iş sahələri və onların alətləri), `tools.py` (bütün modulların alətlərinə vahid giriş), `llm.py` (rejimlər və sistem təlimatı), `state.py` (sessiya, təsdiqlər), `static/index.html` (interfeys). Söhbət tarixçəsi brauzerin yaddaşında (localStorage) saxlanır.

## Struktur

```
worker-assistant/
├── uzlasma/              # Mühasibat modulunun uzlaşma mühərriki
│   ├── config.py         # şirkətə xas ayarlar: dözüm, ƏDV, sütun adları
│   ├── loaders.py        # mənbələri oxuyur, dublikatları işarələyir
│   ├── matcher.py        # ödənişləri qaimələrə bağlayır
│   ├── checks.py         # ƏDV, açıq qaimə, dövr və müqavilə yoxlamaları
│   ├── report.py         # Excel hesabatı
│   └── cli.py            # komanda sətri (python -m uzlasma)
├── assistant/            # Worker's Assistant: söhbət interfeysi (python -m assistant)
│   └── modullar/         # hər iş sahəsi ayrıca modul (muhasibat.py, ...)
├── tests/
│   ├── data/             # SAXTA test məlumatı və cavab açarı
│   └── test_cavab_acari.py
└── tools/
    └── gen_test_data.py  # test məlumatı generatoru
```

Yeni şirkət və ya yeni fayl formatı üçün əsasən yalnız `uzlasma/config.py` dəyişir.

## Uyğunlaşdırma qaydaları

Bank (ƏDV-siz məbləğ) və ƏDV depozit ayrıca uyğunlaşdırılır. Keçidlər ən etibarlıdan başlayır:

1. **Təyinatda qaimə nömrəsi** (məs. `AA1005`): etibar yüksək. Qismən və artıq ödəniş də burada tutulur.
2. **VÖEN + yeganə eyni məbləğli açıq qaimə:** etibar orta.
3. **VÖEN + bir neçə qaimənin cəmi = ödəniş** (ən çox 4 qaimə): etibar orta.
4. **Qalanlar:** istisna. Eyni məbləğli bir neçə qaimə varsa, sistem seçmir, mühasibə göstərir.

Dözüm 3 qəpikdir (`config.DOZUM`).

## Hesabat vərəqləri

Xülasə · İstisnalar (riskə görə rənglənib) · Bağlantılar (Təsdiq: Bəli/Xeyr) · Borc siyahısı · Rəhbərlik · Məktub qaralamaları

## Məlumat təhlükəsizliyi

- Real şirkət məlumatlarını **heç vaxt** repoya əlavə etməyin. `data/` və `hesabatlar/` qovluqları `.gitignore`-dadır.
- Mühərrik xarici AI xidmətinə heç bir məlumat göndərmir, hər şey lokal işləyir.

## Yol xəritəsi

- [ ] Real bank çıxarışı oxuyucuları (hər bank üçün ayrıca modul)
- [ ] E-qaimə PDF oxuyucusu
- [ ] Mühasibin "Təsdiq" sütununu növbəti işə salmada oxumaq
- [ ] Avansların sonrakı qaimələrlə bağlanması
- [ ] Valyuta əməliyyatları və məzənnə fərqi
- [ ] Daxili mühasibat proqramı üçün import faylı
- [ ] E-poçt və Telegram bildirişləri
- [ ] Müqavilə reyestri modulu

## Yeni modul əlavə etmək

1. `assistant/modullar/` qovluğunda fayl yaradın, məsələn `muqavile.py`.
2. Faylda üç şey təyin edin:
   - `MODUL`: ad, təsvir, status və interfeysdəki hazır sorğular;
   - `ALETLER`: `{alət_adı: (funksiya, təsvir, parametrlər)}`. Hər funksiya `{"metn": ..., "ui": ...}` qaytarır;
   - `MECBURI`: mütləq parametrlər.
3. Modulu `assistant/modullar/__init__.py` faylındakı `AKTIV` siyahısına əlavə edin və `PLANLASDIRILIR`-dan silin.

AI yeni alətləri avtomatik görür, sistem təlimatına modulun təsviri öz-özünə əlavə olunur.

Qayda: hesablama və qərar məntiqi modulun kodunda olur, AI yalnız hansı aləti çağıracağını seçir və nəticəni izah edir.

## Serverə yerləşdirmə (TeamCity → Docker, test.server)

Fayllar: `Dockerfile`, `.dockerignore`, `deploy/docker-compose.worker-assistant.yml`, `deploy/teamcity_build.sh`.

1. **Serverdə bir dəfə** — məlumat qovluğu (konteyner `uid 10001` istifadəçisi ilə işləyir):
   ```sh
   sudo mkdir -p /asis/worker-assistant && sudo chown 10001:10001 /asis/worker-assistant
   ```
2. **Stack faylına** `deploy/docker-compose.worker-assistant.yml`-dakı `worker_assistant` servisini əlavə edin (`asis_timer` ilə eyni fayl, eyni `asis_dev_network`, `node.hostname == test.server`). Portu (`published: 4460`) serverdə boş olanla əvəz edin.
3. **TeamCity:** VCS root — bu GitHub repo; build addımı — Command Line: `sh deploy/teamcity_build.sh`; agent test.server-də olmalıdır (obraz registry-yə göndərilmir, lokal qalır). `WA_PAROL` — "Password" tipli parametr, stack deploy zamanı ötürülür. Stack adı fərqlidirsə: `env.WA_STACK`.
4. **İlk dəfə:** `WA_PAROL=... docker stack deploy -c <stack faylı> asis_dev`. Sonrakı build-lər servisi yeni obrazla özü yeniləyir.

Qeydlər:

- Obraz build olunarkən testlər işləyir — test sınsa, obraz yaranmır.
- `WA_HOST=0.0.0.0` olanda `WA_PAROL` mütləqdir, yoxsa server işə düşmür. Brauzer giriş pəncərəsi açır (istifadəçi: `WA_ISTIFADECI`, susmaya görə `muhasib`).
- Sağlamlıq yoxlaması: `GET /saglamliq` (girişsiz, məlumat qaytarmır).
- Real məlumat yalnız `/asis/worker-assistant` volume-dadır, obraza və repoya düşmür. `WA_LLM=qayda` — məlumat serverdən kənara çıxmır.
- Pilot versiyada sessiya bir nəfərlikdir (yaddaşdadır): eyni anda iki nəfər işləsə, eyni faylı görəcək. Ona görə `replicas: 1`.
- Əlaqə HTTP-dir; parol şəbəkədə açıq gedir. Test şəbəkəsindən kənara açılacaqsa, qarşısına HTTPS (reverse proxy) qoyulmalıdır.
