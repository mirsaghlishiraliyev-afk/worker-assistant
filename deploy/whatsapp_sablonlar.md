# WhatsApp şablonları (diş klinikası)

Biznesin təşəbbüsü ilə göndərilən WhatsApp mesajları (xəstəyə yazılma təsdiqi, xatırlatma) yalnız **Meta-da təsdiqlənmiş şablonlarla** gedə bilər.
Bu şablonları Meta Business Manager → WhatsApp Manager → Message templates bölməsində yaradın:

- Kateqoriya: **Utility** (xidmət bildirişi)
- Dil: **Azerbaijani (az)** — `WA_WHATSAPP_DIL`
- Ad və mətn aşağıdakı kimi, hərfi-hərfinə ( `{{1}}`, `{{2}}`… — parametrlər, sıra ilə)

Mətn assistant/modullar/qebul.py-dakı `SABLONLAR` ilə eyni olmalıdır (test bunu yoxlayır).

## `qebul_tesdiq` — Qəbula yazılma (təsdiq) (pasiyentə)

```
Hörmətli {{1}}! {{2}} tarixində {{3}} həkimin qəbuluna yazıldınız. Qəbul vaxtınız: {{4}}. Müraciət səbəbi: {{5}}. Gecikəcəyiniz və ya gələ bilməyəcəyiniz halda, xahiş edirik, həkimi əvvəlcədən məlumatlandırın: {{6}}.
```

Parametrlər: {{1}} xəstənin adı, {{2}} yazılma tarixi, {{3}} həkim, {{4}} qəbul vaxtı, {{5}} müraciət səbəbi, {{6}} həkimin nömrəsi (yoxdursa WA_KLINIKA_TEL)

## `qebul_xatirlatma` — Qəbul xatırlatması (pasiyentə)

```
Hörmətli {{1}}! Xatırladırıq: {{2}} {{3}} həkimin qəbulundasınız. Müraciət səbəbi: {{4}}. Gecikəcəksinizsə, xahiş edirik, həkimi dərhal məlumatlandırın: {{5}}.
```

Parametrlər: {{1}} xəstənin adı, {{2}} qəbul vaxtı, {{3}} həkim, {{4}} müraciət səbəbi, {{5}} həkimin nömrəsi (yoxdursa WA_KLINIKA_TEL)

## `hekim_xatirlatma` — Həkimə xatırlatma (həkimə)

```
Hörmətli {{1}}! {{2}} {{3}} qəbulunuzdadır. Müraciət səbəbi: {{4}}. İşiniz yubanarsa, xahiş edirik, pasiyenti əvvəlcədən məlumatlandırın: {{5}}.
```

Parametrlər: {{1}} həkim, {{2}} qəbul vaxtı, {{3}} xəstənin adı, {{4}} müraciət səbəbi, {{5}} pasiyentin nömrəsi


## Stok nəzarəti (diş klinikası)

### `stok_rehber` — Rəhbərə stok bildirişi

```
Salam! Klinikanın stokunda {{1}} məhsul minimum həddən azdır: {{2}}. Təchizatçılara sifariş göndərmək üçün Worker's Assistant-da «Təchizatçılara sifariş hazırla» seçin.
```

Parametrlər: {{1}} azalan məhsulların sayı, {{2}} məhsullar siyahısı (nöqtəli vergüllə)

### `stok_sifaris` — Təchizatçıya sifariş

```
Salam, {{1}}! Diş klinikası üçün sifariş vermək istəyirik: {{2}}. Xahiş edirik, mövcudluğu, qiyməti və çatdırılma vaxtını bildirin. Təşəkkür edirik.
```

Parametrlər: {{1}} təchizatçının adı, {{2}} məhsullar siyahısı (nöqtəli vergüllə)

Rəhbərin nömrəsi: `WA_REHBER_TEL` və ya çatda «Rəhbərin telefonu 050 123 45 67». Təchizatçıların nömrələri stok faylının «Təchizatçının telefonu» sütunundan götürülür.

## Qoşulma

1. Meta for Developers-də tətbiq yaradın, WhatsApp məhsulunu əlavə edin, klinikanın nömrəsini qoşun.
2. Daimi token (System User) alın.
3. Serveri bu dəyişənlərlə işə salın:

```
WA_WHATSAPP_TOKEN=...          # daimi token
WA_WHATSAPP_PHONE_ID=...       # Phone number ID
WA_XATIRLATMA_SAAT=24,2        # qəbuldan 24 və 2 saat əvvəl (susmaya görə 2)
WA_KLINIKA_TEL=+994 12 ...     # həkimin nömrəsi bilinməyəndə
```

API qurulmayanda sistem mesajları hazırlayır və hər biri üçün `wa.me` linki verir: WhatsApp hazır mətnlə açılır, «Göndər»i istifadəçi özü basır.

**Məxfilik:** mesajda müraciət səbəbi (sağlamlıq məlumatı) var. Xəstənin WhatsApp ilə bildiriş almağa razılığı qəbula yazılarkən alınmalıdır.
