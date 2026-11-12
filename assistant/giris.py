"""Giriş: işçinin şirkət e-poçtu ilə — Outlook-a girdiyi ünvan və parol.

İki üsul var (biri seçilir):
  1) IMAP (XDMX-in mail.gov.az poçt serveri): parol poçt serverində IMAP LOGIN (SSL, 993) ilə yoxlanır.
  2) Active Directory (LDAPS): şirkətin öz domen serveri olsa.
Sistem parolu saxlamır. Uğurlu girişdən sonra brauzerə imzalı, HttpOnly kuki verilir (WA_SESSIYA_SAAT saat).

IMAP üçün mühit dəyişənləri:
  WA_IMAP_HOST    - poçt serveri, məs. mail.gov.az
  WA_IMAP_PORT    - 993 (SSL)
  WA_IMAP_DOMEN   - icazəli e-poçt domeni, məs. asf.gov.az  (mail.gov.az-da başqa qurumların işçiləri girə bilməsin;
                    «ad.soyad» yazılsa, @asf.gov.az əlavə olunur). Vergüllə bir neçə domen olar.
  WA_IMAP_CA      - (istəyə bağlı) poçt serverinin sertifikatını imzalayan CA faylı
  WA_IMAP_SERT_ADI - sertifikatın üzərindəki ad host adından fərqlidirsə (mail.gov.az: sertifikat «*.mail.gov.az»-dır,
                    bu isə «mail.gov.az»-ın özünü əhatə etmir). Zəncir və vaxt adi qaydada yoxlanır, ad isə bununla
                    tutuşdurulur. Sertifikat yenilənəndə də işləyir (WA_IMAP_PIN-dən fərqli olaraq).
  WA_IMAP_PIN     - (müvəqqəti) poçt serverinin sertifikatının SHA-256 barmaq izi. Server sertifikatı vaxtı keçmiş və ya
                    başqa ada verilmiş olanda (mail.gov.az-da 2026-10: CN=cg.gov.az, 06.10.2026-da bitib) adi yoxlama
                    keçmir. Yoxlamanı söndürmək əvəzinə məhz bu sertifikata «sancılır»: başqa sertifikat (ortadakı
                    adam) gəlsə, giriş dayanır. XDMX sertifikatı yeniləyəndə WA_IMAP_PIN silinməlidir.
                    Barmaq izi: openssl s_client -connect mail.gov.az:993 </dev/null 2>/dev/null | openssl x509 -noout -fingerprint -sha256
  İcazəli siyahı  - (istəyə bağlı) WA_ICAZELI="a@asf.gov.az,b@asf.gov.az" və/və ya WA_DATA/icazeli_istifadeciler.txt
                    (hər sətirdə bir ünvan). Siyahı varsa, yalnız ondakılar girə bilər.

Gmail (Google ilə daxil ol, OAuth) üçün mühit dəyişənləri:
  WA_GOOGLE_CLIENT_ID     - Google Cloud Console → APIs & Services → Credentials → «OAuth client ID» (Web application)
  WA_GOOGLE_CLIENT_SECRET - həmin müştərinin sirri (Docker secret «wa_google_secret» də olar)
  WA_GOOGLE_REDIRECT      - (istəyə bağlı) geri dönüş ünvanı; verilməsə sorğudan qurulur: <ünvan>/giris/google/geri.
                            Bu ünvan Google Console-da «Authorized redirect URIs»-ə yazılmalıdır.
  Gmail parolu bu sistemə heç vaxt gəlmir: istifadəçi Google-un öz səhifəsində daxil olur.
  Gmail ünvanları yalnız icazəli siyahıda olanda girə bilər (WA_ICAZELI və ya icazeli_istifadeciler.txt);
  siyahı boşdursa, Gmail ilə giriş bağlıdır (real maliyyə məlumatı hər Google hesabına açılmasın).

Active Directory üçün mühit dəyişənləri:
  WA_LDAP_URL     - domen serveri, məs. ldaps://dc01.asf.gov.az:636 (ldap://...:389 olsa, StartTLS istifadə olunur)
  WA_LDAP_DOMEN   - istifadəçi adının domeni, məs. asf.gov.az  ("ad" -> ad@asf.gov.az; "ASF\\ad" və "ad@..." də olar)
  WA_LDAP_BASE    - axtarış bazası, məs. DC=asf,DC=gov,DC=az
  WA_LDAP_QRUP    - (istəyə bağlı) yalnız bu qrupun üzvləri girə bilər, məs. CN=WA_Istifadeciler,OU=Groups,DC=asf,DC=gov,DC=az
  WA_LDAP_CA      - (istəyə bağlı) domen serverinin sertifikatını imzalayan CA faylı (qurumun daxili CA-sı)
  WA_SESSIYA_SAAT - kukinin müddəti, saat (8)
  Kuki açarı: Docker secret «wa_sessiya_acari» (/run/secrets/wa_sessiya_acari) və ya WA_SESSIYA_ACARI;
              verilməsə, hər başlanğıcda təsadüfi yaranır (server yenidən başlayanda hamı yenidən daxil olur).
"""
import base64
import hashlib
import hmac
import json
import os
import pathlib
import re
import secrets
import threading
import time

KUKI = "wa_giris"


def _env(ad, susmaya=""):
    return os.environ.get(ad, susmaya).strip()


def _secret(fayl, env):
    if _env(env):
        return _env(env)
    try:
        return pathlib.Path(fayl).read_text(encoding="utf-8").strip()
    except OSError:
        return ""


LDAP_URL = _env("WA_LDAP_URL")
DOMEN = _env("WA_LDAP_DOMEN")
BAZA = _env("WA_LDAP_BASE")
QRUP = _env("WA_LDAP_QRUP")
CA = _env("WA_LDAP_CA")
IMAP_HOST = _env("WA_IMAP_HOST")
IMAP_PORT = int(_env("WA_IMAP_PORT", "993") or 993)
IMAP_DOMENLER = [d.strip().lower() for d in _env("WA_IMAP_DOMEN").split(",") if d.strip()]
IMAP_CA = _env("WA_IMAP_CA")
IMAP_SERT_ADI = _env("WA_IMAP_SERT_ADI").lower()
# bir neçə barmaq izi vergüllə (poçt serveri bir neçə node-dan ibarətdirsə, hər birinin sertifikatı fərqli ola bilər)
IMAP_PINLER = {p for p in (re.sub(r"[^0-9a-f]", "", x.lower().replace("sha256 fingerprint=", ""))
                           for x in _env("WA_IMAP_PIN").split(",")) if p}
IMAP_PIN = bool(IMAP_PINLER)
GOOGLE_ID = _env("WA_GOOGLE_CLIENT_ID")
GOOGLE_SECRET = _secret("/run/secrets/wa_google_secret", "WA_GOOGLE_CLIENT_SECRET")
GOOGLE_REDIRECT = _env("WA_GOOGLE_REDIRECT")
# Telefon rejimi (Android APK): server yalnız telefonun özündə (127.0.0.1) açıqdır — kənardan heç kim qoşula bilmir.
# Giriş parolsuzdur: poçt ünvanı (Outlook və ya Gmail) istifadəçini tanımaq üçün soruşulur və yadda qalır.
# Android-də avtomatik (sys.getandroidapilevel), və ya WA_YERLI_GIRIS=1; söndürmək: WA_YERLI_GIRIS=0.
import sys  # noqa: E402

YERLI = _env("WA_YERLI_GIRIS") == "1" or (hasattr(sys, "getandroidapilevel") and _env("WA_YERLI_GIRIS") != "0")
SESSIYA_SAAT = float(_env("WA_SESSIYA_SAAT", "8760" if YERLI else "8") or 8)  # telefonda: bir il


def _acar():
    a = _secret("/run/secrets/wa_sessiya_acari", "WA_SESSIYA_ACARI")
    if a or not YERLI:
        return a or secrets.token_hex(32)
    fayl = pathlib.Path(_env("WA_DATA", "data")) / ".sessiya_acari"  # telefonda saxlanır: tətbiq yenidən açılanda giriş qalır
    try:
        a = fayl.read_text(encoding="utf-8").strip()
    except OSError:
        a = ""
    if not a:
        a = secrets.token_hex(32)
        try:
            fayl.parent.mkdir(parents=True, exist_ok=True)
            fayl.write_text(a, encoding="utf-8")
        except OSError:
            pass
    return a


ACAR = _acar().encode()


def aktiv():
    """Giriş tələb olunurmu (poçt, domen serveri, Google və ya telefon rejimi təyin edilibsə)."""
    return bool(IMAP_HOST or LDAP_URL or GOOGLE_ID or YERLI)


def outlook_aktiv():
    return bool(IMAP_HOST or LDAP_URL or YERLI)


def google_aktiv():
    return bool(GOOGLE_ID and GOOGLE_SECRET) or YERLI


def yerli_yoxla(login, usul="outlook"):
    """Telefon rejimi: ünvan düzgün yazılıbsa istifadəçi; parol yoxlanmır. Səhv ünvan: GirisXetasi."""
    adres = (login or "").strip().lower()
    if usul == "outlook" and "@" not in adres and IMAP_DOMENLER:
        adres = f"{adres}@{IMAP_DOMENLER[0]}"
    if not re.fullmatch(r"[a-z0-9._%+\-]{1,64}@[a-z0-9.\-]+\.[a-z]{2,}", adres):
        raise GirisXetasi("Poçt ünvanını tam yazın, məs.: ad.soyad@asf.gov.az və ya ad@gmail.com")
    if usul == "gmail" and not adres.endswith(("@gmail.com", "@googlemail.com")):
        raise GirisXetasi("Gmail ünvanı @gmail.com ilə bitməlidir.")
    if usul == "outlook" and adres.endswith(("@gmail.com", "@googlemail.com")):
        raise GirisXetasi("Bu Gmail ünvanıdır — yuxarıda «Gmail»i seçin.")
    yerli = adres.split("@")[0]
    ad = " ".join(h.capitalize() for h in re.split(r"[._\-]+", yerli) if h)
    return {"login": adres, "ad": ad or adres, "mail": adres}


def icazeli_siyahi():
    """Yalnız bu ünvanlar girə bilər (boş: domendəki hamı)."""
    s = {x.strip().lower() for x in _env("WA_ICAZELI").split(",") if x.strip()}
    fayl = pathlib.Path(_env("WA_DATA", "data")) / "icazeli_istifadeciler.txt"
    try:
        s |= {x.strip().lower() for x in fayl.read_text(encoding="utf-8").splitlines() if x.strip() and not x.startswith("#")}
    except OSError:
        pass
    return s


def epoct(login):
    """«ad.soyad» və ya «ad.soyad@asf.gov.az» -> tam ünvan; icazəsiz domen və ya yanlış yazılış: None."""
    login = (login or "").strip().lower()
    if "@" not in login:
        if not IMAP_DOMENLER:
            return None
        login = f"{login}@{IMAP_DOMENLER[0]}"
    yerli, _, dom = login.partition("@")
    if not re.fullmatch(r"[a-z0-9._\-]{1,64}", yerli) or not re.fullmatch(r"[a-z0-9.\-]{3,100}", dom):
        return None
    if IMAP_DOMENLER and dom not in IMAP_DOMENLER:
        return None
    return login


def imap_yoxla(login, parol):
    """Poçt serverində IMAP LOGIN. Uğurlu: {"login", "ad", "mail"}; səhv ad/parol: None; başqa problem: GirisXetasi."""
    import imaplib
    import socket
    import ssl

    adres = epoct(login)
    if not adres:
        dom = ", ".join("@" + d for d in IMAP_DOMENLER) or "şirkət"
        raise GirisXetasi(f"Yalnız {dom} e-poçt ünvanları ilə daxil olmaq olar.")
    if not parol:
        return None
    siyahi = icazeli_siyahi()
    if siyahi and adres not in siyahi:
        raise GirisXetasi("Bu sistemə giriş icazəniz yoxdur. Administratora müraciət edin.")
    ctx = ssl.create_default_context(cafile=IMAP_CA or None)
    if IMAP_PIN:  # sertifikat zənciri yox, konkret sertifikatın barmaq izi yoxlanır (aşağıda)
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
    elif IMAP_SERT_ADI:  # zəncir və vaxt yoxlanır (CERT_REQUIRED qalır), ad isə aşağıda WA_IMAP_SERT_ADI ilə
        ctx.check_hostname = False
    try:
        m = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, ssl_context=ctx, timeout=10)
    except ssl.SSLCertVerificationError as x:
        raise GirisXetasi(f"Poçt serverinin ({IMAP_HOST}) sertifikatı yoxlanmadı: {x.verify_message}. "
                          f"Administrator: WA_IMAP_PIN-ə bax (assistant/giris.py).")
    except (OSError, ssl.SSLError, socket.timeout) as x:
        raise GirisXetasi(f"Poçt serverinə ({IMAP_HOST}) qoşulmaq alınmadı ({type(x).__name__})")
    if IMAP_SERT_ADI and not IMAP_PIN:
        sert = m.sock.getpeercert() or {}
        adlar = {v.lower() for k, v in sert.get("subjectAltName", ()) if k == "DNS"} | \
                {v.lower() for rdn in sert.get("subject", ()) for k, v in rdn if k == "commonName"}
        if IMAP_SERT_ADI not in adlar:
            try:
                m.shutdown()
            except Exception:
                pass
            print(f"[giriş] {IMAP_HOST} sertifikatının adları: {sorted(adlar)} — WA_IMAP_SERT_ADI={IMAP_SERT_ADI} deyil", flush=True)
            raise GirisXetasi(f"Poçt serverinin ({IMAP_HOST}) sertifikatı gözlənilən ada verilməyib. Parol göndərilmədi.")
    if IMAP_PIN:
        der = m.sock.getpeercert(binary_form=True) or b""
        alinan = hashlib.sha256(der).hexdigest()
        if alinan not in IMAP_PINLER:  # parol göndərilmədən dayanırıq
            sehv = [p for p in IMAP_PINLER if len(p) != 64]
            print(f"[giriş] {IMAP_HOST} sertifikatının barmaq izi: {':'.join(alinan[i:i + 2] for i in range(0, 64, 2)).upper()} — "
                  f"WA_IMAP_PIN-dəki ilə uyğun gəlmir" + (f" (WA_IMAP_PIN düzgün deyil: 64 hex simvol olmalıdır, "
                                                          f"oxunan: {sehv})" if sehv else ""), flush=True)
            try:
                m.shutdown()
            except Exception:
                pass
            raise GirisXetasi(f"Poçt serverinin ({IMAP_HOST}) sertifikatı dəyişib və gözlənilən deyil. Parol göndərilmədi. "
                              f"Administrator: alınan barmaq izi server log-undadır — yoxlayıb WA_IMAP_PIN-i yeniləsin.")
    try:
        m.login(adres, parol)
    except imaplib.IMAP4.error:
        return None
    except (OSError, ssl.SSLError) as x:
        raise GirisXetasi(f"Poçt serveri ilə əlaqə kəsildi ({type(x).__name__})")
    finally:
        try:
            m.logout()
        except Exception:
            pass
    yerli = adres.split("@")[0]
    ad = " ".join(h.capitalize() for h in re.split(r"[._\-]+", yerli) if h)
    return {"login": adres, "ad": ad or adres, "mail": adres}


class GirisXetasi(Exception):
    """İstifadəçiyə göstərilən xəta (parol səhvi deyil: server əlçatmazdır, icazə yoxdur və s.)."""


# ---------------- istifadəçi adı ----------------
def normallasdir(login):
    """«ad», «ASF\\ad», «ad@asf.gov.az» -> (sAMAccountName, userPrincipalName)."""
    login = (login or "").strip()
    if "\\" in login:
        login = login.split("\\", 1)[1]
    if "@" in login:
        sam, dom = login.split("@", 1)
    else:
        sam, dom = login, DOMEN
    sam = sam.strip()
    if not re.fullmatch(r"[A-Za-z0-9._\-]{1,64}", sam):  # LDAP filtrinə xüsusi simvol düşməsin
        return None, None
    return sam.lower(), (f"{sam}@{dom}" if dom else sam).lower()


def _ldap_esc(s):
    return "".join(f"\\{ord(c):02x}" if c in '*()\\\0' else c for c in s)


# ---------------- domen serverində yoxlama ----------------
def ad_yoxla(login, parol):
    """Domen hesabını yoxlayır. Uğurlu: {"login", "ad", "mail"}; səhv ad/parol: None; başqa problem: GirisXetasi."""
    import ssl

    from ldap3 import NONE, SIMPLE, Connection, Server, Tls
    from ldap3.core.exceptions import LDAPException

    sam, upn = normallasdir(login)
    if not sam or not parol:  # boş parol: AD-də «anonim bind» uğurlu görünə bilər — qəbul etmirik
        return None
    tls = Tls(validate=ssl.CERT_REQUIRED, ca_certs_file=CA or None)
    try:
        server = Server(LDAP_URL, use_ssl=LDAP_URL.lower().startswith("ldaps://"), tls=tls, get_info=NONE, connect_timeout=6)
        con = Connection(server, user=upn, password=parol, authentication=SIMPLE, receive_timeout=10, raise_exceptions=False)
        if not con.open():
            raise GirisXetasi("Domen serverinə qoşulmaq alınmadı")
        if LDAP_URL.lower().startswith("ldap://") and not con.start_tls():  # parol açıq getməsin
            raise GirisXetasi("Domen serveri ilə şifrəli əlaqə (StartTLS) qurulmadı")
        if not con.bind():
            if con.result.get("result") == 49:  # invalidCredentials
                return None
            raise GirisXetasi(f"Domen serveri girişi qəbul etmədi: {con.result.get('description', '')}")
        ad, mail = sam, ""
        if BAZA:
            filtr = f"(&(objectClass=user)(sAMAccountName={_ldap_esc(sam)})"
            if QRUP:  # iç-içə qruplar da nəzərə alınır (LDAP_MATCHING_RULE_IN_CHAIN)
                filtr += f"(memberOf:1.2.840.113556.1.4.1941:={_ldap_esc(QRUP)})"
            filtr += ")"
            con.search(BAZA, filtr, attributes=["displayName", "mail"], size_limit=1)
            if not con.entries:
                if QRUP:
                    raise GirisXetasi("Bu sistemə giriş icazəniz yoxdur. İT şöbəsinə müraciət edin.")
            else:
                e = con.entries[0]
                ad = str(e.displayName) if "displayName" in e and e.displayName.value else sam
                mail = str(e.mail) if "mail" in e and e.mail.value else ""
        con.unbind()
        return {"login": sam, "ad": ad, "mail": mail}
    except GirisXetasi:
        raise
    except LDAPException as x:
        raise GirisXetasi(f"Domen serverinə qoşulmaq alınmadı ({type(x).__name__})")


def yoxla(login, parol):
    return imap_yoxla(login, parol) if IMAP_HOST else ad_yoxla(login, parol)


YOXLAYICI = yoxla  # testlərdə əvəz olunur


# ---------------- Gmail: Google ilə daxil ol (OAuth 2.0, PKCE) ----------------
GOOGLE_AUTH = "https://accounts.google.com/o/oauth2/v2/auth"
GOOGLE_TOKEN = "https://oauth2.googleapis.com/token"
OAUTH_KUKI = "wa_oauth"


def google_basla(redirect):
    """Google-a yönləndirmə ünvanı və müvəqqəti kuki (state + PKCE verifier, imzalı, 10 dəqiqə)."""
    from urllib.parse import urlencode
    state, verifier = secrets.token_urlsafe(24), secrets.token_urlsafe(48)
    challenge = _b64(hashlib.sha256(verifier.encode()).digest())
    url = GOOGLE_AUTH + "?" + urlencode({
        "client_id": GOOGLE_ID, "redirect_uri": redirect, "response_type": "code", "scope": "openid email profile",
        "state": state, "code_challenge": challenge, "code_challenge_method": "S256", "prompt": "select_account"})
    return url, _imzala({"s": state, "v": verifier, "r": redirect, "exp": int(time.time() + 600)})


def google_kuki_oxu(deyer, state):
    """Geri dönüşdə: kukidəki state Google-un qaytardığı ilə eynidirsə (verifier, redirect); deyilsə None."""
    d = _imza_oxu(deyer)
    if not d or d.get("exp", 0) < time.time() or not state or not hmac.compare_digest(d.get("s", ""), state):
        return None
    return d["v"], d["r"]


def google_iddia_yoxla(iddia):
    """id_token-dakı məlumat: düzgün müştəri, təsdiqlənmiş e-poçt, icazəli siyahı. Uğurlu: istifadəçi; yoxsa GirisXetasi."""
    if iddia.get("aud") != GOOGLE_ID or iddia.get("iss") not in ("https://accounts.google.com", "accounts.google.com") \
            or iddia.get("exp", 0) < time.time():
        raise GirisXetasi("Google cavabı yoxlanmadı. Yenidən cəhd edin.")
    mail = str(iddia.get("email", "")).strip().lower()
    if not mail or iddia.get("email_verified") not in (True, "true"):
        raise GirisXetasi("Google hesabının e-poçtu təsdiqlənməyib.")
    siyahi = icazeli_siyahi()
    if not siyahi:
        raise GirisXetasi("Gmail ilə giriş üçün icazəli istifadəçilər siyahısı qurulmayıb. Administratora müraciət edin.")
    if mail not in siyahi:
        raise GirisXetasi(f"{mail} ünvanının bu sistemə giriş icazəsi yoxdur. Administratora müraciət edin.")
    yerli = mail.split("@")[0]
    ad = iddia.get("name") or " ".join(h.capitalize() for h in re.split(r"[._\-]+", yerli) if h)
    return {"login": mail, "ad": ad or mail, "mail": mail}


def google_yoxla(kod, verifier, redirect):
    """Google-dan alınan kodu tokenə dəyişir. id_token birbaşa Google-un token ünvanından (HTTPS) gəlir,
    ona görə imzasız oxumaq olar (Google sənədləri); sahələr yenə də yoxlanır."""
    import httpx
    try:
        r = httpx.post(GOOGLE_TOKEN, data={"code": kod, "client_id": GOOGLE_ID, "client_secret": GOOGLE_SECRET,
                                           "redirect_uri": redirect, "grant_type": "authorization_code",
                                           "code_verifier": verifier}, timeout=10)
    except httpx.HTTPError as x:
        raise GirisXetasi(f"Google ilə əlaqə alınmadı ({type(x).__name__})")
    if r.status_code != 200 or "id_token" not in r.json():
        raise GirisXetasi("Google girişi təsdiqlənmədi. Yenidən cəhd edin.")
    try:
        iddia = json.loads(_b64d(r.json()["id_token"].split(".")[1]))
    except (ValueError, IndexError):
        raise GirisXetasi("Google cavabı oxunmadı.")
    return google_iddia_yoxla(iddia)


GOOGLE_YOXLAYICI = google_yoxla  # testlərdə əvəz olunur


# ---------------- kuki ----------------
def _b64(b):
    return base64.urlsafe_b64encode(b).rstrip(b"=").decode()


def _b64d(s):
    return base64.urlsafe_b64decode(s + "=" * (-len(s) % 4))


def _imzala(d):
    yuk = _b64(json.dumps(d, ensure_ascii=False).encode())
    return f"{yuk}.{_b64(hmac.new(ACAR, yuk.encode(), hashlib.sha256).digest())}"


def _imza_oxu(deyer):
    try:
        yuk, imza = (deyer or "").split(".", 1)
        if not hmac.compare_digest(imza, _b64(hmac.new(ACAR, yuk.encode(), hashlib.sha256).digest())):
            return None
        return json.loads(_b64d(yuk))
    except (ValueError, json.JSONDecodeError):
        return None


def kuki_yarat(istifadeci):
    return _imzala({"l": istifadeci["login"], "a": istifadeci.get("ad", ""), "exp": int(time.time() + SESSIYA_SAAT * 3600)})


def kuki_oxu(deyer):
    d = _imza_oxu(deyer)
    return d if isinstance(d, dict) and d.get("exp", 0) > time.time() and d.get("l") else None


# ---------------- parol təxmininə qarşı ----------------
_UGURSUZ, _KILID = {}, threading.Lock()
MAKS_CEHD, PENCERE = 5, 600  # 10 dəqiqədə 5 uğursuz cəhd -> 10 dəqiqə gözləmə


def bloklanib(acar):
    with _KILID:
        v = [t for t in _UGURSUZ.get(acar, []) if t > time.time() - PENCERE]
        _UGURSUZ[acar] = v
        return len(v) >= MAKS_CEHD


def ugursuz(acar):
    with _KILID:
        _UGURSUZ.setdefault(acar, []).append(time.time())


def ugurlu(acar):
    with _KILID:
        _UGURSUZ.pop(acar, None)


# ---------------- giriş səhifəsi ----------------
def _yerli_sehife(xeta, login, usul, secim, mobil=""):
    """Telefon rejimi: Outlook / Gmail seçimi, yalnız poçt ünvanı (parol yoxdur — server telefonun özündədir)."""
    from html import escape

    def forma(u, izah, numune):
        return (f'''<form class="usul" id="u-{u}" method="post" action="/giris"{'' if usul == u else ' hidden'}>
<input type="hidden" name="usul" value="{u}"><p class="izah">{izah}</p>
<label for="l-{u}">E-poçt ünvanı</label>
<input id="l-{u}" name="login" type="email" inputmode="email" value="{escape(login) if usul == u else ''}" placeholder="{numune}" autocomplete="email" required>
<button type="submit">Daxil ol</button></form>''')
    dom = escape((IMAP_DOMENLER or [DOMEN or "asf.gov.az"])[0])
    bas, son = sehife(xeta, login, usul, _qabiq=True).split("<!--FORMLAR-->")
    return bas + secim + forma("outlook", "Şirkət poçtunuzu yazın — sizi tanımaq üçündür. Telefonda parol tələb olunmur: server "
                                         "telefonunuzun özündədir.", f"ad.soyad@{dom}") + \
        forma("gmail", "Gmail ünvanınızı yazın — sizi tanımaq üçündür. Telefonda parol tələb olunmur.", "ad@gmail.com") + mobil + \
        (f'<div class="xeta">{escape(xeta)}</div>' if xeta else "") + (MOBIL_JS if mobil else "") + son


# Mobil giriş: telefon nömrəsi + SMS / WhatsApp təsdiq kodu (assistant/hesab.py). app.py təyin edir: kod çatdırıla bilərmi.
MOBIL_MUMKUN = lambda: False  # noqa: E731

MOBIL_JS = """
<script>
(() => {
  const f = document.getElementById("u-mobil"), dg = document.getElementById("kodGonder"); if(!f || !dg) return;
  const info = document.getElementById("kodInfo");
  dg.onclick = async () => {
    const tel = f.tel.value.trim(); if(!tel){ f.tel.focus(); return; }
    const kanal = (f.querySelector("input[name=kanal]:checked") || {}).value || "sms";
    info.textContent = "Göndərilir…"; info.className = "info";
    let j = {};
    try { j = await (await fetch("/giris/kod", {method: "POST", body: new URLSearchParams({tel, kanal})})).json(); } catch { j = {metn: "Serverə qoşulmaq alınmadı"}; }
    info.textContent = j.metn + (j.ekranda ? "  Kod: " + j.ekranda : ""); info.className = "info " + (j.ok ? "ok" : "no");
    if(j.ok){ document.getElementById("kodHisse").hidden = false; f.kod.focus(); document.getElementById("kodGonder").textContent = "Kodu yenidən göndər"; }
  };
})();
</script>"""


def _mobil_forma(usul, tel=""):
    from html import escape
    return f'''<form class="usul" id="u-mobil" method="post" action="/giris/mobil"{'' if usul == 'mobil' else ' hidden'}>
<p class="izah">Telefon nömrənizi yazın — təsdiq kodu SMS və ya WhatsApp ilə gələcək. Poçtun parolunu unutmusunuzsa da buradan girin:
sistem sizi nömrənizlə tanıyır. İlk dəfədirsə, qısa qeydiyyat açılacaq.</p>
<label for="tel">Telefon nömrəsi</label>
<input id="tel" name="tel" type="tel" inputmode="tel" value="{escape(tel)}" placeholder="050 123 45 67" autocomplete="tel" required>
<div class="kanal"><label><input type="radio" name="kanal" value="sms" checked> SMS</label>
<label><input type="radio" name="kanal" value="whatsapp"> WhatsApp</label></div>
<button type="button" id="kodGonder">Kod göndər</button>
<div id="kodHisse" hidden><label for="kod">Təsdiq kodu</label>
<input id="kod" name="kod" inputmode="numeric" autocomplete="one-time-code" maxlength="6" placeholder="6 rəqəm">
<button type="submit">Daxil ol</button></div>
<div class="info" id="kodInfo"></div></form>'''


def _secim(usul, ol, gm, mb):
    usullar = [u for u, var in (("outlook", ol), ("gmail", gm), ("mobil", mb)) if var]
    if len(usullar) < 2:
        return ""
    ad = {"outlook": "Outlook", "gmail": "Gmail", "mobil": "Mobil"}
    return '<div class="secim" role="tablist" style="grid-template-columns:repeat(' + str(len(usullar)) + ',1fr)">' + "".join(
        f'<button type="button" role="tab" lang="en" data-u="{u}" aria-selected="{str(usul == u).lower()}">{ad[u]}</button>'
        for u in usullar) + "</div>"


def sehife(xeta="", login="", usul="", _qabiq=False, tel=""):
    """Giriş səhifəsi: Outlook (şirkət poçtu), Gmail və ya Mobil (telefon + kod). Yalnız işləyən üsullar göstərilir."""
    from html import escape
    ol, gm, mb = outlook_aktiv(), google_aktiv(), MOBIL_MUMKUN()
    usul = usul if usul in ("outlook", "gmail", "mobil") else ("outlook" if ol else "gmail" if gm else "mobil")
    secim = _secim(usul, ol, gm, mb)
    if YERLI and not _qabiq:  # telefon rejimi: hər iki üsulda yalnız ünvan (parolsuz)
        return _yerli_sehife(xeta, login, usul, secim, (_mobil_forma(usul, tel) if mb else ""))
    gmail = f'''<div class="usul" id="u-gmail"{'' if usul == 'gmail' else ' hidden'}>
<p class="izah">Google hesabınızla daxil olun. Parolu Google-un öz səhifəsində yazırsınız — bu sistemə gəlmir.</p>
<a class="google" href="/giris/google"><svg viewBox="0 0 48 48" width="18" height="18" aria-hidden="true"><path fill="#FFC107" d="M43.6 20.5H42V20H24v8h11.3C33.7 32.7 29.2 36 24 36c-6.6 0-12-5.4-12-12s5.4-12 12-12c3.1 0 5.8 1.2 7.9 3.1l5.7-5.7C34 6.1 29.3 4 24 4 12.9 4 4 12.9 4 24s8.9 20 20 20 20-8.9 20-20c0-1.3-.1-2.4-.4-3.5z"/><path fill="#FF3D00" d="M6.3 14.7l6.6 4.8C14.7 15.1 19 12 24 12c3.1 0 5.8 1.2 7.9 3.1l5.7-5.7C34 6.1 29.3 4 24 4 16.3 4 9.7 8.3 6.3 14.7z"/><path fill="#4CAF50" d="M24 44c5.2 0 9.9-2 13.4-5.2l-6.2-5.2C29.2 35.1 26.7 36 24 36c-5.2 0-9.6-3.3-11.3-7.9l-6.5 5C9.5 39.6 16.2 44 24 44z"/><path fill="#1976D2" d="M43.6 20.5H42V20H24v8h11.3c-.8 2.2-2.2 4.2-4.1 5.6l6.2 5.2C37 39.2 44 34 44 24c0-1.3-.1-2.4-.4-3.5z"/></svg>Google ilə daxil ol</a>
</div>''' if gm else ""
    return f"""<!doctype html><html lang="az"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>Worker's Assistant — giriş</title>
<link rel="icon" href="/static/favicon.svg?v=boz">
<style>
:root{{--red:#ff2a2a;--text:#e8e8ea;--muted:#8a8a90;--amber:#ffb020}}
*{{box-sizing:border-box}}body{{margin:0;min-height:100vh;display:grid;place-items:center;background:#050505;color:var(--text);
font:15px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif;padding:16px}}
.kart{{width:100%;max-width:380px;background:linear-gradient(180deg,#111113,#08080a);border:1px solid #232327;border-radius:4px;padding:28px 26px}}
.bas{{display:flex;align-items:center;gap:12px;margin-bottom:6px}}.bas img{{width:38px;height:38px}}
h1{{font:700 15px ui-monospace,Consolas,monospace;letter-spacing:.22em;text-transform:uppercase;margin:0}}
.skaner{{height:2px;background:#1a0505;margin:14px 0 18px;position:relative;overflow:hidden}}
.skaner::after{{content:"";position:absolute;top:0;left:-30%;width:30%;height:100%;background:var(--red);box-shadow:0 0 10px var(--red);animation:s 2.2s ease-in-out infinite alternate}}
@keyframes s{{to{{left:100%}}}}
p.izah{{color:var(--muted);font-size:13px;margin:0 0 16px}}
label{{display:block;font:700 10px ui-monospace,monospace;letter-spacing:.16em;text-transform:uppercase;color:var(--amber);margin:12px 0 5px}}
input{{width:100%;padding:10px 12px;background:#000;border:1px solid #2a2a2e;border-radius:2px;color:var(--text);font-size:15px}}
input:focus{{outline:none;border-color:var(--red);box-shadow:0 0 8px rgba(255,40,40,.35)}}
button{{width:100%;margin-top:20px;padding:11px;background:#1a0505;border:1px solid var(--red);color:#fff;border-radius:2px;
font:700 12px ui-monospace,monospace;letter-spacing:.2em;text-transform:uppercase;cursor:pointer}}
button:hover{{background:#2a0808;box-shadow:0 0 12px rgba(255,40,40,.45)}}
.secim{{display:grid;grid-template-columns:1fr 1fr;gap:0;margin:0 0 16px;border:1px solid #2a2a2e;border-radius:2px}}
.secim button{{margin:0;padding:9px;background:transparent;border:0;color:var(--muted);letter-spacing:.14em}}
.secim button[aria-selected=true]{{background:#1a0505;color:#fff;box-shadow:inset 0 -2px 0 var(--red)}}
.secim button:hover{{box-shadow:inset 0 -2px 0 var(--red)}}
a.google{{display:flex;align-items:center;justify-content:center;gap:10px;margin-top:8px;padding:11px;background:#fff;color:#1f1f1f;
border-radius:2px;text-decoration:none;font-weight:600;font-size:14px}}
a.google:hover{{box-shadow:0 0 12px rgba(255,255,255,.25)}}
.kanal{{display:flex;gap:18px;margin:12px 0 0;color:var(--text)}}
.kanal label{{display:flex;align-items:center;gap:6px;margin:0;font:inherit;letter-spacing:0;text-transform:none;color:var(--text)}}
.kanal input{{width:auto;accent-color:var(--red)}}
.info{{margin-top:12px;font-size:13px;color:var(--muted)}}.info.ok{{color:#3cff7a}}.info.no{{color:#ff8a80}}
.xeta{{margin-top:14px;padding:9px 11px;border:1px solid #5a0a0a;background:#160303;color:#ff8a80;font-size:13px;border-radius:2px}}
</style></head><body><div class="kart">
<div class="bas"><img src="/static/logo.svg" alt=""><h1 lang="en">Worker's Assistant</h1></div>
<div class="skaner"></div>
{"<!--FORMLAR-->" if _qabiq else secim}
{"" if _qabiq else f'''<form class="usul" id="u-outlook" method="post" action="/giris" autocomplete="on"{'' if usul == 'outlook' else ' hidden'}>
<p class="izah">Şirkət e-poçtunuzla daxil olun — Outlook-a yazdığınız ünvan və parol.</p>
<label for="l">E-poçt ünvanı</label>
<input id="l" name="login" value="{escape(login)}" placeholder="ad.soyad@{escape((IMAP_DOMENLER or [DOMEN or 'asf.gov.az'])[0])}" autocomplete="username" required>
<label for="p">Parol</label>
<input id="p" name="parol" type="password" autocomplete="current-password" required>
<button type="submit">Daxil ol</button>
</form>''' if ol else ""}
{"" if _qabiq else gmail}
{"" if _qabiq or not mb else _mobil_forma(usul, tel)}
{f'<div class="xeta">{escape(xeta)}</div>' if xeta and not _qabiq else ''}
</div>{MOBIL_JS if mb and not _qabiq else ""}
<script>
document.querySelectorAll(".secim button").forEach(b => b.onclick = () => {{
  document.querySelectorAll(".secim button").forEach(x => x.setAttribute("aria-selected", x === b));
  document.querySelectorAll(".usul").forEach(u => u.hidden = u.id !== "u-" + b.dataset.u);
  const l = document.getElementById("l"); if(b.dataset.u === "outlook" && l) l.focus();
}});
</script></body></html>"""
