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
  WA_IMAP_PIN     - (müvəqqəti) poçt serverinin sertifikatının SHA-256 barmaq izi. Server sertifikatı vaxtı keçmiş və ya
                    başqa ada verilmiş olanda (mail.gov.az-da 2026-10: CN=cg.gov.az, 06.10.2026-da bitib) adi yoxlama
                    keçmir. Yoxlamanı söndürmək əvəzinə məhz bu sertifikata «sancılır»: başqa sertifikat (ortadakı
                    adam) gəlsə, giriş dayanır. XDMX sertifikatı yeniləyəndə WA_IMAP_PIN silinməlidir.
                    Barmaq izi: openssl s_client -connect mail.gov.az:993 </dev/null 2>/dev/null | openssl x509 -noout -fingerprint -sha256
  İcazəli siyahı  - (istəyə bağlı) WA_ICAZELI="a@asf.gov.az,b@asf.gov.az" və/və ya WA_DATA/icazeli_istifadeciler.txt
                    (hər sətirdə bir ünvan). Siyahı varsa, yalnız ondakılar girə bilər.

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
IMAP_PIN = re.sub(r"[^0-9a-f]", "", _env("WA_IMAP_PIN").lower().replace("sha256 fingerprint=", ""))
SESSIYA_SAAT = float(_env("WA_SESSIYA_SAAT", "8") or 8)
ACAR = (_secret("/run/secrets/wa_sessiya_acari", "WA_SESSIYA_ACARI") or secrets.token_hex(32)).encode()


def aktiv():
    """Giriş tələb olunurmu (poçt və ya domen serveri təyin edilibsə)."""
    return bool(IMAP_HOST or LDAP_URL)


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
    try:
        m = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, ssl_context=ctx, timeout=10)
    except ssl.SSLCertVerificationError as x:
        raise GirisXetasi(f"Poçt serverinin ({IMAP_HOST}) sertifikatı yoxlanmadı: {x.verify_message}. "
                          f"Administrator: WA_IMAP_PIN-ə bax (assistant/giris.py).")
    except (OSError, ssl.SSLError, socket.timeout) as x:
        raise GirisXetasi(f"Poçt serverinə ({IMAP_HOST}) qoşulmaq alınmadı ({type(x).__name__})")
    if IMAP_PIN:
        der = m.sock.getpeercert(binary_form=True) or b""
        if hashlib.sha256(der).hexdigest() != IMAP_PIN:  # parol göndərilmədən dayanırıq
            try:
                m.shutdown()
            except Exception:
                pass
            raise GirisXetasi(f"Poçt serverinin ({IMAP_HOST}) sertifikatı dəyişib və gözlənilən deyil. Parol göndərilmədi. "
                              f"Administrator: sertifikatı yoxlayıb WA_IMAP_PIN-i yeniləsin (və ya XDMX sertifikatı yeniləyibsə, silsin).")
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


# ---------------- kuki ----------------
def _b64(b):
    return base64.urlsafe_b64encode(b).rstrip(b"=").decode()


def _b64d(s):
    return base64.urlsafe_b64decode(s + "=" * (-len(s) % 4))


def kuki_yarat(istifadeci):
    yuk = _b64(json.dumps({"l": istifadeci["login"], "a": istifadeci.get("ad", ""),
                           "exp": int(time.time() + SESSIYA_SAAT * 3600)}, ensure_ascii=False).encode())
    imza = _b64(hmac.new(ACAR, yuk.encode(), hashlib.sha256).digest())
    return f"{yuk}.{imza}"


def kuki_oxu(deyer):
    try:
        yuk, imza = (deyer or "").split(".", 1)
        if not hmac.compare_digest(imza, _b64(hmac.new(ACAR, yuk.encode(), hashlib.sha256).digest())):
            return None
        d = json.loads(_b64d(yuk))
        return d if d.get("exp", 0) > time.time() and d.get("l") else None
    except (ValueError, json.JSONDecodeError):
        return None


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
def sehife(xeta="", login=""):
    from html import escape
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
.xeta{{margin-top:14px;padding:9px 11px;border:1px solid #5a0a0a;background:#160303;color:#ff8a80;font-size:13px;border-radius:2px}}
</style></head><body><form class="kart" method="post" action="/giris" autocomplete="on">
<div class="bas"><img src="/static/logo.svg" alt=""><h1 lang="en">Worker's Assistant</h1></div>
<div class="skaner"></div>
<p class="izah">Şirkət e-poçtunuzla daxil olun — Outlook-a yazdığınız ünvan və parol.</p>
<label for="l">E-poçt ünvanı</label>
<input id="l" name="login" value="{escape(login)}" placeholder="ad.soyad@{escape((IMAP_DOMENLER or [DOMEN or 'asf.gov.az'])[0])}" autocomplete="username" required autofocus>
<label for="p">Parol</label>
<input id="p" name="parol" type="password" autocomplete="current-password" required>
<button type="submit">Daxil ol</button>
{f'<div class="xeta">{escape(xeta)}</div>' if xeta else ''}
</form></body></html>"""
