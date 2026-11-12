"""İstifadəçi hesabları: telefon nömrəsi ilə tanınma, SMS / WhatsApp təsdiq kodu, qeydiyyat.

Hər hesabın açarı girişdir (login): poçt ünvanı (Outlook/Gmail) və ya — yalnız telefonla qeydiyyatdan keçibsə — nömrə
(994XXXXXXXXX). Telefon nömrəsi məcburidir və hesaba bağlanır: poçtun parolunu unudan istifadəçi nömrəsi ilə girir və
sistem onu öz hesabına (söhbətlər, peşə, tapşırıqlar) salır. Mobil giriş yalnız təsdiqlənmiş nömrə ilə mümkündür.

Kodun çatdırılması:
  - WhatsApp: WhatsApp Business Cloud API (qebul modulu ilə eyni: WA_WHATSAPP_TOKEN, WA_WHATSAPP_PHONE_ID),
    Meta-da «Authentication» kateqoriyalı şablon: WA_KOD_SABLONU (susmaya görə «giris_kodu»)
  - SMS: istənilən HTTP SMS xidməti — WA_SMS_URL, məs. https://sms.example.az/send?to={tel}&text={metn}
    (WA_SMS_METOD=POST olsa {"to": tel, "text": metn} JSON göndərilir; WA_SMS_TOKEN → Authorization: Bearer)
  - Heç biri qurulmayıbsa: telefonda (APK, server telefonun özündədir) və ya WA_KOD_TEST=1 ilə kod ekranda göstərilir;
    serverdə isə kod göndərilə bilmir — nömrə qeydiyyatda «təsdiqlənməyib» kimi yazılır, mobil giriş bağlı qalır.
Kod: 6 rəqəm, 5 dəqiqə etibarlı, 5 cəhd; eyni nömrəyə 60 saniyədə bir, saatda 5 kod.
"""
import hashlib
import hmac
import json
import os
import pathlib
import re
import secrets
import threading
import time
from urllib.parse import quote

from . import giris as G
from .modullar import qebul as QB
from .state import DATA

FAYL = DATA / "hesablar.json"
MECBURI = True  # telefon və peşə məcburidir (qeydiyyat tamamlanmadan sistem açılmır); testlərdə söndürülə bilər
KOD_OMRU, KOD_CEHD, KOD_ARA, KOD_SAAT = 300, 5, 60, 5
_KILID = threading.Lock()
_KODLAR = {}  # tel -> {"hash", "bitir", "cehd", "gonderilib": [vaxtlar]}


def _env(ad, susmaya=""):
    return os.environ.get(ad, susmaya).strip()


# ======================= hesablar =======================
def _oxu():
    try:
        return json.loads(FAYL.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _yaz(d):
    FAYL.parent.mkdir(parents=True, exist_ok=True)
    tmp = FAYL.with_suffix(".tmp")
    tmp.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(FAYL)


def al(login):
    return _oxu().get(login) if login else None


def tel_ile(tel):
    """Təsdiqlənmiş nömrəyə bağlı hesab."""
    t = QB.telefon(tel)
    return next((h for h in _oxu().values() if t and h.get("tel") == t and h.get("tel_tesdiq")), None)


MAIL_RE = re.compile(r"[a-z0-9._%+\-]{1,64}@[a-z0-9.\-]+\.[a-z]{2,}")


def poct(x):
    x = str(x or "").strip().lower()
    return x if MAIL_RE.fullmatch(x) else ""


def mail_ile(mail):
    """Poçt ünvanı bağlı olan hesab (hesabın açarı ünvanın özü və ya telefonla yaradılıb ünvan sonra bağlanıb)."""
    m = poct(mail)
    if not m:
        return None
    d = _oxu()
    return d.get(m) or next((h for h in d.values() if h.get("mail") == m), None)


def mail_tesdiqle(login):
    with _KILID:
        d = _oxu()
        if login in d:
            d[login]["mail_tesdiq"] = True
            _yaz(d)


def saxla(login, ad, tel, tesdiq, mail=""):
    t = QB.telefon(tel)
    with _KILID:
        d = _oxu()
        if t and tesdiq:  # nömrə başqa hesabda təsdiqlidirsə — oradan götürülür (bir nömrə, bir hesab)
            for k, h in d.items():
                if k != login and h.get("tel") == t and h.get("tel_tesdiq"):
                    h["tel_tesdiq"] = False
        h = d.get(login, {"login": login, "yaradildi": time.time()})
        h.update(ad=" ".join(str(ad or "").split())[:80], tel=t, tel_tesdiq=bool(tesdiq) or (h.get("tel") == t and h.get("tel_tesdiq", False)))
        m = poct(mail) or poct(login)
        if m:
            h["mail_tesdiq"] = h.get("mail_tesdiq", False) if h.get("mail") == m else (m == login)  # poçtla girib — təsdiqlidir
            h["mail"] = m
        d[login] = h
        _yaz(d)
    return h


def tam_mi(login):
    """Qeydiyyat tamamdırmı: ad, telefon, poçt ünvanı və peşə."""
    from . import profil as PR
    h = al(login)
    return bool(h and h.get("ad") and h.get("tel") and (h.get("mail") or poct(login)) and PR.al(login))


def ad_tap(login="", tel=""):
    """Ad soyad: hesabdan, tapşırıqlar siyahısından (sistemə girənlər), yoxdursa — boş (istifadəçi özü yazır)."""
    h = al(login) or (tel_ile(tel) if tel else None)
    if h and h.get("ad"):
        return h["ad"]
    from .modullar import tapsiriq as TP
    return TP._istifadeciler().get(login, "") if login and "@" in login else ""


# ======================= təsdiq kodu =======================
def kanallar():
    """Hansı kanallar işləyir: {"sms": bool, "whatsapp": bool, "ekranda": bool}."""
    return {"sms": bool(_env("WA_SMS_URL")), "whatsapp": QB.api_aktiv(), "ekranda": G.YERLI or _env("WA_KOD_TEST") == "1"}


def mobil_mumkun():
    k = kanallar()
    return k["sms"] or k["whatsapp"] or k["ekranda"]


def _hash(tel, kod):
    return hmac.new(G.ACAR, f"{tel}:{kod}".encode(), hashlib.sha256).hexdigest()


def _sms_gonder(tel, metn):
    import httpx
    url, bas = _env("WA_SMS_URL"), {}
    if _env("WA_SMS_TOKEN"):
        bas["Authorization"] = f"Bearer {_env('WA_SMS_TOKEN')}"
    if _env("WA_SMS_METOD", "GET").upper() == "POST":
        r = httpx.post(url.replace("{tel}", tel).replace("{metn}", quote(metn)), json={"to": tel, "text": metn}, headers=bas, timeout=15)
    else:
        r = httpx.get(url.replace("{tel}", quote(tel)).replace("{metn}", quote(metn)), headers=bas, timeout=15)
    if r.status_code >= 300:
        raise RuntimeError(f"SMS xidməti {r.status_code}")


def _wa_gonder(tel, kod):
    import httpx
    sablon = _env("WA_KOD_SABLONU", "giris_kodu")
    r = httpx.post(f"https://graph.facebook.com/{QB.VERSIYA}/{QB.PHONE_ID}/messages", headers={"Authorization": f"Bearer {QB.TOKEN}"},
                   timeout=15, json={"messaging_product": "whatsapp", "to": tel, "type": "template", "template": {
                       "name": sablon, "language": {"code": QB.DIL}, "components": [
                           {"type": "body", "parameters": [{"type": "text", "text": kod}]},
                           {"type": "button", "sub_type": "url", "index": "0", "parameters": [{"type": "text", "text": kod}]}]}})
    if r.status_code >= 300:
        raise RuntimeError(f"WhatsApp API {r.status_code}")


GONDERICI = {"sms": lambda tel, kod: _sms_gonder(tel, f"Worker's Assistant giriş kodu: {kod}. 5 dəqiqə etibarlıdır. Heç kimə deməyin."),
             "whatsapp": _wa_gonder}  # testlərdə əvəz olunur


def kod_gonder(tel, kanal):
    """Kod göndərir. Qaytarır: {"ok", "metn", "ekranda": kod və ya None}; xəta: {"ok": False, "metn"}."""
    t = QB.telefon(tel)
    if not t:
        return {"ok": False, "metn": "Telefon nömrəsi oxunmadı. Məs.: 050 123 45 67 və ya +994 50 123 45 67."}
    k = kanallar()
    kanal = kanal if kanal in ("sms", "whatsapp") else "sms"
    indi = time.time()
    with _KILID:
        q = _KODLAR.get(t, {"gonderilib": []})
        q["gonderilib"] = [v for v in q["gonderilib"] if v > indi - 3600]
        if q["gonderilib"] and indi - q["gonderilib"][-1] < KOD_ARA:
            return {"ok": False, "metn": f"Yeni kod {int(KOD_ARA - (indi - q['gonderilib'][-1]))} saniyədən sonra göndərilə bilər."}
        if len(q["gonderilib"]) >= KOD_SAAT:
            return {"ok": False, "metn": "Bu nömrəyə çox kod göndərilib. Bir saat sonra yenidən cəhd edin."}
        kod = f"{secrets.randbelow(10 ** 6):06d}"
        q.update(hash=_hash(t, kod), bitir=indi + KOD_OMRU, cehd=0)
        q["gonderilib"].append(indi)
        _KODLAR[t] = q
    ad = "WhatsApp" if kanal == "whatsapp" else "SMS"
    if k[kanal]:
        try:
            GONDERICI[kanal](t, kod)
        except Exception as x:
            return {"ok": False, "metn": f"{ad} ilə kod göndərilmədi ({x}). Digər üsulu seçin və ya sonra cəhd edin."}
        return {"ok": True, "metn": f"Kod {ad} ilə {QB.tel_goster(t)} nömrəsinə göndərildi.", "ekranda": None}
    if k["ekranda"]:  # telefonda (server telefonun özündədir) və ya sınaq rejimi
        return {"ok": True, "metn": f"{ad} xidməti qurulmayıb — sınaq rejimi: kod ekranda göstərilir.", "ekranda": kod}
    return {"ok": False, "metn": f"{ad} xidməti qurulmayıb — kod göndərilə bilmir. Administrator: WA_SMS_URL və ya WhatsApp API."}


def kod_yoxla(tel, kod):
    t = QB.telefon(tel)
    with _KILID:
        q = _KODLAR.get(t)
        if not q or not q.get("hash") or q["bitir"] < time.time():
            return False, "Kodun vaxtı keçib və ya kod göndərilməyib. Yenidən göndərin."
        q["cehd"] += 1
        if q["cehd"] > KOD_CEHD:
            q["hash"] = None
            return False, "Çox səhv cəhd. Yeni kod göndərin."
        if not hmac.compare_digest(q["hash"], _hash(t, re.sub(r"\D", "", str(kod or "")))):
            return False, f"Kod səhvdir (qalan cəhd: {KOD_CEHD - q['cehd']})."
        q["hash"] = None  # bir dəfəlik
    return True, ""


# ======================= qeydiyyat səhifəsi =======================
def qeydiyyat_sehifesi(xeta="", ad="", tel="", tesdiqli=False, secilmis=(), izah="", mail="", mail_sabit=False):
    """Ad soyad, telefon (məcburi, təsdiq kodu ilə) və peşə (məcburi). Giriş səhifəsinin üslubu ilə."""
    from html import escape
    from . import profil as PR
    bas, son = G.sehife(_qabiq=True).split("<!--FORMLAR-->")
    sec = set(secilmis)

    def pese(p, alt=False):
        return (f'<label class="pq{" alt" if alt else ""}"><input type="checkbox" name="peseler" value="{p["kod"]}"'
                f'{" checked" if p["kod"] in sec else ""}><span>{escape(p["ad"])}</span></label>')
    siyahi = "".join(pese(p) + "".join(pese(a, True) for a in PR.PESELER if a.get("ust") == p["kod"])
                     for p in PR.PESELER if not p.get("ust"))
    if tesdiqli:
        tel_hisse = (f'<input id="tel" name="tel" type="tel" value="{escape(QB.tel_goster(QB.telefon(tel)) or tel)}" readonly>'
                     '<div class="info ok">✓ Nömrə təsdiqlənib</div>')
    else:
        tel_hisse = (f'<input id="tel" name="tel" type="tel" inputmode="tel" value="{escape(tel)}" placeholder="050 123 45 67" '
                     'autocomplete="tel" required>')
        if mobil_mumkun():
            tel_hisse += ('<div class="kanal"><label><input type="radio" name="kanal" value="sms" checked> SMS</label>'
                          '<label><input type="radio" name="kanal" value="whatsapp"> WhatsApp</label></div>'
                          '<button type="button" id="kodGonder">Təsdiq kodu göndər</button>'
                          '<div id="kodHisse" hidden><label for="kod">Təsdiq kodu</label>'
                          '<input id="kod" name="kod" inputmode="numeric" autocomplete="one-time-code" maxlength="6" placeholder="6 rəqəm"></div>'
                          '<div class="info" id="kodInfo"></div>')
        else:
            tel_hisse += '<div class="info">SMS / WhatsApp xidməti qurulmayıb: nömrə təsdiqsiz yazılır (mobil giriş üçün sonra təsdiqlənməlidir).</div>'
    forma = (f'''<form id="u-mobil" method="post" action="/qeydiyyat">
<p class="izah">{escape(izah or "Qeydiyyatı tamamlayın — bu məlumatlar bir dəfə soruşulur.")}</p>
<label for="ad">Ad soyad</label>
<input id="ad" name="ad" value="{escape(ad)}" placeholder="Adınızı və soyadınızı yazın" autocomplete="name" required>
<label for="tel">Telefon nömrəsi (məcburi)</label>{tel_hisse}
<label for="mail">Poçt ünvanı (məcburi)</label>
<input id="mail" name="mail" type="email" inputmode="email" value="{escape(mail)}" placeholder="ad.soyad@asf.gov.az və ya ad@gmail.com"
 autocomplete="email" {"readonly" if mail_sabit else "required"}>
<div class="info">Telefonla girib çıxsanız, sonra bu poçtla girəndə qaldığınız yerdən davam edəcəksiniz.</div>
<label>Peşə (məcburi — bir neçəsini seçə bilərsiniz)</label><div class="pqlar">{siyahi}</div>
<button type="submit">Qeydiyyatı tamamla</button>
{f'<div class="xeta">{escape(xeta)}</div>' if xeta else ''}</form>''')
    stil = ("<style>.pqlar{display:grid;gap:6px;margin-top:6px}.pq{display:flex;align-items:center;gap:10px;margin:0;padding:8px 10px;"
            "border:1px solid #2a2a2e;border-radius:2px;font:14px system-ui,sans-serif;letter-spacing:0;text-transform:none;color:#e8e8ea}"
            ".pq.alt{margin-left:22px;border-left:2px solid #5a0a0a}.pq input{width:auto;accent-color:#ff2a2a}"
            ".pq:has(input:checked){border-color:#ff2a2a;background:#160404}input[readonly]{color:#9a8f8f}</style>")
    return bas + stil + forma + son + (G.MOBIL_JS if not tesdiqli and mobil_mumkun() else "")


def bagla_sehifesi(mail, h, xeta=""):
    """Poçtla giriş telefonla yaradılmış hesaba: hesabın nömrəsinə (gizlədilmiş) kod göndərilir."""
    from html import escape
    bas, son = G.sehife(_qabiq=True).split("<!--FORMLAR-->")
    t = h.get("tel", "")
    gizli = f"+{t[:3]} {t[3:5]} *** ** {t[-2:]}" if len(t) == 12 else "—"
    if not (h.get("tel_tesdiq") and mobil_mumkun()):
        forma = (f'<p class="izah">«{escape(mail)}» telefonla yaradılmış hesaba bağlıdır, amma bu nömrəyə təsdiq kodu göndərilə bilmir. '
                 'Telefonla daxil olun və ya administratora müraciət edin.</p><a class="google" href="/giris">Geri</a>')
        return bas + forma + son
    forma = f'''<form id="u-bagla" method="post" action="/giris/bagla">
<p class="izah">«{escape(mail)}» ünvanı telefonla yaradılmış hesabınıza bağlıdır. Qaldığınız yerdən davam etmək üçün hesabın
nömrəsinə (<b>{gizli}</b>) göndərilən kodu yazın.</p>
<div class="kanal"><label><input type="radio" name="kanal" value="sms" checked> SMS</label>
<label><input type="radio" name="kanal" value="whatsapp"> WhatsApp</label></div>
<button type="button" id="baglaGonder">Kod göndər</button>
<label for="kod">Təsdiq kodu</label>
<input id="kod" name="kod" inputmode="numeric" autocomplete="one-time-code" maxlength="6" placeholder="6 rəqəm" required>
<button type="submit">Davam et</button>
<div class="info" id="kodInfo"></div>
{f'<div class="xeta">{escape(xeta)}</div>' if xeta else ''}</form>
<script>
document.getElementById("baglaGonder").onclick = async () => {{
  const kanal = (document.querySelector("input[name=kanal]:checked") || {{}}).value || "sms", info = document.getElementById("kodInfo");
  let j = {{}};
  try {{ j = await (await fetch("/giris/bagla/kod", {{method: "POST", body: new URLSearchParams({{kanal}})}})).json(); }} catch {{ j = {{metn: "Serverə qoşulmaq alınmadı"}}; }}
  info.textContent = j.metn + (j.ekranda ? "  Kod: " + j.ekranda : ""); info.className = "info " + (j.ok ? "ok" : "no");
}};
</script>'''
    return bas + forma + son
