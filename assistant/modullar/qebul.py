"""Diş klinikası: qəbula yazılmış xəstələrə və həkimlərə WhatsApp şablon mesajları.

Mənbə — klinikanın qəbul jurnalı (Excel): Xəstə, Telefon, Qəbul tarixi (+ Saat), Həkim, Problem, Yazılma tarixi,
istəyə bağlı «Həkimin telefonu». Mesajlar:
  1) Təsdiq (pasiyentə): nə vaxt yazıldı, nə vaxt və hansı həkimin qəbulunda olacaq, hansı problemlə.
     Gecikəcəksə — həkimi əvvəlcədən məlumatlandırmağı xahiş edir (həkimin əlaqəsi ilə).
  2) Xatırlatma, qəbula yaxın (WA_XATIRLATMA_SAAT saat qalmış):
     - pasiyentə: qəbulu xatırladır, gecikəcəksə həkimi məlumatlandırmağı xahiş edir;
     - həkimə: kim, nə vaxt, hansı problemlə gəlir; işi yubanarsa pasiyenti məlumatlandırmağı xahiş edir (pasiyentin nömrəsi ilə).

Göndərmə iki cür olur:
  - WhatsApp Business Cloud API qurulubsa (WA_WHATSAPP_TOKEN, WA_WHATSAPP_PHONE_ID): Meta-da təsdiqlənmiş şablonlarla
    sistem özü göndərir. Əl ilə göndərmə yalnız istifadəçinin təsdiqindən sonra; avtomatik xatırlatmalar yalnız
    istifadəçi «Xatırlatmaları aktiv et» dedikdən sonra işləyir.
  - Qurulmayıbsa: hər mesaj üçün wa.me linki — basanda WhatsApp hazır mətnlə açılır, istifadəçi «Göndər»i özü basır.
Şablonların mətni Meta-da eyni olmalıdır: deploy/whatsapp_sablonlar.md.

Mühit dəyişənləri:
  WA_WHATSAPP_TOKEN, WA_WHATSAPP_PHONE_ID  - WhatsApp Business Cloud API (Meta for Developers)
  WA_WHATSAPP_VERSIYA                      - Graph API versiyası (v21.0)
  WA_WHATSAPP_DIL                          - şablonların dil kodu (az)
  WA_XATIRLATMA_SAAT                       - qəbuldan neçə saat əvvəl xatırladılsın, vergüllə (2; məs. «24,2»)
  WA_KLINIKA_TEL                           - həkimin nömrəsi bilinməyəndə pasiyentə veriləcək əlaqə (qəbul şöbəsi)
"""
import datetime as dt
import json
import os
import pathlib
import re
import threading
from urllib.parse import quote

import pandas as pd

from .excel import _sorgu, norm


def _env(ad, susmaya=""):
    return os.environ.get(ad, susmaya).strip()


TOKEN = _env("WA_WHATSAPP_TOKEN")
PHONE_ID = _env("WA_WHATSAPP_PHONE_ID")
VERSIYA = _env("WA_WHATSAPP_VERSIYA", "v21.0")
DIL = _env("WA_WHATSAPP_DIL", "az")
XATIRLATMA_SAAT = sorted({float(x) for x in _env("WA_XATIRLATMA_SAAT", "2").split(",") if x.strip()}, reverse=True) or [2.0]
KLINIKA_TEL = _env("WA_KLINIKA_TEL")
QOVLUQ = None  # state.py təyin edir: WA_DATA/qebul
_KILID = threading.Lock()


def _indi():
    return dt.datetime.now()  # testlərdə əvəz olunur


def api_aktiv():
    return bool(TOKEN and PHONE_ID)


# ======================= şablonlar =======================
# Mətn Meta-da yaradılan şablonla hərfi-hərfinə eyni olmalıdır ({{n}} — parametrlər, sıra ilə).
SABLONLAR = {
    "qebul_tesdiq": {
        "kime": "Pasiyent", "ad": "Qəbula yazılma (təsdiq)",
        "metn": "Hörmətli {{1}}! {{2}} tarixində {{3}} həkimin qəbuluna yazıldınız. Qəbul vaxtınız: {{4}}. "
                "Müraciət səbəbi: {{5}}. Gecikəcəyiniz və ya gələ bilməyəcəyiniz halda, xahiş edirik, həkimi "
                "əvvəlcədən məlumatlandırın: {{6}}.",
        "parametrler": ["xeste", "yazilma", "hekim", "vaxt", "problem", "hekim_elaqe"]},
    "qebul_xatirlatma": {
        "kime": "Pasiyent", "ad": "Qəbul xatırlatması",
        "metn": "Hörmətli {{1}}! Xatırladırıq: {{2}} {{3}} həkimin qəbulundasınız. Müraciət səbəbi: {{4}}. "
                "Gecikəcəksinizsə, xahiş edirik, həkimi dərhal məlumatlandırın: {{5}}.",
        "parametrler": ["xeste", "vaxt", "hekim", "problem", "hekim_elaqe"]},
    "hekim_xatirlatma": {
        "kime": "Həkim", "ad": "Həkimə xatırlatma",
        "metn": "Hörmətli {{1}}! {{2}} {{3}} qəbulunuzdadır. Müraciət səbəbi: {{4}}. İşiniz yubanarsa, xahiş edirik, "
                "pasiyenti əvvəlcədən məlumatlandırın: {{5}}.",
        "parametrler": ["hekim", "vaxt", "xeste", "problem", "xeste_tel_goster"]},
}


def _doldur(sablon, q):
    s = SABLONLAR[sablon]
    deyerler = [str(q[p]) for p in s["parametrler"]]
    return re.sub(r"\{\{(\d+)\}\}", lambda m: deyerler[int(m.group(1)) - 1], s["metn"]), deyerler


# ======================= telefon, tarix =======================
def telefon(x):
    """Azərbaycan nömrəsi -> 994XXXXXXXXX (WhatsApp formatı); oxunmursa ''."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    if isinstance(x, float):
        x = str(int(x))
    r = re.sub(r"\D", "", str(x))
    if r.startswith("00"):
        r = r[2:]
    if len(r) == 10 and r.startswith("0"):
        r = "994" + r[1:]
    elif len(r) == 9:
        r = "994" + r
    return r if re.fullmatch(r"994\d{9}", r) else ""


def tel_goster(t):
    return f"+{t[:3]} {t[3:5]} {t[5:8]} {t[8:10]} {t[10:]}" if t else ""


def _vaxt(tarix, saat=None):
    if tarix is None or (isinstance(tarix, float) and pd.isna(tarix)) or str(tarix).strip() == "":
        return None
    t = pd.to_datetime(str(tarix), errors="coerce", dayfirst=not re.match(r"\d{4}-", str(tarix)))
    if pd.isna(t):
        return None
    t = t.to_pydatetime()
    if saat is not None and not (isinstance(saat, float) and pd.isna(saat)) and str(saat).strip():
        if isinstance(saat, float) and 0 <= saat < 1:  # Excel vaxt kəsri
            dq = round(saat * 24 * 60)
            return t.replace(hour=dq // 60, minute=dq % 60, second=0)
        m = re.search(r"(\d{1,2})[:.](\d{2})", str(saat))
        if m:
            return t.replace(hour=int(m.group(1)), minute=int(m.group(2)), second=0)
    return t


def vaxt_goster(t):
    return f"{t:%d.%m.%Y}, saat {t:%H:%M}" if t.hour or t.minute else f"{t:%d.%m.%Y}"


# ======================= qəbul jurnalını oxumaq =======================
ROLLAR = [  # sıra vacibdir: «Həkimin telefonu» əvvəl, «Qəbul tarixi» «Yazılma tarixi»ndən ayrı
    ("hekim_tel", r"hekim\w*\s*(tel|nomre|mobil|whatsapp)"),
    ("yazilma", r"yazilma|yazildig|qeydiyyat tarix|muraciet tarix"),
    ("saat", r"^saat|qebul saat|\bvaxt$|qebul vaxt"),
    ("qebul", r"qebul|tarix"),
    ("tel", r"telefon|\btel\b|mobil|whatsapp|elaqe|nomre"),
    ("hekim", r"hekim|doktor|\bdr\b"),
    ("xeste", r"xeste|pasiyent|pasient|ad soyad|soyad|musteri"),
    ("problem", r"problem|sikayet|sebeb|xidmet|diaqnoz|muraciet"),
]


def _rollar(v):
    out = {}
    for ad in v["sutunlar"]:
        n = norm(ad).strip()
        for rol, ifade in ROLLAR:
            if rol not in out and re.search(ifade, n):
                out[rol] = ad
                break
    return out


def _vereq(s):
    """Qəbul jurnalı olan vərəq: xəstə, telefon və qəbul tarixi sütunları olmalıdır."""
    if not getattr(s, "excel", None):
        return None, None, None
    for ad, v in s.excel["vereqler"].items():
        r = _rollar(v)
        if {"xeste", "tel", "qebul"} <= set(r):
            return ad, v, r
    return None, None, None


def _hekim_tel_yaddas():
    try:
        return json.loads((pathlib.Path(QOVLUQ) / "hekim_telefonlari.json").read_text(encoding="utf-8")) if QOVLUQ else {}
    except (OSError, ValueError):
        return {}


def qebullar(s):
    """Jurnaldakı qəbullar (xam dəyərlər normallaşdırılıb)."""
    ad, v, r = _vereq(s)
    if not v:
        return None
    df = _sorgu(s, f"SELECT * FROM {v['cedvel']}")
    sql = dict(zip(v["sutunlar"], v["sql"]))
    yaddas = _hekim_tel_yaddas()
    out = []
    for x in df.to_dict("records"):
        def al(rol):
            c = r.get(rol)
            val = x.get(sql[c]) if c else None
            return None if val is None or (isinstance(val, float) and pd.isna(val)) else val
        hekim = str(al("hekim") or "").strip()
        vaxt = _vaxt(al("qebul"), al("saat"))
        yaz = _vaxt(al("yazilma"))
        htel = telefon(al("hekim_tel")) or telefon(yaddas.get(norm(hekim).strip()))
        xtel = telefon(al("tel"))
        q = {"setir": int(x["_setir"]), "xeste": str(al("xeste") or "").strip(), "tel": xtel, "hekim": hekim or "—",
             "hekim_tel": htel, "problem": str(al("problem") or "").strip() or "müayinə",
             "vaxt_dt": vaxt, "vaxt": vaxt_goster(vaxt) if vaxt else "",
             "yazilma": f"{yaz:%d.%m.%Y}" if yaz else f"{_indi():%d.%m.%Y}",
             "hekim_elaqe": tel_goster(htel) or KLINIKA_TEL or "klinikanın qəbul şöbəsi",
             "xeste_tel_goster": tel_goster(xtel) or "nömrə yoxdur"}
        q["problemler"] = [p for p, ok in [("xəstənin telefonu yoxdur və ya oxunmur", xtel), ("qəbul tarixi oxunmur", vaxt),
                                           ("xəstənin adı yoxdur", q["xeste"])] if not ok]
        out.append(q)
    return {"vereq": ad, "rollar": r, "qebullar": out}


# ======================= göndərmə =======================
def _link(tel, metn):
    return f"https://wa.me/{tel}?text={quote(metn)}"


def _api_gonder(tel, sablon, parametrler):
    import httpx
    r = httpx.post(f"https://graph.facebook.com/{VERSIYA}/{PHONE_ID}/messages",
                   headers={"Authorization": f"Bearer {TOKEN}"}, timeout=15,
                   json={"messaging_product": "whatsapp", "to": tel, "type": "template",
                         "template": {"name": sablon, "language": {"code": DIL},
                                      "components": [{"type": "body", "parameters": [{"type": "text", "text": p} for p in parametrler]}]}})
    if r.status_code >= 300:
        try:
            x = r.json().get("error", {}).get("message", "")
        except ValueError:
            x = r.text[:200]
        raise RuntimeError(f"WhatsApp API {r.status_code}: {x}")
    return True


GONDERICI = _api_gonder  # testlərdə əvəz olunur


def _mesajlar(q, nov):
    """Bir qəbul üçün mesajlar: [(sablon, kimə, ad, tel)]."""
    if nov == "tesdiq":
        return [("qebul_tesdiq", "Pasiyent", q["xeste"], q["tel"])]
    return [("qebul_xatirlatma", "Pasiyent", q["xeste"], q["tel"]), ("hekim_xatirlatma", "Həkim", q["hekim"], q["hekim_tel"])]


def _acar(nov, sablon, tel, q):
    return f"{sablon}|{tel}|{q['vaxt_dt']:%Y-%m-%dT%H:%M}" if q.get("vaxt_dt") else f"{sablon}|{tel}|?"


def _jurnal_fayli():
    return pathlib.Path(QOVLUQ or ".") / "gonderilenler.json"


def _gonderilenler():
    try:
        return json.loads(_jurnal_fayli().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _qeyd_et(acar):
    with _KILID:
        d = _gonderilenler()
        d[acar] = _indi().isoformat(timespec="seconds")
        _jurnal_fayli().parent.mkdir(parents=True, exist_ok=True)
        _jurnal_fayli().write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")


def _hazirla(qs, nov, gonder):
    """Mesaj kartları; gonder=True və API varsa — həm də göndərir (əvvəl göndərilənlər təkrar getmir)."""
    kartlar, ugurlu, xeta, kecildi = [], 0, 0, 0
    evvel = _gonderilenler()
    for q in qs:
        for sablon, kime, ad, tel in _mesajlar(q, nov):
            metn, param = _doldur(sablon, q)
            k = {"kime": kime, "ad": ad, "tel": tel_goster(tel), "metn": metn, "sablon": SABLONLAR[sablon]["ad"]}
            acar = _acar(nov, sablon, tel, q)
            if not tel:
                k["status"] = "telefon yoxdur" if kime == "Pasiyent" else "həkimin telefonu yoxdur"
            elif acar in evvel:
                k["status"] = f"göndərilib {evvel[acar][:16].replace('T', ' ')}"
                kecildi += 1
            elif gonder and api_aktiv():
                try:
                    GONDERICI(tel, sablon, param)
                    _qeyd_et(acar)
                    k["status"] = "göndərildi ✓"
                    ugurlu += 1
                except Exception as x:  # şəbəkə, token, şablon təsdiqlənməyib və s.
                    k["status"] = f"göndərilmədi: {x}"
                    xeta += 1
            else:
                k["link"] = _link(tel, metn)
            kartlar.append(k)
    return kartlar, ugurlu, xeta, kecildi


# ======================= alətlər =======================
def _yoxdur(s):
    if getattr(s, "excel", None):
        return {"metn": "Bu faylda qəbul jurnalı tapılmadı: «Xəstə», «Telefon» və «Qəbul tarixi» sütunları olan vərəq lazımdır "
                        "(istəyə bağlı: «Saat», «Həkim», «Həkimin telefonu», «Problem», «Yazılma tarixi»).", "ui": None}
    return {"metn": "Qəbul jurnalı yüklənməyib. 📎 ilə klinikanın qəbul Excel faylını əlavə edin.", "ui": None}


def qebul_siyahi(s):
    d = qebullar(s)
    if not d:
        return _yoxdur(s)
    qs = d["qebullar"]
    setirler = [[str(q["setir"]), q["xeste"], q["xeste_tel_goster"] if q["tel"] else "—", q["vaxt"] or "—", q["hekim"],
                 tel_goster(q["hekim_tel"]) or "—", q["problem"], "; ".join(q["problemler"]) or "hazır"] for q in qs]
    pr = sum(1 for q in qs if q["problemler"])
    htelsiz = sorted({q["hekim"] for q in qs if not q["hekim_tel"]})
    metn = (f"«{d['vereq']}»: {len(qs)} qəbul." + (f" {pr} qəbulda problem var (son sütuna baxın)." if pr else "") +
            (f" Həkimin telefonu yoxdur: {', '.join(htelsiz)} — yazın, məs.: «{htelsiz[0]} telefonu 050 123 45 67»." if htelsiz else "") +
            " Mesajlar üçün: «Təsdiq mesajlarını hazırla» və ya «Xatırlatmaları hazırla».")
    return {"metn": metn, "ui": {"tip": "cedvel", "basliq": f"Qəbullar · {len(qs)}",
                                 "sutunlar": ["Sətir", "Xəstə", "Telefon", "Qəbul", "Həkim", "Həkimin telefonu", "Problem", "Vəziyyət"],
                                 "setirler": setirler}}


def _sual(metn, esas, legv="Ləğv et"):
    return {"tip": "sual", "variantlar": [{"etiket": esas[0], "mesaj": esas[1], "esas": True},
                                          {"etiket": legv, "mesaj": "Mesajları göndərmə", "ikinci": True}]}


def qebul_mesajlari(s, nov="tesdiq", gonder=False, saat=None):
    """nov=tesdiq — bütün yazılmışlara təsdiq; nov=xatirlatma — yaxın saatlarda qəbulu olanlara (pasiyent + həkim)."""
    d = qebullar(s)
    if not d:
        return _yoxdur(s)
    nov = "xatirlatma" if str(nov).startswith("xat") else "tesdiq"
    indi = _indi()
    qs = [q for q in d["qebullar"] if q["vaxt_dt"] and q["vaxt_dt"] > indi]
    kecmis = sum(1 for q in d["qebullar"] if q["vaxt_dt"] and q["vaxt_dt"] <= indi)
    if nov == "xatirlatma":
        h = float(saat) if saat else 24.0
        yaxin = qs[0] if qs else None
        qs = [q for q in qs if q["vaxt_dt"] <= indi + dt.timedelta(hours=h)]
        if not qs and yaxin:
            yaxin = min((q for q in d["qebullar"] if q["vaxt_dt"] and q["vaxt_dt"] > indi), key=lambda q: q["vaxt_dt"])
            return {"metn": f"Növbəti {h:g} saatda qəbul yoxdur. Ən yaxın qəbul: {yaxin['vaxt']} — {yaxin['xeste']} "
                            f"({yaxin['hekim']}). Xatırlatmanı qəbula yaxın hazırlayın; daha geniş aralıq üçün yazın, "
                            f"məs.: «48 saatlıq xatırlatmaları hazırla».", "ui": None}
    if not qs:
        return {"metn": ("Növbəti " + (f"{int(h)} saatda" if nov == "xatirlatma" else "günlərdə") + " qəbul yoxdur."
                         + (f" Keçmiş tarixli {kecmis} qəbul nəzərə alınmadı." if kecmis else "")), "ui": None}
    kartlar, ugurlu, xeta, kecildi = _hazirla(qs, nov, gonder)
    ad = "Təsdiq" if nov == "tesdiq" else "Xatırlatma"
    if api_aktiv() and gonder:
        metn = f"{ad} mesajları: {ugurlu} göndərildi" + (f", {xeta} göndərilmədi" if xeta else "") + \
               (f", {kecildi} əvvəl göndərilmişdi" if kecildi else "") + "."
        return {"metn": metn, "ui": {"tip": "whatsapp", "kartlar": kartlar}}
    metn = f"{len(kartlar)} {ad.lower()} mesajı hazırdır ({len(qs)} qəbul" + (", pasiyent və həkim" if nov == "xatirlatma" else "") + ")."
    if kecmis:
        metn += f" Keçmiş tarixli {kecmis} qəbul nəzərə alınmadı."
    if api_aktiv():
        hazir = sum(1 for k in kartlar if "link" in k)
        metn += f" WhatsApp ilə {hazir} mesaj göndərilsin? Mətnlərə baxıb təsdiq edin."
        return {"metn": metn, "ui": {"tip": "whatsapp", "kartlar": kartlar, "sual": _sual(metn, (
            "✓ Bəli, göndər", f"{ad} mesajlarını göndər"))}}
    metn += (" WhatsApp Business API qurulmayıb: hər mesajın «WhatsApp-da aç» düyməsini basın — mətn hazır açılacaq, "
             "«Göndər»i özünüz basırsınız.")
    return {"metn": metn, "ui": {"tip": "whatsapp", "kartlar": kartlar}}


def hekim_telefonu(s, hekim, telefon_nomresi):
    t = telefon(telefon_nomresi)
    if not t:
        return {"metn": f"«{telefon_nomresi}» nömrə kimi oxunmadı. Məs.: 050 123 45 67 və ya +994 50 123 45 67.", "ui": None}
    if QOVLUQ:
        with _KILID:
            d = _hekim_tel_yaddas()
            d[norm(hekim).strip()] = t
            pathlib.Path(QOVLUQ).mkdir(parents=True, exist_ok=True)
            (pathlib.Path(QOVLUQ) / "hekim_telefonlari.json").write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
    return {"metn": f"Oldu: {hekim} — {tel_goster(t)}. Yadda saxladım, xatırlatmalar bu nömrəyə gedəcək.", "ui": None}


# ======================= avtomatik xatırlatmalar =======================
def _plan_fayli():
    return pathlib.Path(QOVLUQ or ".") / "xatirlatma_plani.json"


def xatirlatma_aktiv(s, aktiv=True):
    """Avtomatik xatırlatmalar: jurnaldakı gələcək qəbullar serverdə saxlanır, hər dəqiqə yoxlanır."""
    if not aktiv:
        try:
            _plan_fayli().unlink()
        except OSError:
            pass
        return {"metn": "Avtomatik xatırlatmalar dayandırıldı.", "ui": None}
    if not api_aktiv():
        return {"metn": "Avtomatik göndərmə üçün WhatsApp Business API qurulmalıdır (WA_WHATSAPP_TOKEN, WA_WHATSAPP_PHONE_ID). "
                        "O vaxta qədər qəbula yaxın «Xatırlatmaları hazırla» yazın — linklərlə özünüz göndərərsiniz.", "ui": None}
    d = qebullar(s)
    if not d:
        return _yoxdur(s)
    qs = [{k: (v.isoformat() if isinstance(v, dt.datetime) else v) for k, v in q.items()}
          for q in d["qebullar"] if q["vaxt_dt"] and q["vaxt_dt"] > _indi()]
    _plan_fayli().parent.mkdir(parents=True, exist_ok=True)
    _plan_fayli().write_text(json.dumps({"qebullar": qs}, ensure_ascii=False, indent=1), encoding="utf-8")
    saatlar = " və ".join(f"{h:g}" for h in XATIRLATMA_SAAT)
    return {"metn": f"Avtomatik xatırlatmalar aktivdir: {len(qs)} qəbul. Qəbuldan {saatlar} saat əvvəl pasiyentə və həkimə "
                    f"WhatsApp mesajı gedəcək. Dayandırmaq üçün: «Xatırlatmaları dayandır».", "ui": None}


def xatirlatmalari_isle(indi=None):
    """Planlayıcı: vaxtı çatmış xatırlatmaları göndərir. Göndərilənlərin sayını qaytarır."""
    if not api_aktiv():
        return 0
    try:
        plan = json.loads(_plan_fayli().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return 0
    indi = indi or _indi()
    say = 0
    for q in plan.get("qebullar", []):
        q = {**q, "vaxt_dt": dt.datetime.fromisoformat(q["vaxt_dt"])}
        qalan = (q["vaxt_dt"] - indi).total_seconds() / 3600
        if qalan <= 0 or qalan > max(XATIRLATMA_SAAT):
            continue
        h = min(x for x in XATIRLATMA_SAAT if x >= qalan)  # ən yaxın xatırlatma pəncərəsi
        for sablon, _, _, tel in _mesajlar(q, "xatirlatma"):
            esas = _acar("xatirlatma", sablon, tel, q)
            acar = esas + f"|{h:g}"
            g = _gonderilenler()
            if not tel or acar in g or (h == min(XATIRLATMA_SAAT) and esas in g):  # əl ilə artıq göndərilib
                continue
            try:
                GONDERICI(tel, sablon, _doldur(sablon, q)[1])
                _qeyd_et(acar)
                say += 1
            except Exception as x:
                print(f"[whatsapp] {sablon} göndərilmədi: {x}", flush=True)
    return say


def planlayici_baslat():
    """Server başlayanda: API qurulubsa, hər dəqiqə vaxtı çatmış xatırlatmaları göndərir."""
    if not api_aktiv():
        return None

    def dovr():
        hadise = threading.Event()
        while not hadise.wait(60):
            try:
                xatirlatmalari_isle()
            except Exception as x:
                print(f"[whatsapp] planlayıcı xətası: {x}", flush=True)
    t = threading.Thread(target=dovr, name="whatsapp-xatirlatma", daemon=True)
    t.start()
    return t


# ======================= lokal (AI-sız) rejim =======================
def qayda(s, mesaj):
    n = norm(mesaj)
    if not re.search(r"whatsapp|qebul|xatirlat|tesdiq mesaj|hekim\w* telefon|telefonu|mesajlari|gonderme", n):
        return None
    m = re.search(r"^(.+?)\s+(?:telefonu|nomresi)\s+([+\d][\d\s\-()]{7,})$", mesaj.strip(), re.I)
    if m:
        return "hekim_telefonu", {"hekim": m.group(1).strip(), "telefon_nomresi": m.group(2)}
    if not _vereq(s)[1]:
        return None
    if re.search(r"gonderme|legv", n):
        return "qebul_siyahi", {}
    if re.search(r"xatirlat\w* (dayandir|sondur|deaktiv)", n):
        return "xatirlatma_aktiv", {"aktiv": False}
    if re.search(r"xatirlat\w* (aktiv|avtomatik|qos)", n):
        return "xatirlatma_aktiv", {"aktiv": True}
    gonder = bool(re.search(r"\bgonder\b|gonderin|gonder$", n))
    if "xatirlat" in n:
        h = re.search(r"(\d{1,3})\s*saat", n)
        return "qebul_mesajlari", {"nov": "xatirlatma", "gonder": gonder, **({"saat": int(h.group(1))} if h else {})}
    if re.search(r"tesdiq|whatsapp|mesaj", n):
        return "qebul_mesajlari", {"nov": "tesdiq", "gonder": gonder}
    return "qebul_siyahi", {}


ALETLER = {
    "qebul_siyahi": (qebul_siyahi, "Diş klinikasının qəbul jurnalı: xəstələr, telefonlar, qəbul vaxtı, həkim, problem, çatışmazlıqlar.", {}),
    "qebul_mesajlari": (qebul_mesajlari, "Qəbula yazılmış xəstələrə WhatsApp mesajları: nov=tesdiq (yazılma təsdiqi) və ya "
                        "nov=xatirlatma (yaxın saatlarda qəbulu olan pasiyentə və həkimə). gonder=true YALNIZ istifadəçi göndərməyi "
                        "açıq təsdiq etdikdə.", {"nov": {"type": "string", "enum": ["tesdiq", "xatirlatma"]},
                                                 "gonder": {"type": "boolean"}, "saat": {"type": "number"}}),
    "hekim_telefonu": (hekim_telefonu, "Həkimin WhatsApp nömrəsini yadda saxlayır (jurnalda yoxdursa).",
                       {"hekim": {"type": "string"}, "telefon_nomresi": {"type": "string"}}),
    "xatirlatma_aktiv": (xatirlatma_aktiv, "Avtomatik WhatsApp xatırlatmalarını aktiv edir (aktiv=true) və ya dayandırır. "
                         "Yalnız istifadəçi açıq istədikdə.", {"aktiv": {"type": "boolean"}}),
}
MECBURI = {"hekim_telefonu": ["hekim", "telefon_nomresi"]}

MODUL = {
    "kod": "qebul",
    "ad": "Qəbullar (WhatsApp)",
    "tesvir": "Diş klinikası: qəbula yazılan xəstələrə təsdiq, qəbula yaxın pasiyentə və həkimə xatırlatma",
    "status": "aktiv",
    "teklifler": [
        {"s": "Qəbulları göstər", "b": "Qəbul jurnalı", "a": "xəstələr, vaxt, həkim, çatışmazlıqlar"},
        {"s": "Təsdiq mesajlarını hazırla", "b": "Qəbula yazılma mesajı", "a": "pasiyentə WhatsApp: tarix, həkim, problem"},
        {"s": "Xatırlatmaları hazırla", "b": "Qəbula yaxın xatırlatma", "a": "pasiyentə və həkimə WhatsApp"},
    ],
}
