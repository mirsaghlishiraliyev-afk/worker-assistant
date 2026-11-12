"""Diş klinikası: həkimlərin stoka tələbləri.

Həkim çatda yazır: «Stoka tələb: Kompozit plomb A2 — 5 ədəd, təcili». Fayl yükləmək lazım deyil.
Tələblər klinika üzrə ortaqdır (WA_DATA/stok/telebler.json): həkim öz tələblərini, klinika rəhbəri hamısını görür.
Rəhbər təsdiq və ya rədd edir («T-004 təsdiq» / «T-004 rədd»). Təsdiqlənmiş tələblər təchizatçılara sifarişə
(stok modulu) əlavə olunur. Həkim tələbi WhatsApp ilə rəhbərə də bildirə bilər.

Kim rəhbərdir: profilində «Diş klinikasının rəhbəri» peşəsi olan istifadəçi. Giriş qurulmayıbsa (bir istifadəçi,
lokal rejim) — hamı.
"""
import json
import pathlib
import re
import threading

from . import qebul as QB
from .excel import norm

QOVLUQ = None  # state.py təyin edir: WA_DATA/stok
_KILID = threading.Lock()
STATUSLAR = {"gozleyir": "gözləyir", "tesdiq": "təsdiqləndi", "redd": "rədd edildi", "sifaris": "sifariş edildi", "legv": "ləğv edildi",
             "silindi": "silindi"}
VAHIDLER = r"ədəd|eded|ampul|qutu|dəst|dest|paket|şüşə|şuşə|tübik|tubik|ml|qram|qr|kq|litr|rulon|cüt|cut|flakon|şpris|spris|karpula|vərəq"

SABLONLAR = {
    "stok_teleb": {
        "kime": "Rəhbər", "ad": "Rəhbərə stok tələbi",
        "metn": "Salam! {{1}} stoka tələb yazıb: {{2}}. Tələbə Worker's Assistant-da baxıb təsdiq və ya rədd edə bilərsiniz.",
        "parametrler": ["kim", "siyahi"]},
}


# ======================= saxlama =======================
def _fayl():
    return pathlib.Path(QOVLUQ or ".") / "telebler.json"


def _oxu():
    try:
        return json.loads(_fayl().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"telebler": []}


def _yaz(d):
    _fayl().parent.mkdir(parents=True, exist_ok=True)
    tmp = _fayl().with_suffix(".tmp")
    tmp.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(_fayl())


def hamisi():
    with _KILID:
        return _oxu()["telebler"]


def tesdiqlenmis():
    return [t for t in hamisi() if t["status"] == "tesdiq"]


def status_deyis(nomreler, status):
    with _KILID:
        d = _oxu()
        for t in d["telebler"]:
            if t["no"] in nomreler:
                t["status"] = status
        _yaz(d)


# ======================= kim =======================
def _kim(s):
    return getattr(s, "istifadeci_login", "") or "", getattr(s, "istifadeci_adi", "") or "Həkim"


def rehberdir(s):
    login, _ = _kim(s)
    if not login:
        return True  # giriş yoxdur: tək istifadəçi
    from .. import profil as PR  # dövri importun qarşısını almaq üçün
    p = PR.al(login)
    return bool(p and "klinika_rehber" in p["peseler"])


# ======================= mətndən tələb =======================
def tehlil(metn):
    """«Stoka tələb: Kompozit plomb A2 — 5 ədəd, təcili» -> {mehsul, miqdar, vahid, tecili}."""
    t = str(metn or "").strip()
    tecili = bool(re.search(r"t[əe]cili|acil|bu\s*g[üu]n", t, re.I))
    t = re.sub(r"^\s*(?:stoka?\s+)?t[əe]l[əe]b(?:\s+yaz(?:ıram|iram)?)?\s*[:\-–—]?\s*", "", t, flags=re.I)
    t = re.sub(r"[,;\s]*\b(?:t[əe]cili|acil|bu\s*g[üu]n)\b[!.,;\s]*", " ", t, flags=re.I)
    t = re.sub(r"\s*\b(?:laz[ıi]md[ıi]r?|laz[ıi]m(?:\s+olacaq)?|ehtiyac\s+var|bitib|qurtar[ıi]b)\b[!.\s]*$", "", t, flags=re.I)
    miqdar, vahid = None, ""
    m = None
    for m in re.finditer(rf"(\d+(?:[.,]\d+)?)\s*({VAHIDLER})\b", t, re.I):
        pass
    if m is None:
        m = re.search(r"(?:^|[\s,—–\-:])(\d+(?:[.,]\d+)?)\s*$", t)
    if m:
        miqdar = float(m.group(1).replace(",", "."))
        vahid = (m.group(2) if m.lastindex and m.lastindex >= 2 and m.group(2) else "").lower()
        t = (t[:m.start()] + " " + t[m.end():])
    mehsul = re.sub(r"\s{2,}", " ", t).strip(" -—–,:;.")
    return {"mehsul": mehsul, "miqdar": miqdar, "vahid": vahid, "tecili": tecili}


def _miqdar(t):
    if t.get("miqdar") is None:
        return "miqdar göstərilməyib"
    return f"{t['miqdar']:g}".replace(".", ",") + (f" {t['vahid']}" if t.get("vahid") else "")


def _setir(t):
    return f"{t['mehsul']} — {_miqdar(t)}" + (" (təcili)" if t.get("tecili") else "")


# ======================= alətlər =======================
def stok_teleb_yaz(s, mehsul, miqdar=None, vahid="", tecili=False, qeyd=""):
    mehsul = " ".join(str(mehsul or "").split())[:120]
    if not mehsul:
        s.teleb_gozle = True
        return {"metn": "Nə lazımdır? Məhsulu və miqdarı yazın, məs.: «Kompozit plomb A2 — 5 ədəd» "
                        "(təcilidirsə, «təcili» əlavə edin).", "ui": None}
    s.teleb_gozle = False
    login, ad = _kim(s)
    with _KILID:
        d = _oxu()
        no = f"T-{len(d['telebler']) + 1:03d}"
        t = {"no": no, "vaxt": QB._indi().isoformat(timespec="minutes"), "login": login, "kim": ad, "mehsul": mehsul,
             "miqdar": float(miqdar) if miqdar not in (None, "") else None, "vahid": str(vahid or "").strip(),
             "tecili": bool(tecili), "qeyd": str(qeyd or "")[:200], "status": "gozleyir"}
        d["telebler"].append(t)
        _yaz(d)
    return {"metn": f"Tələb {no} qeydə alındı: {_setir(t)}. Klinika rəhbəri sistemdə görəcək və təsdiq edəcək.",
            "ui": {"tip": "sual", "variantlar": [
                {"etiket": "Rəhbərə WhatsApp ilə bildir", "mesaj": f"{no} tələbini rəhbərə bildir", "esas": True},
                {"etiket": "Tələblərim", "mesaj": "Stok tələblərini göstər", "ikinci": True}]}}


def stok_telebleri(s, status=None):
    rehber = rehberdir(s)
    login, _ = _kim(s)
    ts = [t for t in hamisi() if t["status"] != "silindi" and (rehber or t["login"] == login)]  # silinənlər görünmür
    if status:
        ts = [t for t in ts if t["status"] == status]
    if not ts:
        return {"metn": "Stok tələbi yoxdur." + ("" if rehber else " Yazmaq üçün: «Stoka tələb: <məhsul> — <miqdar>»."), "ui": None}
    sira = {"gozleyir": 0, "tesdiq": 1, "sifaris": 2, "redd": 3, "legv": 4, "silindi": 5}
    ts = sorted(ts, key=lambda t: (sira[t["status"]], not t.get("tecili"), t["vaxt"]))
    setirler = [[t["no"], t["mehsul"], _miqdar(t), "təcili" if t.get("tecili") else "", STATUSLAR[t["status"]], t["kim"],
                 f"{t['vaxt'][8:10]}.{t['vaxt'][5:7]} {t['vaxt'][11:16]}"] for t in ts]  # telefonda vacib sütunlar əvvəl
    gozleyen = [t for t in ts if t["status"] == "gozleyir"]
    ui = {"tip": "cedvel", "basliq": "Stok tələbləri" if rehber else "Tələblərim",
          "sutunlar": ["№", "Məhsul", "Miqdar", "", "Vəziyyət", "Kim", "Tarix"], "setirler": setirler}
    metn = f"{len(ts)} tələb" + (f", {len(gozleyen)}-i qərar gözləyir" if gozleyen else "") + "."
    if rehber and gozleyen:
        metn += " Qərar üçün düymələri basın və ya yazın: «T-001 təsdiq» / «T-001 rədd»."
        var = []
        for t in gozleyen[:3]:
            var += [{"etiket": f"✓ {t['no']} təsdiq", "mesaj": f"{t['no']} təsdiq", "esas": True},
                    {"etiket": f"✗ {t['no']} rədd", "mesaj": f"{t['no']} rədd", "ikinci": True}]
        if len(gozleyen) > 1:
            var.append({"etiket": f"✓ Hamısını təsdiq et ({len(gozleyen)})", "mesaj": "Bütün tələbləri təsdiq et"})
        return {"metn": metn, "ui": {"tip": "bloklar", "bloklar": [ui, {"tip": "sual", "variantlar": var}]}}
    return {"metn": metn, "ui": ui}


def stok_teleb_qerar(s, nomre, qerar):
    """Rəhbərin qərarı: tesdiq / redd / sifaris (sifariş edildi kimi işarələ). nomre='hamisi' — gözləyənlərin hamısı."""
    q = {"tesdiq": "tesdiq", "təsdiq": "tesdiq", "redd": "redd", "rədd": "redd", "sifaris": "sifaris", "legv": "legv",
         "ləğv": "legv", "sil": "silindi", "silindi": "silindi"}.get(norm(qerar).strip(), None)
    if q in ("legv", "silindi"):  # tələbi yazan özü (və ya rəhbər) gözləyən tələbi ləğv edir / silir — məs. səhvən yazılıb
        login, _ = _kim(s)
        hedef = [t["no"] for t in hamisi() if t["no"] in _nomreler(nomre) and t["status"] == "gozleyir"
                 and (rehberdir(s) or t["login"] == login)]
        if not hedef:
            return {"metn": f"«{nomre}»: {'silinəcək' if q == 'silindi' else 'ləğv ediləcək'} gözləyən tələb tapılmadı "
                            f"(yalnız öz tələbinizi {'silə' if q == 'silindi' else 'ləğv edə'} bilərsiniz).", "ui": None}
        status_deyis(set(hedef), q)  # silinən siyahılarda görünmür, nömrəsi isə təkrarlanmır
        return {"metn": f"Oldu: {', '.join(hedef)} — {STATUSLAR[q]}.", "ui": None}
    if not rehberdir(s):
        return {"metn": "Tələbi yalnız klinika rəhbəri təsdiq və ya rədd edə bilər.", "ui": None}
    if not q:
        return {"metn": "Qərar «təsdiq», «rədd» və ya «sifariş edildi» olmalıdır.", "ui": None}
    ts = hamisi()
    if str(nomre).lower() in ("hamisi", "hamısı", "*"):
        hedef = [t["no"] for t in ts if t["status"] == "gozleyir"]
    else:
        istenen = _nomreler(nomre)
        hedef = [t["no"] for t in ts if t["no"] in istenen]
    if not hedef:
        return {"metn": f"«{nomre}» tələbi tapılmadı.", "ui": None}
    status_deyis(set(hedef), q)
    metn = f"Oldu: {', '.join(hedef)} — {STATUSLAR[q]}."
    if q == "tesdiq":
        metn += " Təsdiqlənmiş tələblər «Təchizatçılara sifariş hazırla»da sifarişə əlavə olunacaq."
    qalan = [t for t in hamisi() if t["status"] == "gozleyir"]
    if qalan and rehberdir(s):
        n = stok_telebleri(s, status="gozleyir")
        return {"metn": metn + "\n\n" + n["metn"], "ui": n["ui"]}
    return {"metn": metn, "ui": None}


def stok_teleb_rehbere(s, nomre=None, gonder=False):
    """Həkimin tələbini (və ya bütün gözləyənləri) WhatsApp ilə klinika rəhbərinə bildirir."""
    from .stok import _kart, _netice, rehber_tel
    login, ad = _kim(s)
    istenen = _nomreler(nomre)
    ts = [t for t in hamisi() if t["status"] == "gozleyir" and (not istenen or t["no"] in istenen)
          and (rehberdir(s) or t["login"] == login)]
    if not ts:
        return {"metn": "Bildiriləcək gözləyən tələb yoxdur.", "ui": None}
    kim = ", ".join(dict.fromkeys(t["kim"] for t in ts))
    siyahi = "; ".join(f"{t['no']} {_setir(t)}" for t in ts)
    tel = rehber_tel()
    k = _kart("stok_teleb", "Klinika rəhbəri", tel, [kim, siyahi], gonder,
              f"stok_teleb|{tel}|" + ",".join(t["no"] for t in ts), sablonlar=SABLONLAR)
    n = _netice([k], gonder, "Rəhbərə tələb bildirişi", "Tələbi rəhbərə göndər")
    if not tel:
        n["metn"] = ("Rəhbərin WhatsApp nömrəsi sistemdə yoxdur — rəhbər və ya administrator yazsın: «Rəhbərin telefonu 050 123 45 67». "
                     "Tələb sistemdə qeydə alınıb, rəhbər onu görəcək.")
    return n


# ======================= lokal (AI-sız) rejim =======================
def _nomreler(nomre):
    """«T-001,T-002», «t1 t2», ["T-001"] -> {"T-001", "T-002"}; boş -> set()."""
    if not nomre:
        return set()
    m = nomre if isinstance(nomre, (list, tuple)) else re.findall(r"t\s*-?\s*\d{1,4}", str(nomre), re.I)
    return {_no(x) for x in m}


def _no(x):
    return "T-" + x.upper().replace("-", "")[1:].zfill(3)


def qayda(s, mesaj):
    """Fayldan asılı deyil: tələb yazmaq, siyahı, qərar, rəhbərə bildiriş. Başqa əmr gələndə «Nə lazımdır?» gözləməsi bitir."""
    plan = _qayda(s, mesaj)
    if plan and plan[0] != "stok_teleb_yaz":
        s.teleb_gozle = False
    return plan


def _qayda(s, mesaj):
    n = norm(mesaj)
    nomreler = ",".join(dict.fromkeys(_no(x) for x in re.findall(r"\bt-?\d{1,4}\b", n)))  # bir neçə: «T-001, T-002»
    m = re.search(r"\bt-?\d{1,4}\b.*\b(tesdiq|redd|sifaris edildi|legv|sil)\b", n)
    if m:
        return "stok_teleb_qerar", {"nomre": nomreler, "qerar": "sifaris" if "sifaris" in m.group(1) else m.group(1)}
    if re.search(r"butun telebleri (tesdiq|redd)", n):
        return "stok_teleb_qerar", {"nomre": "hamisi", "qerar": "redd" if "redd" in n else "tesdiq"}
    m = re.search(r"\bt-?\d{1,4}\b.*rehber", n)
    if m:  # «T-001 və T-002 ni rəhbərə göndər» — «tələb» sözü olmasa da
        return "stok_teleb_rehbere", {"nomre": nomreler,
                                      "gonder": bool(re.search(r"\bgonder\b", n))}
    if re.search(r"telebi rehbere (gonder|bildir)", n):
        return "stok_teleb_rehbere", {"gonder": "gonder" in n}
    if re.search(r"telebleri(ni|m)?\b|^telebler\b", n) and not re.search(r"teleb\s*[:\-–—]", n):
        return "stok_telebleri", {}
    if re.search(r"^\s*(stoka?\s+)?teleb\b|^\s*teleb yaz", n):
        return "stok_teleb_yaz", tehlil(mesaj)
    # «Nə lazımdır?» sualından sonrakı cavab tələbdir — amma nömrəli (T-…) və ya əmr sözlü mesaj deyil
    if getattr(s, "teleb_gozle", False) and len(n.split()) <= 15 and not re.search(
            r"\?|goster|xulase|legv|imtina|rehber|tesdiq|redd|gonder|bildir|sifaris|\bt-?\d{1,4}\b", n):
        return "stok_teleb_yaz", tehlil(mesaj)
    if getattr(s, "teleb_gozle", False) and re.search(r"legv|imtina", n):
        s.teleb_gozle = False
    return None


ALETLER = {
    "stok_teleb_yaz": (stok_teleb_yaz, "Həkimin stoka tələbi: məhsul, miqdar, vahid, təcili. Klinika rəhbəri görür və təsdiq edir.",
                       {"mehsul": {"type": "string"}, "miqdar": {"type": "number"}, "vahid": {"type": "string"},
                        "tecili": {"type": "boolean"}, "qeyd": {"type": "string"}}),
    "stok_telebleri": (stok_telebleri, "Stok tələbləri: həkim öz tələblərini, rəhbər hamısını görür (gözləyənlər əvvəl).",
                       {"status": {"type": "string", "enum": ["gozleyir", "tesdiq", "redd", "sifaris"]}}),
    "stok_teleb_qerar": (stok_teleb_qerar, "Klinika rəhbərinin qərarı: nomre (T-001 və ya 'hamisi'), qerar: tesdiq/redd/sifaris.",
                         {"nomre": {"type": "string"}, "qerar": {"type": "string"}}),
    "stok_teleb_rehbere": (stok_teleb_rehbere, "Gözləyən tələbi WhatsApp ilə klinika rəhbərinə bildirir. gonder=true YALNIZ istifadəçi "
                           "açıq təsdiq etdikdə.", {"nomre": {"type": "string"}, "gonder": {"type": "boolean"}}),
}
MECBURI = {"stok_teleb_yaz": ["mehsul"], "stok_teleb_qerar": ["nomre", "qerar"]}

MODUL = {
    "kod": "teleb",
    "ad": "Stok tələbləri",
    "tesvir": "Diş həkimləri stoka tələb yazır, klinika rəhbəri təsdiq edir, təsdiqlənənlər sifarişə düşür",
    "status": "aktiv",
    "teklifler": [
        {"s": "Stoka tələb yaz", "b": "Stoka tələb", "a": "nə lazımdır, nə qədər, təcilidirmi"},
        {"s": "Stok tələblərini göstər", "b": "Stok tələbləri", "a": "gözləyən, təsdiqlənən, sifariş edilən"},
    ],
}
