"""Tapşırıqlar: rəhbər işçiyə tapşırıq verir, işçi icra edir; aralarında təsdiq və məlumat sorğuları.

Rəhbər (profilində «Rəhbər» və ya onun alt peşəsi olan istifadəçi) işçinin adını və ya e-poçtunu yazıb tapşırıq verir:
    «Aynur Məmmədovaya tapşırıq: Oktyabr hesabatını hazırla, müddət 15.11.2026»
    «aynur@asf.gov.az tapşırıq: Müqaviləni yoxla, müddət 3 gün»
İşçi icranı yazır: «TP-001 icraya başladım», «TP-001 addım: məlumatlar toplandı», «TP-001 tamamlandı».
Rəhbər izləyir: «Tapşırıqlar» (verdiyi tapşırıqlar, gecikənlər əvvəl), «TP-001» (addımların tarixçəsi).
İşçi rəhbərə sorğu yazır — ayrıca və ya tapşırığa bağlı:
    «Təsdiq sorğusu: 20–22 noyabr ezamiyyəyə getməyə icazə»
    «TP-001 üzrə məlumat sorğusu: hesabat hansı formada olmalıdır?»
Rəhbər cavab verir: «S-001 təsdiq», «S-001 rədd», «S-001 cavab: PDF formasında».

Məlumat klinika/qurum üzrə ortaqdır (WA_DATA/tapsiriq/): hər kəs yalnız özünə aid olanları görür.
İşçilərin siyahısı sistemə daxil olan istifadəçilərdən yığılır (ad və e-poçt); hələ daxil olmayan işçiyə e-poçtu ilə
tapşırıq vermək olar — ilk girişdə görəcək.
"""
import datetime as dt
import json
import pathlib
import re
import threading

from . import qebul as QB
from .excel import norm

QOVLUQ = None  # state.py təyin edir: WA_DATA/tapsiriq
_KILID = threading.Lock()

STATUS = {"yeni": "yeni", "icrada": "icradadır", "tamamlandi": "tamamlandı", "legv": "ləğv edildi"}
SORGU_NOV = {"tesdiq": "Təsdiq", "melumat": "Məlumat"}
SORGU_STATUS = {"gozleyir": "cavab gözləyir", "tesdiq": "təsdiqləndi", "redd": "rədd edildi", "cavab": "cavablandı"}


# ======================= saxlama =======================
def _fayl(ad):
    return pathlib.Path(QOVLUQ or ".") / ad


def _oxu(ad, susmaya):
    try:
        return json.loads(_fayl(ad).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return susmaya


def _yaz(ad, d):
    _fayl(ad).parent.mkdir(parents=True, exist_ok=True)
    tmp = _fayl(ad).with_suffix(".tmp")
    tmp.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(_fayl(ad))


def _baza():
    return _oxu("baza.json", {"tapsiriqlar": [], "sorgular": []})


def _indi():
    return QB._indi()


# ======================= istifadəçilər =======================
def qeyd_et(login, ad):
    """Daxil olan istifadəçini siyahıya yazır (rəhbər işçini adı ilə tapa bilsin)."""
    if not login:
        return
    with _KILID:
        d = _oxu("istifadeciler.json", {})
        if d.get(login) != ad:
            d[login] = ad
            _yaz("istifadeciler.json", d)


def _istifadeciler():
    return _oxu("istifadeciler.json", {})


def _ad(login):
    return _istifadeciler().get(login) or login


def istifadeci_tap(terim):
    """Ad, soyad və ya e-poçt -> (login, ad); tapılmasa və ya bir neçə uyğun gəlsə — ValueError (izahla)."""
    terim = " ".join(str(terim or "").split()).strip(" ,:;—-")
    if not terim:
        raise ValueError("Kimə? İşçinin adını və ya e-poçtunu yazın.")
    if "@" in terim:
        login = terim.lower()
        return login, _istifadeciler().get(login, login)
    n = norm(terim).strip()
    # Azərbaycan dilində yönlük şəkilçisi: «Aynura», «Rəşadə», «Leylaya» -> kök
    namizedler = [n] + [n[:-len(s)] for s in ("ya", "ye", "a", "e") if n.endswith(s) and len(n) - len(s) >= 3]
    tapilan = {}
    for login, ad in _istifadeciler().items():
        sozler = norm(ad).split() + [norm(login.split("@")[0])]
        tam = norm(ad)
        for k in namizedler:
            if tam == k or tam.startswith(k + " ") or any(s == k for s in sozler) or (len(k) >= 4 and any(s.startswith(k) for s in sozler)):
                tapilan[login] = ad
                break
    if len(tapilan) == 1:
        return next(iter(tapilan.items()))
    if len(tapilan) > 1:
        raise ValueError(f"«{terim}» bir neçə nəfərə uyğun gəlir: " + ", ".join(f"{a} ({lg})" for lg, a in tapilan.items())
                         + ". E-poçtu ilə yazın.")
    raise ValueError(f"«{terim}» adlı işçi sistemdə tapılmadı (hələ daxil olmayıb). E-poçtu ilə yazın, məs.: "
                     f"«ad.soyad@asf.gov.az tapşırıq: …».")


def _kim(s):
    return getattr(s, "istifadeci_login", "") or "", getattr(s, "istifadeci_adi", "") or "İstifadəçi"


def rehberdir(s):
    """Tapşırıq vermək hüququ: profilində «Rəhbər» (və ya alt peşəsi) olan istifadəçi; giriş yoxdursa — hamı."""
    login, _ = _kim(s)
    if not login:
        return True
    from .. import profil as PR
    return "rehber" in PR.effektiv(PR.al(login))


# ======================= müddət =======================
def muddet_oxu(metn, indi=None):
    """«15.11.2026», «15.11», «3 gün», «2 həftə», «sabah», «bu gün», «birigün» -> date; oxunmursa None."""
    indi = indi or _indi()
    indi = indi.date() if isinstance(indi, dt.datetime) else indi
    n = norm(metn or "")
    m = re.search(r"(\d{1,2})[./](\d{1,2})(?:[./](\d{2,4}))?", n)
    if m:
        il = int(m.group(3)) if m.group(3) else indi.year
        il = il + 2000 if il < 100 else il
        try:
            d = dt.date(il, int(m.group(2)), int(m.group(1)))
        except ValueError:
            return None
        return d.replace(year=d.year + 1) if not m.group(3) and d < indi else d
    m = re.search(r"(\d+)\s*(gun|hefte|ay)", n)
    if m:
        k = int(m.group(1))
        return indi + dt.timedelta(days=k * {"gun": 1, "hefte": 7, "ay": 30}[m.group(2)])
    if "birigun" in n:
        return indi + dt.timedelta(days=2)
    if "sabah" in n:
        return indi + dt.timedelta(days=1)
    if re.search(r"bu\s*gun", n):
        return indi
    if re.search(r"hefte(nin)? sonu", n):
        return indi + dt.timedelta(days=(4 - indi.weekday()) % 7 or 7)
    return None


def _tarix(d):
    return f"{dt.date.fromisoformat(d):%d.%m.%Y}" if d else "—"


def _gecikir(t):
    return bool(t.get("muddet")) and t["status"] in ("yeni", "icrada") and dt.date.fromisoformat(t["muddet"]) < _indi().date()


def _status(t):
    return STATUS[t["status"]] + (" · ⚠ gecikir" if _gecikir(t) else "")


def _no(prefiks, x):
    r = re.sub(r"[^0-9]", "", x)
    return f"{prefiks}-{int(r):03d}" if r else ""


# ======================= tapşırıqlar =======================
def _tap(no):
    return next((t for t in _baza()["tapsiriqlar"] if t["no"] == no), None)


def _deyis(no, f):
    with _KILID:
        d = _baza()
        t = next((x for x in d["tapsiriqlar"] if x["no"] == no), None)
        if t is None:
            return None
        f(t)
        _yaz("baza.json", d)
        return t


def tapsiriq_ver(s, isci, metn, muddet=None):
    if not rehberdir(s):
        return {"metn": "Tapşırıq vermək üçün profilinizdə «Rəhbər» peşəsi olmalıdır («Peşələrim»).", "ui": None}
    metn = " ".join(str(metn or "").split())
    if not metn:
        return {"metn": "Tapşırığın mətni boşdur. Məs.: «Aynura tapşırıq: Oktyabr hesabatını hazırla, müddət 15.11».", "ui": None}
    try:
        isci_login, isci_ad = istifadeci_tap(isci)
    except ValueError as x:
        return {"metn": str(x), "ui": None}
    son = muddet_oxu(muddet) if muddet and not isinstance(muddet, dt.date) else muddet
    if muddet and not son:
        return {"metn": f"«{muddet}» müddət kimi oxunmadı. Məs.: «15.11.2026», «3 gün», «sabah».", "ui": None}
    login, ad = _kim(s)
    with _KILID:
        d = _baza()
        no = f"TP-{len(d['tapsiriqlar']) + 1:03d}"
        t = {"no": no, "vaxt": _indi().isoformat(timespec="minutes"), "rehber": login, "rehber_ad": ad,
             "isci": isci_login, "isci_ad": isci_ad, "metn": metn[:1000], "muddet": son.isoformat() if son else None,
             "status": "yeni", "addimlar": [{"vaxt": _indi().isoformat(timespec="minutes"), "kim": ad, "metn": "Tapşırıq verildi"}]}
        d["tapsiriqlar"].append(t)
        _yaz("baza.json", d)
    return {"metn": f"Tapşırıq {no} verildi: {isci_ad} — «{metn}»" + (f", müddət {_tarix(t['muddet'])}" if son else ", müddətsiz")
                    + ". İşçi sistemə girəndə görəcək; icranı «Tapşırıqlar»da izləyə bilərsiniz.", "ui": None}


def _setirler(ts, rol):
    return [[t["no"], t["metn"][:70], t["isci_ad"] if rol == "rehber" else t["rehber_ad"], _tarix(t["muddet"]), _status(t),
             t["addimlar"][-1]["metn"][:50] if t["addimlar"] else ""] for t in ts]


def _sirala(ts):
    sira = {"icrada": 1, "yeni": 0, "tamamlandi": 2, "legv": 3}
    return sorted(ts, key=lambda t: (not _gecikir(t), sira[t["status"]], t.get("muddet") or "9999", t["no"]))


def tapsiriqlar(s):
    """Mənə verilən və mənim verdiyim tapşırıqlar (gecikənlər və yenilər əvvəl)."""
    login, _ = _kim(s)
    d = _baza()["tapsiriqlar"]
    mene = _sirala([t for t in d if t["isci"] == login])
    verdiyim = _sirala([t for t in d if t["rehber"] == login])
    if not mene and not verdiyim:
        return {"metn": "Tapşırıq yoxdur." + (" Vermək üçün: «<işçinin adı və ya e-poçtu> tapşırıq: <mətn>, müddət <tarix>»."
                                              if rehberdir(s) else ""), "ui": None}
    bloklar, hisse = [], []
    if mene:
        yeni = sum(t["status"] == "yeni" for t in mene)
        hisse.append(f"sizə verilən {len(mene)}" + (f" ({yeni} yeni)" if yeni else ""))
        bloklar.append({"tip": "cedvel", "basliq": "Mənə verilən tapşırıqlar", "setirler": _setirler(mene, "isci"),
                        "sutunlar": ["№", "Tapşırıq", "Rəhbər", "Müddət", "Vəziyyət", "Son addım"]})
    if verdiyim:
        gec = sum(_gecikir(t) for t in verdiyim)
        tam = sum(t["status"] == "tamamlandi" for t in verdiyim)
        hisse.append(f"verdiyiniz {len(verdiyim)} ({tam} tamamlanıb" + (f", {gec} gecikir" if gec else "") + ")")
        bloklar.append({"tip": "cedvel", "basliq": "Verdiyim tapşırıqlar", "setirler": _setirler(verdiyim, "rehber"),
                        "sutunlar": ["№", "Tapşırıq", "İşçi", "Müddət", "Vəziyyət", "Son addım"]})
    acilacaq = [t for t in mene + verdiyim if t["status"] in ("yeni", "icrada")][:4]
    if acilacaq:
        bloklar.append({"tip": "sual", "variantlar": [{"etiket": f"{t['no']} — ətraflı", "mesaj": t["no"]} for t in acilacaq]})
    return {"metn": "Tapşırıqlar: " + "; ".join(hisse) + ". Ətraflı və addımlar üçün nömrəsini yazın, məs.: «"
                    + (acilacaq or mene or verdiyim)[0]["no"] + "».", "ui": {"tip": "bloklar", "bloklar": bloklar}}


def tapsiriq_bax(s, no):
    t = _tap(_no("TP", str(no)))
    login, _ = _kim(s)
    if not t or (login and login not in (t["rehber"], t["isci"])):
        return {"metn": f"{no} tapşırığı tapılmadı.", "ui": None}
    rol = "isci" if t["isci"] == login else "rehber"
    melumat = {"tip": "cedvel", "basliq": f"{t['no']} · {_status(t)}", "sutunlar": ["", ""], "setirler": [
        ["Tapşırıq", t["metn"]], ["Rəhbər", t["rehber_ad"]], ["İşçi", t["isci_ad"]],
        ["Verilib", t["vaxt"].replace("T", " ")], ["Müddət", _tarix(t["muddet"])], ["Vəziyyət", _status(t)]]}
    addimlar = {"tip": "cedvel", "basliq": "İcra addımları", "sutunlar": ["Vaxt", "Kim", "Addım"],
                "setirler": [[a["vaxt"].replace("T", " "), a["kim"], a["metn"]] for a in t["addimlar"]]}
    var = []
    if t["status"] in ("yeni", "icrada"):
        if rol == "isci" or not login:
            if t["status"] == "yeni":
                var.append({"etiket": "▶ İcraya başladım", "mesaj": f"{t['no']} icraya başladım", "esas": True})
            var += [{"etiket": "✎ Addım yaz", "mesaj": f"{t['no']} addım yaz"},
                    {"etiket": "✓ Tamamlandı", "mesaj": f"{t['no']} tamamlandı", "esas": t["status"] == "icrada"},
                    {"etiket": "? Rəhbərə sorğu", "mesaj": f"{t['no']} üzrə sorğu yaz", "ikinci": True}]
        if rol == "rehber" or not login:
            var += [{"etiket": "✎ Qeyd yaz", "mesaj": f"{t['no']} addım yaz", "ikinci": True},
                    {"etiket": "✗ Ləğv et", "mesaj": f"{t['no']} ləğv et", "ikinci": True}]
    elif rol == "rehber" or not login:
        var.append({"etiket": "↺ Yenidən icraya", "mesaj": f"{t['no']} yenidən icraya", "ikinci": True})
    bag = [x for x in _baza()["sorgular"] if x.get("tapsiriq") == t["no"]]
    bloklar = [melumat, addimlar]
    if bag:
        bloklar.append(_sorgu_cedveli(bag, "Bu tapşırıq üzrə sorğular"))
    if var:
        bloklar.append({"tip": "sual", "variantlar": var})
    return {"metn": f"{t['no']}: «{t['metn']}» — {_status(t)}" + (f", müddət {_tarix(t['muddet'])}" if t["muddet"] else "") + ".",
            "ui": {"tip": "bloklar", "bloklar": bloklar}}


def tapsiriq_status(s, no, status, qeyd=""):
    no = _no("TP", str(no))
    t = _tap(no)
    login, ad = _kim(s)
    if not t or (login and login not in (t["rehber"], t["isci"])):
        return {"metn": f"{no} tapşırığı tapılmadı.", "ui": None}
    if status not in STATUS:
        return {"metn": "Vəziyyət: icrada, tamamlandi və ya legv olmalıdır.", "ui": None}
    if status == "legv" and login and login != t["rehber"]:
        return {"metn": "Tapşırığı yalnız onu verən rəhbər ləğv edə bilər.", "ui": None}
    if t["status"] in ("tamamlandi", "legv") and status != "icrada":
        return {"metn": f"{no} artıq «{STATUS[t['status']]}» vəziyyətindədir.", "ui": None}
    if t["status"] in ("tamamlandi", "legv") and login and login != t["rehber"]:
        return {"metn": "Bağlanmış tapşırığı yalnız rəhbər yenidən icraya qaytara bilər.", "ui": None}
    yazi = {"icrada": "İcraya başlanıldı" if t["status"] == "yeni" else "Yenidən icraya qaytarıldı",
            "tamamlandi": "Tamamlandı", "legv": "Ləğv edildi"}[status] + (f": {qeyd}" if qeyd else "")
    t = _deyis(no, lambda x: (x.update(status=status), x["addimlar"].append(
        {"vaxt": _indi().isoformat(timespec="minutes"), "kim": ad, "metn": yazi})))
    n = tapsiriq_bax(s, no)
    return {"metn": f"Oldu: {no} — {STATUS[status]}." + (" Rəhbər «Tapşırıqlar»da görəcək." if login == t["isci"] else ""),
            "ui": n["ui"]}


def tapsiriq_addim(s, no, metn=""):
    no = _no("TP", str(no))
    t = _tap(no)
    login, ad = _kim(s)
    if not t or (login and login not in (t["rehber"], t["isci"])):
        return {"metn": f"{no} tapşırığı tapılmadı.", "ui": None}
    metn = " ".join(str(metn or "").split())
    if not metn:
        s.tp_gozle = {"nov": "addim", "no": no}
        return {"metn": f"{no} üzrə hansı addımı atdınız? Qısa yazın (məs.: «Məlumatlar toplandı, cədvəl hazırlanır»).", "ui": None}
    s.tp_gozle = None
    yeni = t["status"] == "yeni" and login == t["isci"]

    def f(x):
        x["addimlar"].append({"vaxt": _indi().isoformat(timespec="minutes"), "kim": ad, "metn": metn[:500]})
        if yeni:
            x["status"] = "icrada"  # işçi addım yazdısa, icraya başlayıb
    _deyis(no, f)
    return {"metn": f"Addım yazıldı: {no} — «{metn}»." + (" Vəziyyət: icradadır." if yeni else ""), "ui": tapsiriq_bax(s, no)["ui"]}


def tapsiriq_muddet(s, no, muddet):
    no = _no("TP", str(no))
    t = _tap(no)
    login, ad = _kim(s)
    if not t or (login and login != t["rehber"]):
        return {"metn": "Müddəti yalnız tapşırığı verən rəhbər dəyişə bilər.", "ui": None}
    son = muddet_oxu(muddet)
    if not son:
        return {"metn": f"«{muddet}» müddət kimi oxunmadı.", "ui": None}
    _deyis(no, lambda x: (x.update(muddet=son.isoformat()), x["addimlar"].append(
        {"vaxt": _indi().isoformat(timespec="minutes"), "kim": ad, "metn": f"Müddət dəyişdi: {son:%d.%m.%Y}"})))
    return {"metn": f"Oldu: {no} müddəti — {son:%d.%m.%Y}.", "ui": None}


# ======================= sorğular (işçi -> rəhbər) =======================
def _sorgu_cedveli(ss, basliq):
    return {"tip": "cedvel", "basliq": basliq, "sutunlar": ["№", "Növ", "Kimdən", "Kimə", "Tapşırıq", "Sorğu", "Vəziyyət", "Cavab"],
            "setirler": [[x["no"], SORGU_NOV[x["nov"]], x["kimden_ad"], x["kime_ad"], x.get("tapsiriq") or "—", x["metn"][:80],
                          SORGU_STATUS[x["status"]], x.get("cavab") or ""] for x in ss]}


def _rehberim(login):
    """İşçinin rəhbəri: ona ən son tapşırıq verən."""
    ts = [t for t in _baza()["tapsiriqlar"] if t["isci"] == login]
    return (ts[-1]["rehber"], ts[-1]["rehber_ad"]) if ts else (None, None)


def sorgu_yarat(s, nov, metn, tapsiriq=None, kime=None):
    nov = "tesdiq" if str(nov).startswith("tesd") or str(nov).startswith("təsd") else "melumat"
    metn = " ".join(str(metn or "").split())
    login, ad = _kim(s)
    tno = _no("TP", str(tapsiriq)) if tapsiriq else None
    if tno:
        t = _tap(tno)
        if not t or (login and login not in (t["rehber"], t["isci"])):
            return {"metn": f"{tapsiriq} tapşırığı tapılmadı.", "ui": None}
        kime_login, kime_ad = (t["rehber"], t["rehber_ad"]) if login != t["rehber"] else (t["isci"], t["isci_ad"])
    elif kime:
        try:
            kime_login, kime_ad = istifadeci_tap(kime)
        except ValueError as x:
            return {"metn": str(x), "ui": None}
    else:
        kime_login, kime_ad = _rehberim(login)
        if not kime_login:
            return {"metn": "Sorğunu kimə göndərim? Rəhbərin e-poçtunu yazın, məs.: «rehber@asf.gov.az təsdiq sorğusu: …».", "ui": None}
    if not metn:
        s.tp_gozle = {"nov": "sorgu", "sorgu_nov": nov, "tapsiriq": tno, "kime": kime_login}
        return {"metn": f"{SORGU_NOV[nov]} sorğusunun mətnini yazın" + (f" ({tno} üzrə)" if tno else "") + f" — {kime_ad} görəcək.",
                "ui": None}
    s.tp_gozle = None
    with _KILID:
        d = _baza()
        no = f"S-{len(d['sorgular']) + 1:03d}"
        x = {"no": no, "vaxt": _indi().isoformat(timespec="minutes"), "nov": nov, "kimden": login, "kimden_ad": ad,
             "kime": kime_login, "kime_ad": kime_ad, "tapsiriq": tno, "metn": metn[:1000], "status": "gozleyir", "cavab": ""}
        d["sorgular"].append(x)
        if tno:  # tapşırığın tarixçəsində də görünsün
            t = next(t for t in d["tapsiriqlar"] if t["no"] == tno)
            t["addimlar"].append({"vaxt": x["vaxt"], "kim": ad, "metn": f"{SORGU_NOV[nov]} sorğusu {no}: {metn[:200]}"})
        _yaz("baza.json", d)
    return {"metn": f"{SORGU_NOV[nov]} sorğusu {no} göndərildi: {kime_ad}" + (f" ({tno} üzrə)" if tno else "") +
                    f" — «{metn}». Cavabı «Sorğular»da görəcəksiniz.", "ui": None}


def sorgular(s):
    login, _ = _kim(s)
    d = _baza()["sorgular"]
    gelen = [x for x in d if x["kime"] == login]
    gonderdiyim = [x for x in d if x["kimden"] == login]
    if not gelen and not gonderdiyim:
        return {"metn": "Sorğu yoxdur. Rəhbərə sorğu üçün: «Təsdiq sorğusu: …» və ya «Məlumat sorğusu: …» "
                        "(tapşırığa bağlı: «TP-001 üzrə təsdiq sorğusu: …»).", "ui": None}
    gelen = sorted(gelen, key=lambda x: (x["status"] != "gozleyir", x["vaxt"]))
    bloklar = []
    if gelen:
        bloklar.append(_sorgu_cedveli(gelen, "Mənə gələn sorğular"))
    if gonderdiyim:
        bloklar.append(_sorgu_cedveli(sorted(gonderdiyim, key=lambda x: x["vaxt"], reverse=True), "Göndərdiyim sorğular"))
    gozleyen = [x for x in gelen if x["status"] == "gozleyir"]
    var = []
    for x in gozleyen[:3]:
        if x["nov"] == "tesdiq":
            var += [{"etiket": f"✓ {x['no']} təsdiq", "mesaj": f"{x['no']} təsdiq", "esas": True},
                    {"etiket": f"✗ {x['no']} rədd", "mesaj": f"{x['no']} rədd", "ikinci": True}]
        var.append({"etiket": f"✎ {x['no']} cavab yaz", "mesaj": f"{x['no']} cavab yaz", "ikinci": x["nov"] == "tesdiq"})
    if var:
        bloklar.append({"tip": "sual", "variantlar": var})
    metn = (f"Sizə gələn {len(gelen)} sorğu" + (f", {len(gozleyen)}-i cavab gözləyir" if gozleyen else "") + "; " if gelen else "") + \
           (f"göndərdiyiniz {len(gonderdiyim)} sorğu." if gonderdiyim else "")
    return {"metn": metn.strip("; ") + ("." if not metn.endswith(".") else ""), "ui": {"tip": "bloklar", "bloklar": bloklar}}


def sorgu_cavab(s, no, qerar="cavab", cavab=""):
    no = _no("S", str(no))
    login, ad = _kim(s)
    with _KILID:
        d = _baza()
        x = next((x for x in d["sorgular"] if x["no"] == no), None)
        if not x or (login and login != x["kime"]):
            return {"metn": f"{no} sorğusu tapılmadı (cavabı yalnız sorğunun ünvanlandığı şəxs verir).", "ui": None}
        qerar = {"tesdiq": "tesdiq", "redd": "redd"}.get(norm(qerar).strip(), "cavab")
        cavab = " ".join(str(cavab or "").split())
        if qerar == "cavab" and not cavab:
            s.tp_gozle = {"nov": "cavab", "no": no}
            return {"metn": f"{no} sorğusuna cavabınızı yazın: «{x['metn'][:120]}»", "ui": None}
        s.tp_gozle = None
        x.update(status=qerar, cavab=cavab or SORGU_STATUS[qerar], cavab_vaxt=_indi().isoformat(timespec="minutes"))
        if x.get("tapsiriq"):
            t = next((t for t in d["tapsiriqlar"] if t["no"] == x["tapsiriq"]), None)
            if t:
                t["addimlar"].append({"vaxt": x["cavab_vaxt"], "kim": ad, "metn": f"{no} — {SORGU_STATUS[qerar]}" + (f": {cavab}" if cavab else "")})
        _yaz("baza.json", d)
    return {"metn": f"Oldu: {no} — {SORGU_STATUS[qerar]}" + (f": «{cavab}»" if cavab else "") + f". {x['kimden_ad']} görəcək.", "ui": None}


# ======================= bildirişlər =======================
def bildirisler(login):
    d = _baza()
    return {"yeni": sum(1 for t in d["tapsiriqlar"] if t["isci"] == login and t["status"] == "yeni"),
            "gecikir": sum(1 for t in d["tapsiriqlar"] if login in (t["isci"], t["rehber"]) and _gecikir(t)),
            "sorgu": sum(1 for x in d["sorgular"] if x["kime"] == login and x["status"] == "gozleyir"),
            "cavab": sum(1 for x in d["sorgular"] if x["kimden"] == login and x["status"] != "gozleyir")}


# ======================= lokal (AI-sız) rejim =======================
TP_RE = re.compile(r"\b(tp\s*-?\s*\d{1,4})\b", re.I)
S_RE = re.compile(r"\b(s\s*-?\s*\d{1,4})\b", re.I)


def _muddeti_ayir(metn):
    m = re.search(r"[,.;]?\s*(?:icra\s+)?(?:müddət[iı]?|muddet[i]?|son\s+tarix|deadline)\s*[:\-—–]?\s*(?P<m>[^,;]+?)\s*[.!]?\s*$", metn, re.I)
    if m:
        return metn[:m.start()].strip(" ,.;"), m.group("m")
    return metn.strip(), None


def qayda(s, mesaj):
    n = norm(mesaj)
    gozle = getattr(s, "tp_gozle", None)
    # rəhbərin tapşırığı: «<kim> tapşırıq: <mətn>, müddət <tarix>»
    m = re.match(r"^\s*(?P<kime>[^:]+?)\s+(?:yeni\s+)?tap[şs][ıi]r[ıi]q(?:\s+ver)?\s*[:\-—–]\s*(?P<metn>.+)$", mesaj, re.I | re.S) or \
        re.match(r"^\s*tap[şs][ıi]r[ıi]q(?:\s+ver)?\s+(?P<kime>\S+@\S+|[^:—–]+?)\s*[:\-—–]\s*(?P<metn>.+)$", mesaj, re.I | re.S)
    if m and not TP_RE.search(m.group("kime")):
        metn, muddet = _muddeti_ayir(m.group("metn"))
        return "tapsiriq_ver", {"isci": m.group("kime"), "metn": metn, **({"muddet": muddet} if muddet else {})}
    # sorğu: «[TP-001 üzrə] təsdiq/məlumat sorğusu[: mətn]», «[e-poçt] məlumat sorğusu: ...»
    m = re.search(r"(tesdiq|melumat)\s+sorgu\w*", n)
    if m or re.search(r"\bsorgu (yaz|gonder|yarat)\b|uzre sorgu", n):
        tm = TP_RE.search(mesaj)
        em = re.search(r"\S+@\S+", mesaj)
        metn = mesaj.split(":", 1)[1] if ":" in mesaj else ""
        nov = (m.group(1) if m else "melumat")
        return "sorgu_yarat", {"nov": nov, "metn": metn.strip(), **({"tapsiriq": tm.group(1)} if tm else {}),
                               **({"kime": em.group(0).rstrip(":,")} if em else {})}
    sm = S_RE.search(mesaj)
    if sm and re.search(r"\b(tesdiq|redd|cavab)", n):
        qerar = "tesdiq" if "tesdiq" in n else ("redd" if "redd" in n else "cavab")
        cavab = mesaj.split(":", 1)[1].strip() if ":" in mesaj else ""
        return "sorgu_cavab", {"no": sm.group(1), "qerar": qerar, **({"cavab": cavab} if cavab else {})}
    tm = TP_RE.search(mesaj)
    if tm:
        no = tm.group(1)
        if re.search(r"addim|qeyd", n):
            metn = mesaj.split(":", 1)[1].strip() if ":" in mesaj else ""
            return "tapsiriq_addim", {"no": no, "metn": metn}
        if re.search(r"yeniden icra|geri qaytar", n):
            return "tapsiriq_status", {"no": no, "status": "icrada"}
        if re.search(r"tamamla|icra olundu|icra etdim|bitirdim|hazirdir|yerine yetirildi", n):
            return "tapsiriq_status", {"no": no, "status": "tamamlandi", **({"qeyd": mesaj.split(":", 1)[1].strip()} if ":" in mesaj else {})}
        if re.search(r"icraya basla|basladim|icradadir|icraya gotur|icraya", n):
            return "tapsiriq_status", {"no": no, "status": "icrada"}
        if re.search(r"\blegv", n):
            return "tapsiriq_status", {"no": no, "status": "legv"}
        if re.search(r"muddet", n):
            md = re.search(r"m[üu]dd[əe]t\w*\s*[:\-—]?\s*(.+)$", mesaj, re.I)
            return "tapsiriq_muddet", {"no": no, "muddet": md.group(1) if md else ""}
        return "tapsiriq_bax", {"no": no}
    if re.search(r"^\s*(menim\s+)?tapsiriqlar|tapsiriqlarim|verdiyim tapsiriq|tapsiriqlari goster", n):
        return "tapsiriqlar", {}
    if re.search(r"^\s*sorgular|sorgularim|gelen sorgu|sorgulari goster", n):
        return "sorgular", {}
    # gözlənilən cavab: addım mətni, sorğu mətni, sorğuya cavab
    if gozle and not re.search(r"\?$|^\s*(legv|imtina)", n) and len(n.split()) <= 60:
        if gozle["nov"] == "addim":
            return "tapsiriq_addim", {"no": gozle["no"], "metn": mesaj}
        if gozle["nov"] == "cavab":
            return "sorgu_cavab", {"no": gozle["no"], "qerar": "cavab", "cavab": mesaj}
        if gozle["nov"] == "sorgu":
            return "sorgu_yarat", {"nov": gozle["sorgu_nov"], "metn": mesaj, **({"tapsiriq": gozle["tapsiriq"]} if gozle["tapsiriq"] else {}),
                                   "kime": gozle["kime"]}
    if gozle and re.search(r"^\s*(legv|imtina)", n):
        s.tp_gozle = None
    return None


ALETLER = {
    "tapsiriq_ver": (tapsiriq_ver, "Rəhbər işçiyə (adı və ya e-poçtu ilə) tapşırıq verir; muddet: '15.11.2026', '3 gün', 'sabah'.",
                     {"isci": {"type": "string"}, "metn": {"type": "string"}, "muddet": {"type": "string"}}),
    "tapsiriqlar": (tapsiriqlar, "Mənə verilən və mənim verdiyim tapşırıqlar (gecikənlər əvvəl).", {}),
    "tapsiriq_bax": (tapsiriq_bax, "Tapşırığın ətraflı vəziyyəti və icra addımlarının tarixçəsi.", {"no": {"type": "string"}}),
    "tapsiriq_status": (tapsiriq_status, "Tapşırığın vəziyyəti: icrada / tamamlandi / legv (ləğv yalnız rəhbər).",
                        {"no": {"type": "string"}, "status": {"type": "string", "enum": ["icrada", "tamamlandi", "legv"]},
                         "qeyd": {"type": "string"}}),
    "tapsiriq_addim": (tapsiriq_addim, "Tapşırıq üzrə atılan addımı (və ya rəhbərin qeydini) yazır.",
                       {"no": {"type": "string"}, "metn": {"type": "string"}}),
    "tapsiriq_muddet": (tapsiriq_muddet, "Rəhbər tapşırığın müddətini dəyişir.", {"no": {"type": "string"}, "muddet": {"type": "string"}}),
    "sorgu_yarat": (sorgu_yarat, "İşçinin rəhbərə təsdiq və ya məlumat sorğusu; ayrıca və ya tapşırığa bağlı (tapsiriq: TP-001).",
                    {"nov": {"type": "string", "enum": ["tesdiq", "melumat"]}, "metn": {"type": "string"},
                     "tapsiriq": {"type": "string"}, "kime": {"type": "string"}}),
    "sorgular": (sorgular, "Mənə gələn və göndərdiyim sorğular.", {}),
    "sorgu_cavab": (sorgu_cavab, "Sorğuya cavab: qerar tesdiq / redd / cavab (+ cavab mətni).",
                    {"no": {"type": "string"}, "qerar": {"type": "string", "enum": ["tesdiq", "redd", "cavab"]}, "cavab": {"type": "string"}}),
}
MECBURI = {"tapsiriq_ver": ["isci", "metn"], "tapsiriq_bax": ["no"], "tapsiriq_status": ["no", "status"], "tapsiriq_addim": ["no"],
           "tapsiriq_muddet": ["no", "muddet"], "sorgu_yarat": ["nov"], "sorgu_cavab": ["no"]}

MODUL = {
    "kod": "tapsiriq",
    "ad": "Tapşırıqlar",
    "tesvir": "Rəhbər tapşırıq verir və icranı izləyir; işçi addımları yazır, rəhbərə təsdiq və məlumat sorğusu göndərir",
    "status": "aktiv",
    "umumi": True,  # peşədən asılı olmayaraq hamıya göstərilir
    "teklifler": [
        {"s": "Tapşırıqlarım", "b": "Tapşırıqlar", "a": "mənə verilən və mənim verdiyim"},
        {"s": "Sorğular", "b": "Sorğular", "a": "təsdiq və məlumat sorğuları"},
    ],
}
