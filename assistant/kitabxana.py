"""İstifadəçinin yüklədiyi fayllar («Fayllarım»): kim hansı Excel-i yükləyib — yadda qalır.

Fayl bir dəfə yüklənir, sonra siyahıdan seçilib açılır (yenidən göndərmək lazım deyil; böyük fayllar keşdən
saniyələr içində açılır). Eyni adlı fayl yenidən yüklənəndə interfeys soruşur: üstünə yazım, ayrıca saxlayım,
yoxsa köhnəni açım. Eyni məzmunlu fayl təkrar saxlanmır.
"""
import hashlib
import json
import re
import threading
import time
import uuid

from .state import DATA, YUKLEMELER

QOVLUQ = DATA / "fayllar"
MAKS = 200  # istifadəçi başına
_KILID = threading.Lock()


def _fayl(login):
    return QOVLUQ / f"{hashlib.sha256((login or 'lokal').encode()).hexdigest()[:20]}.json"


def _oxu(login):
    try:
        return json.loads(_fayl(login).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"login": login, "fayllar": []}


def _yaz(login, d):
    QOVLUQ.mkdir(parents=True, exist_ok=True)
    f = _fayl(login)
    t = f.with_suffix(".tmp")
    t.write_text(json.dumps(d, ensure_ascii=False), encoding="utf-8")
    t.replace(f)


def sha1(yol):
    h = hashlib.sha1()
    with open(yol, "rb") as f:
        while b := f.read(1 << 20):
            h.update(b)
    return h.hexdigest()


def _acar(ad):
    return re.sub(r"\s+", " ", (ad or "").strip().lower())


def _ictimai(e):
    return {k: e[k] for k in ("id", "ad", "olcu", "vaxt", "nov")}


def siyahi(login):
    with _KILID:
        return [_ictimai(e) for e in _oxu(login)["fayllar"]]


def al(login, fid):
    with _KILID:
        return next((e for e in _oxu(login)["fayllar"] if e["id"] == fid), None)


def yoxla(login, fayllar):
    """Yükləmədən əvvəl: eyni adlı fayl artıq varmı? [{ad, id, vaxt, olcu, eyni_olcu}]"""
    with _KILID:
        movcud = _oxu(login)["fayllar"]
    out = []
    for f in fayllar:
        e = next((x for x in movcud if _acar(x["ad"]) == _acar(f.get("ad"))), None)
        if e:
            out.append({"ad": f.get("ad"), "id": e["id"], "vaxt": e["vaxt"], "olcu": e["olcu"],
                        "eyni_olcu": int(f.get("olcu") or -1) == e["olcu"]})
    return out


def _istifade_olunur(yol, istisna_login=None):
    """Fayl başqa qeyddə (və ya başqa istifadəçidə) istifadə olunurmu — silməzdən əvvəl."""
    for f in QOVLUQ.glob("*.json"):
        try:
            d = json.loads(f.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        if any(e["yol"] == str(yol) for e in d.get("fayllar", [])):
            return True
    return False


def _diskden_sil(yol):
    try:
        p = (YUKLEMELER / pathlib_ad(yol)).resolve()
        if p.parent == YUKLEMELER.resolve() and not _istifade_olunur(yol):
            p.unlink(missing_ok=True)
    except OSError:
        pass


def pathlib_ad(yol):
    import pathlib
    return pathlib.Path(str(yol)).name


def elave(login, ad, yol, nov, rejim=""):
    """Yüklənmiş faylı siyahıya yazır. rejim: "uzerine" — eyni adlı köhnəni əvəz edir; əks halda ayrıca saxlanır.
    Eyni məzmun artıq varsa, yeni nüsxə saxlanmır (köhnə qeyd qaytarılır). Qaytarır: (qeyd, yol)."""
    h = sha1(yol)
    olcu = yol.stat().st_size
    with _KILID:
        d = _oxu(login)
        eyni = next((e for e in d["fayllar"] if e.get("sha1") == h), None)
        if eyni:  # eyni məzmun: yeni nüsxəyə ehtiyac yoxdur
            eyni["vaxt"], eyni["ad"] = time.time(), ad
            d["fayllar"] = [eyni] + [e for e in d["fayllar"] if e is not eyni]
            _yaz(login, d)
            silinecek = yol if str(yol) != eyni["yol"] else None
            netice = (eyni, eyni["yol"])
        else:
            kohneler = [e for e in d["fayllar"] if _acar(e["ad"]) == _acar(ad)] if rejim == "uzerine" else []
            q = {"id": uuid.uuid4().hex[:12], "ad": ad, "yol": str(yol), "olcu": olcu, "sha1": h, "vaxt": time.time(), "nov": nov}
            d["fayllar"] = [q] + [e for e in d["fayllar"] if e not in kohneler]
            artiq = d["fayllar"][MAKS:]
            d["fayllar"] = d["fayllar"][:MAKS]
            _yaz(login, d)
            silinecek = None
            netice = (q, str(yol))
            kohneler += artiq
    if eyni:
        if silinecek:
            _diskden_sil(silinecek)
    else:
        for e in kohneler:
            _diskden_sil(e["yol"])
    return netice


def sil(login, fid):
    with _KILID:
        d = _oxu(login)
        e = next((x for x in d["fayllar"] if x["id"] == fid), None)
        d["fayllar"] = [x for x in d["fayllar"] if x["id"] != fid]
        _yaz(login, d)
    if e:
        _diskden_sil(e["yol"])
    return e is not None
