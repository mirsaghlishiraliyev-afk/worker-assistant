"""Söhbət tarixçəsi — hər istifadəçinin öz faylında (serverdə, WA_DATA/sohbetler/).

Əvvəl söhbətlər brauzerdə (localStorage) saxlanırdı: kompüter dəyişəndə itirdi, eyni brauzerdə başqa
istifadəçi də görürdü. İndi hər domen/e-poçt istifadəçisinin söhbətləri ayrıdır və istənilən kompüterdən açılır.
"""
import hashlib
import json
import re
import threading
import time

from .state import DATA

QOVLUQ = DATA / "sohbetler"
MAKS_SOHBET = 100            # istifadəçi başına (köhnələr silinir)
MAKS_OLCU = 3 * 1024 * 1024  # bir söhbət (cədvəllər daxil)
_KILID = threading.Lock()
ID_RE = re.compile(r"[a-z0-9]{1,24}")


def _fayl(login):
    ad = hashlib.sha256((login or "lokal").encode()).hexdigest()[:20]  # fayl adında e-poçt görünməsin
    return QOVLUQ / f"{ad}.json"


def _oxu(login):
    try:
        return json.loads(_fayl(login).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"login": login, "sohbetler": []}


def _yaz(login, d):
    QOVLUQ.mkdir(parents=True, exist_ok=True)
    f = _fayl(login)
    tmp = f.with_suffix(".tmp")
    tmp.write_text(json.dumps(d, ensure_ascii=False), encoding="utf-8")
    tmp.replace(f)  # yarımçıq yazılmış fayl qalmasın


def siyahi(login):
    with _KILID:
        return [{"id": s["id"], "ad": s.get("ad", ""), "vaxt": s.get("vaxt", 0)} for s in _oxu(login)["sohbetler"]]


def al(login, sid):
    with _KILID:
        return next((s for s in _oxu(login)["sohbetler"] if s["id"] == sid), None)


def saxla(login, sid, ad, mesajlar):
    if not ID_RE.fullmatch(sid or ""):
        raise ValueError("Söhbət nömrəsi düzgün deyil")
    s = {"id": sid, "ad": (ad or "")[:80], "vaxt": time.time(), "mesajlar": mesajlar}
    if len(json.dumps(s, ensure_ascii=False).encode()) > MAKS_OLCU:
        raise ValueError("Söhbət çox böyükdür — yeni söhbət açın")
    with _KILID:
        d = _oxu(login)
        d["login"] = login
        d["sohbetler"] = [s] + [x for x in d["sohbetler"] if x["id"] != sid]
        d["sohbetler"] = sorted(d["sohbetler"], key=lambda x: -x.get("vaxt", 0))[:MAKS_SOHBET]
        _yaz(login, d)


def sil(login, sid):
    with _KILID:
        d = _oxu(login)
        d["sohbetler"] = [x for x in d["sohbetler"] if x["id"] != sid]
        _yaz(login, d)
