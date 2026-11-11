"""İstifadəçi profili: işçinin məşğul olduğu peşələr (hər istifadəçinin öz faylında, WA_DATA/profil/).

Girişdən sonra peşələr seçilməyibsə, interfeys soruşur; cavab yadda qalır və istənilən kompüterdən keçərlidir.
Seçilmiş peşələrə uyğun modullar yan paneldə öndə göstərilir, boş ekrandakı nümunə sorğular da onlardan gəlir.
"""
import hashlib
import json
import threading
import time

from .state import DATA

QOVLUQ = DATA / "profil"
_KILID = threading.Lock()

# Peşə -> uyğun modulların kodları (assistant/modullar)
PESELER = [
    {"kod": "muhasib", "ad": "Mühasib", "tesvir": "bank daxilolmaları, uzlaşma, 1C", "modullar": ["bank", "muhasibat", "excel"]},
    {"kod": "maliyye", "ad": "Maliyyə / iqtisadçı", "tesvir": "hesabatlar, Excel təhlili", "modullar": ["excel", "muhasibat"]},
    {"kod": "huquq", "ad": "Hüquqşünas", "tesvir": "müqavilələr, müddətlər", "modullar": ["muqavile"]},
    {"kod": "karguzar", "ad": "Kargüzar", "tesvir": "məktublar, sənəd dövriyyəsi", "modullar": ["senedler"]},
    {"kod": "kadr", "ad": "Kadr mütəxəssisi", "tesvir": "məzuniyyət, əmrlər, tabel", "modullar": ["kadr"]},
    {"kod": "rehber", "ad": "Rəhbər", "tesvir": "xülasələr, nəzarət", "modullar": ["bank", "muhasibat", "excel"]},
]
KODLAR = {p["kod"] for p in PESELER}
MAKS_DIGER = 80


def _fayl(login):
    ad = hashlib.sha256((login or "lokal").encode()).hexdigest()[:20]  # fayl adında e-poçt görünməsin
    return QOVLUQ / f"{ad}.json"


def al(login):
    """Seçilmiş peşələr: {"peseler": [...], "diger": "..."}; hələ seçilməyibsə None."""
    with _KILID:
        try:
            d = json.loads(_fayl(login).read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return None
    return {"peseler": [k for k in d.get("peseler", []) if k in KODLAR], "diger": d.get("diger", "")}


def saxla(login, peseler, diger=""):
    peseler = [k for k in dict.fromkeys(peseler or []) if k in KODLAR]
    diger = " ".join(str(diger or "").split())[:MAKS_DIGER]
    if not peseler and not diger:
        raise ValueError("Ən azı bir peşə seçin və ya «Digər» sahəsinə yazın.")
    d = {"login": login, "peseler": peseler, "diger": diger, "vaxt": time.time()}
    with _KILID:
        QOVLUQ.mkdir(parents=True, exist_ok=True)
        f = _fayl(login)
        tmp = f.with_suffix(".tmp")
        tmp.write_text(json.dumps(d, ensure_ascii=False), encoding="utf-8")
        tmp.replace(f)
    return {"peseler": peseler, "diger": diger}


def modullar(profil):
    """Profilə uyğun modul kodları (seçim sırası ilə)."""
    out = []
    for k in (profil or {}).get("peseler", []):
        for m in next(p for p in PESELER if p["kod"] == k)["modullar"]:
            if m not in out:
                out.append(m)
    return out
