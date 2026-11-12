"""Diş klinikası: stok nəzarəti. Azalan məhsullar barədə klinika rəhbərini məlumatlandırır, lazım olanda
təchizatçılara WhatsApp sifariş mesajı hazırlayır (və ya göndərir).

Mənbə — stok (anbar) Excel faylı: Məhsul, Qalıq, Minimum qalıq; istəyə bağlı: Vahid, Təchizatçı,
Təchizatçının telefonu, Sifariş miqdarı. Qalığı minimumdan az və ya ona bərabər olan məhsul «azalan» sayılır,
qalığı 0 olan — «bitib».

Sifariş miqdarı: «Sifariş miqdarı» sütunu varsa ondan; yoxdursa qalığı minimumun iki qatına çatdıran miqdar.
WhatsApp göndərmə qəbul modulunun altyapısı ilə (assistant/modullar/qebul.py): API qurulubsa — istifadəçinin
təsdiqindən sonra Meta şablonu ilə; qurulmayıbsa — wa.me linki. Sistem heç nəyi özbaşına sifariş etmir.

Mühit dəyişənləri:
  WA_REHBER_TEL - klinika rəhbərinin WhatsApp nömrəsi (çatda da yazmaq olar: «Rəhbərin telefonu 050 ...»)
"""
import hashlib
import json
import math
import os
import pathlib
import re

import pandas as pd

from . import qebul as QB
from . import teleb as TELEB
from .excel import _sorgu, norm

SABLONLAR = {
    "stok_rehber": {
        "kime": "Rəhbər", "ad": "Rəhbərə stok bildirişi",
        "metn": "Salam! Klinikanın stokunda {{1}} məhsul minimum həddən azdır: {{2}}. Təchizatçılara sifariş göndərmək "
                "üçün Worker's Assistant-da «Təchizatçılara sifariş hazırla» seçin.",
        "parametrler": ["say", "siyahi"]},
    "stok_sifaris": {
        "kime": "Təchizatçı", "ad": "Təchizatçıya sifariş",
        "metn": "Salam, {{1}}! Diş klinikası üçün sifariş vermək istəyirik: {{2}}. Xahiş edirik, mövcudluğu, qiyməti "
                "və çatdırılma vaxtını bildirin. Təşəkkür edirik.",
        "parametrler": ["techizatci", "siyahi"]},
}

ROLLAR = [  # sıra vacibdir: «Minimum qalıq» qalıqdan, «Sifariş miqdarı» miqdardan, telefon təchizatçıdan əvvəl
    ("tel", r"telefon|\btel\b|mobil|whatsapp|elaqe"),
    ("minimum", r"minimum|\bmin\b|kritik|hedd"),
    ("sifaris", r"sifaris"),
    ("techizatci", r"techizatci|tedarukcu|satici|postavsik|firma"),
    ("vahid", r"vahid|olcu"),
    ("qaliq", r"qaliq|stok|miqdar|movcud|anbarda|say$"),
    ("mehsul", r"mehsul|material|derman|\bmal\b|\bad\b|adi$|adlar"),
]


def _rehber_fayli():
    return pathlib.Path(QB.QOVLUQ or ".") / "rehber_telefonu.json"


def rehber_tel():
    try:
        t = json.loads(_rehber_fayli().read_text(encoding="utf-8")).get("tel", "")
    except (OSError, ValueError):
        t = ""
    return QB.telefon(t or os.environ.get("WA_REHBER_TEL", ""))


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
    if not getattr(s, "excel", None):
        return None, None, None
    for ad, v in s.excel["vereqler"].items():
        r = _rollar(v)
        if {"mehsul", "qaliq", "minimum"} <= set(r):
            return ad, v, r
    return None, None, None


def _eded(x):
    try:
        f = float(str(x).replace(",", ".").replace(" ", "")) if x is not None else None
    except ValueError:
        return None
    return None if f is None or pd.isna(f) else f


def _goster(x):
    return f"{x:g}".replace(".", ",") if x is not None else "—"


def stok(s):
    ad, v, r = _vereq(s)
    if not v:
        return None
    df = _sorgu(s, f"SELECT * FROM {v['cedvel']}")
    sql = dict(zip(v["sutunlar"], v["sql"]))
    out = []
    for x in df.to_dict("records"):
        def al(rol):
            c = r.get(rol)
            val = x.get(sql[c]) if c else None
            return None if val is None or (isinstance(val, float) and pd.isna(val)) else val
        mehsul = str(al("mehsul") or "").strip()
        if not mehsul:
            continue
        qaliq, mn = _eded(al("qaliq")), _eded(al("minimum"))
        if qaliq is None or mn is None:
            veziyyet = "oxunmur"
        elif qaliq <= 0:
            veziyyet = "bitib"
        elif qaliq <= mn:
            veziyyet = "azdır"
        else:
            veziyyet = "normal"
        sif = _eded(al("sifaris"))
        if veziyyet in ("bitib", "azdır"):
            sif = sif if sif and sif > 0 else max(math.ceil(2 * mn - qaliq), 1)
        else:
            sif = None
        out.append({"setir": int(x["_setir"]), "mehsul": mehsul, "vahid": str(al("vahid") or "").strip(),
                    "qaliq": qaliq, "minimum": mn, "veziyyet": veziyyet, "sifaris": sif,
                    "techizatci": str(al("techizatci") or "").strip(), "tel": QB.telefon(al("tel"))})
    _techizatcilari_yadda_saxla(out)
    return {"vereq": ad, "mehsullar": out}


# Stok faylı yüklənəndə məhsul → təchizatçı (və telefonu) yadda saxlanır: stok faylı açıq olmayanda da
# həkimlərin təsdiqlənmiş tələbləri üzrə sifariş hazırlamaq olsun
def _techizatci_fayli():
    return pathlib.Path(TELEB.QOVLUQ or ".") / "techizatcilar.json"


def _techizatcilari_yadda_saxla(mehsullar):
    yeni = {norm(m["mehsul"]).strip(): {"mehsul": m["mehsul"], "techizatci": m["techizatci"], "tel": m["tel"], "vahid": m["vahid"]}
            for m in mehsullar if m.get("techizatci")}
    if not yeni:
        return
    try:
        d = json.loads(_techizatci_fayli().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        d = {}
    if any(d.get(k) != v for k, v in yeni.items()):
        d.update(yeni)
        _techizatci_fayli().parent.mkdir(parents=True, exist_ok=True)
        _techizatci_fayli().write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")


def _yaddasdaki_mehsullar():
    try:
        d = json.loads(_techizatci_fayli().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return []
    return [dict(v, qaliq=None, minimum=None, veziyyet="", sifaris=None) for v in d.values()]


def _yoxdur(s):
    if getattr(s, "excel", None):
        return {"metn": "Bu faylda stok cədvəli tapılmadı: «Məhsul», «Qalıq» və «Minimum qalıq» sütunları olan vərəq lazımdır "
                        "(istəyə bağlı: «Vahid», «Təchizatçı», «Təchizatçının telefonu», «Sifariş miqdarı»).", "ui": None}
    return {"metn": "Stok faylı yüklənməyib. 📎 ilə anbar (stok) Excel faylını əlavə edin.", "ui": None}


def _azalanlar(d):
    return sorted([m for m in d["mehsullar"] if m["veziyyet"] in ("bitib", "azdır")],
                  key=lambda m: (m["veziyyet"] != "bitib", (m["qaliq"] or 0) / (m["minimum"] or 1)))


def _vahid(m):
    return f" {m['vahid']}" if m["vahid"] else ""


# ======================= alətlər =======================
def stok_veziyyeti(s, yalniz_azalan=False):
    d = stok(s)
    if not d:
        return _yoxdur(s)
    az = _azalanlar(d)
    ms = az if yalniz_azalan else sorted(d["mehsullar"], key=lambda m: ["bitib", "azdır", "oxunmur", "normal"].index(m["veziyyet"]))
    setirler = [[m["mehsul"], _goster(m["qaliq"]) + _vahid(m), _goster(m["minimum"]),
                 {"bitib": "⛔ bitib", "azdır": "⚠ azdır", "normal": "✓ normal", "oxunmur": "? oxunmur"}[m["veziyyet"]],
                 (_goster(m["sifaris"]) + _vahid(m)) if m["sifaris"] else "—", m["techizatci"] or "—"] for m in ms]
    bitib = sum(1 for m in az if m["veziyyet"] == "bitib")
    if az:
        metn = (f"Stokda {len(az)} məhsul minimum həddədir və ya ondan azdır" + (f", {bitib}-i bitib" if bitib else "") + ": " +
                ", ".join(m["mehsul"] for m in az[:5]) + ("…" if len(az) > 5 else "") + ".")
        variantlar = [{"etiket": "Rəhbərə bildir", "mesaj": "Rəhbərə stok bildirişi hazırla", "esas": True},
                      {"etiket": "Təchizatçılara sifariş hazırla", "mesaj": "Təchizatçılara sifariş hazırla"}]
    else:
        metn = f"Stok qaydasındadır: {len(d['mehsullar'])} məhsulun heç biri minimum həddə çatmayıb."
        variantlar = []
    ui = {"tip": "cedvel", "basliq": "Azalan məhsullar" if yalniz_azalan else "Stok",
          "sutunlar": ["Məhsul", "Qalıq", "Minimum", "Vəziyyət", "Sifariş (təklif)", "Təchizatçı"], "setirler": setirler}
    if variantlar:  # cədvəl + növbəti addım düymələri
        return {"metn": metn, "ui": {"tip": "bloklar", "bloklar": [ui, {"tip": "sual", "variantlar": variantlar}]}}
    return {"metn": metn, "ui": ui}


def _kart(sablon, ad, tel, param, gonder, acar, sablonlar=None):
    s = (sablonlar or SABLONLAR)[sablon]
    metn = re.sub(r"\{\{(\d+)\}\}", lambda m: param[int(m.group(1)) - 1], s["metn"])
    k = {"kime": s["kime"], "ad": ad, "tel": QB.tel_goster(tel), "metn": metn, "sablon": s["ad"], "_acar": acar}
    evvel = QB._gonderilenler()
    if not tel:
        k["status"] = "telefon yoxdur"
    elif acar in evvel:
        k["status"] = f"göndərilib {evvel[acar][:16].replace('T', ' ')}"
    elif gonder and QB.api_aktiv():
        try:
            QB.GONDERICI(tel, sablon, param)
            QB._qeyd_et(acar)
            k["status"] = "göndərildi ✓"
        except Exception as x:
            k["status"] = f"göndərilmədi: {x}"
    else:
        k["link"] = QB._link(tel, metn)
    return k


def _acar(sablon, tel, siyahi):
    return f"{sablon}|{tel}|{QB._indi():%Y-%m-%d}|{hashlib.sha256(siyahi.encode()).hexdigest()[:10]}"  # eyni gün eyni siyahı bir dəfə


def _netice(kartlar, gonder, ad, sual_mesaj):
    if QB.api_aktiv() and gonder:
        ok = sum(1 for k in kartlar if k.get("status") == "göndərildi ✓")
        sehv = sum(1 for k in kartlar if str(k.get("status", "")).startswith("göndərilmədi"))
        return {"metn": f"{ad}: {ok} göndərildi" + (f", {sehv} göndərilmədi" if sehv else "") + ".", "ui": {"tip": "whatsapp", "kartlar": kartlar}}
    if QB.api_aktiv():
        return {"metn": f"{ad} hazırdır. Mətnə baxıb təsdiq edin — WhatsApp ilə göndərilsin?",
                "ui": {"tip": "whatsapp", "kartlar": kartlar, "sual": {"tip": "sual", "variantlar": [
                    {"etiket": "✓ Bəli, göndər", "mesaj": sual_mesaj, "esas": True},
                    {"etiket": "Ləğv et", "mesaj": "Stok mesajlarını göndərmə", "ikinci": True}]}}}
    return {"metn": f"{ad} hazırdır. WhatsApp Business API qurulmayıb: «WhatsApp-da aç» düyməsini basın — mətn hazır açılacaq, "
                    f"«Göndər»i özünüz basırsınız.", "ui": {"tip": "whatsapp", "kartlar": kartlar}}


def stok_rehbere(s, gonder=False):
    """Klinika rəhbərinə azalan məhsullar barədə WhatsApp bildirişi."""
    d = stok(s)
    if not d:
        return _yoxdur(s)
    az = _azalanlar(d)
    if not az:
        return {"metn": "Bildiriləcək bir şey yoxdur: stokda azalan məhsul yoxdur.", "ui": None}
    tel = rehber_tel()
    siyahi = "; ".join(f"{m['mehsul']} — qalıq {_goster(m['qaliq'])}{_vahid(m)} (min {_goster(m['minimum'])})" for m in az)
    k = _kart("stok_rehber", "Klinika rəhbəri", tel, [str(len(az)), siyahi], gonder, _acar("stok_rehber", tel, siyahi))
    n = _netice([k], gonder, "Rəhbərə stok bildirişi", "Rəhbərə stok bildirişini göndər")
    if not tel:
        n["metn"] = "Rəhbərin WhatsApp nömrəsi yoxdur — yazın, məs.: «Rəhbərin telefonu 050 123 45 67». Mətn aşağıdadır."
    return n


def stok_sifaris(s, gonder=False, techizatci=None):
    """Azalan məhsullar üzrə təchizatçılara (hər birinə öz məhsulları ilə) WhatsApp sifariş mesajı."""
    d = stok(s)
    faylsiz = d is None  # stok faylı açıq deyil: yalnız təsdiqlənmiş tələblər (təchizatçı yaddaşdan)
    if faylsiz:
        if not TELEB.tesdiqlenmis():
            return {"metn": "Təsdiqlənmiş stok tələbi yoxdur. Azalan məhsullar üzrə sifariş üçün stok faylını 📎 ilə yükləyin.", "ui": None}
        d = {"mehsullar": _yaddasdaki_mehsullar()}
    az = [dict(m, telebler=[]) for m in _azalanlar(d)]
    # həkimlərin təsdiqlənmiş tələbləri: stokdakı məhsulla tutuşdurulur (təchizatçı ondan), miqdar böyüyü götürülür
    uygunsuz = []
    for t in TELEB.tesdiqlenmis():
        tn = norm(t["mehsul"]).strip()
        m = next((x for x in az if tn and (tn in norm(x["mehsul"]) or norm(x["mehsul"]) in tn)), None)
        if m is None:
            st = next((x for x in d["mehsullar"] if tn and (tn in norm(x["mehsul"]) or norm(x["mehsul"]) in tn)), None)
            if st:
                m = dict(st, sifaris=0, telebler=[])
                az.append(m)
        if m is None:
            uygunsuz.append(t)
            continue
        m["sifaris"] = max(m["sifaris"] or 0, t["miqdar"] or 0) or m["sifaris"]
        m["telebler"].append(t["no"])
    az = [m for m in az if not techizatci or norm(techizatci) in norm(m["techizatci"])]
    if not az and not uygunsuz:
        return {"metn": "Sifariş lazım deyil: stokda azalan məhsul və təsdiqlənmiş tələb yoxdur.", "ui": None}
    qruplar = {}
    for m in az:
        qruplar.setdefault((m["techizatci"] or "Təchizatçı göstərilməyib", m["tel"]), []).append(m)
    for t in uygunsuz:  # stokda olmayan yeni məhsul: təchizatçısı bilinmir
        qruplar.setdefault(("Təchizatçı göstərilməyib", ""), []).append(
            {"mehsul": t["mehsul"], "sifaris": t["miqdar"], "vahid": t["vahid"], "telebler": [t["no"]]})
    kartlar, kart_telebleri = [], []
    for (ad, tel), ms in qruplar.items():
        siyahi = "; ".join(f"{m['mehsul']} — {_goster(m['sifaris']) if m['sifaris'] else 'miqdarı dəqiqləşdirin'}{_vahid(m)}" for m in ms)
        kartlar.append(_kart("stok_sifaris", ad, tel, [ad, siyahi], gonder, _acar("stok_sifaris", tel, siyahi)))
        kart_telebleri.append({no for m in ms for no in m.get("telebler", [])})
    for k, nolar in zip(kartlar, kart_telebleri):  # göndərilən sifarişin tələbləri «sifariş edildi» olur
        if k.get("status") == "göndərildi ✓" and nolar:
            TELEB.status_deyis(nolar, "sifaris")
    n = _netice(kartlar, gonder, f"{len(kartlar)} təchizatçıya sifariş", "Təchizatçılara sifarişi göndər")
    tno = sorted({no for x in kart_telebleri for no in x})
    if tno:
        n["metn"] += f" Həkimlərin təsdiqlənmiş tələbləri də daxildir: {', '.join(tno)}."
    if faylsiz:
        n["metn"] = "Stok faylı açıq deyil — sifariş həkimlərin təsdiqlənmiş tələblərindən hazırlandı. " + n["metn"]
        if uygunsuz:
            n["metn"] += (" Təchizatçısı bilinməyən məhsullar var: təchizatçılar stok faylından yadda qalır — stok faylını bir dəfə "
                          "yükləsəniz, sonrakı sifarişlərdə avtomatik tapılacaq.")
    telsiz = [k["ad"] for k in kartlar if not k["tel"] and k["ad"] != "Təchizatçı göstərilməyib"]
    if telsiz:
        n["metn"] += f" Telefonu olmayan: {', '.join(telsiz)} — stok faylında «Təchizatçının telefonu» sütununa yazın."
    return n


def rehber_telefonu(s, telefon_nomresi):
    t = QB.telefon(telefon_nomresi)
    if not t:
        return {"metn": f"«{telefon_nomresi}» nömrə kimi oxunmadı. Məs.: 050 123 45 67.", "ui": None}
    _rehber_fayli().parent.mkdir(parents=True, exist_ok=True)
    _rehber_fayli().write_text(json.dumps({"tel": t}), encoding="utf-8")
    return {"metn": f"Oldu: klinika rəhbərinin nömrəsi {QB.tel_goster(t)}. Stok bildirişləri bu nömrəyə gedəcək.", "ui": None}


# ======================= lokal (AI-sız) rejim =======================
def qayda(s, mesaj):
    n = norm(mesaj)
    m = re.search(r"rehber\w*\s+(?:telefonu|nomresi)\s+([+\d][\d\s\-()]{7,})$", n)
    if m:
        return "rehber_telefonu", {"telefon_nomresi": m.group(1)}
    if not re.search(r"stok|anbar|azalan|biten|bitib|sifaris|techizatci|rehber", n):
        return None
    if not _vereq(s)[1]:  # stok faylı yoxdur: yalnız təsdiqlənmiş tələblər üzrə sifariş mümkündür
        if re.search(r"sifaris|techizatci", n) and not re.search(r"gonderme|legv", n):
            return "stok_sifaris", {"gonder": bool(re.search(r"\bgonder\b|gonderin", n))}
        return None
    if re.search(r"gonderme|legv", n):
        return "stok_veziyyeti", {}
    gonder = bool(re.search(r"\bgonder\b|gonderin", n))
    if re.search(r"sifaris|techizatci", n):
        return "stok_sifaris", {"gonder": gonder}
    if "rehber" in n:
        return "stok_rehbere", {"gonder": gonder}
    if re.search(r"azalan|biten|bitib|az qalan", n):
        return "stok_veziyyeti", {"yalniz_azalan": True}
    return "stok_veziyyeti", {}


ALETLER = {
    "stok_veziyyeti": (stok_veziyyeti, "Diş klinikasının stoku: qalıq, minimum, vəziyyət (bitib/azdır/normal), sifariş təklifi, təchizatçı.",
                       {"yalniz_azalan": {"type": "boolean"}}),
    "stok_rehbere": (stok_rehbere, "Klinika rəhbərinə azalan məhsullar barədə WhatsApp bildirişi. gonder=true YALNIZ istifadəçi "
                     "açıq təsdiq etdikdə.", {"gonder": {"type": "boolean"}}),
    "stok_sifaris": (stok_sifaris, "Azalan məhsullar üzrə hər təchizatçıya WhatsApp sifariş mesajı. gonder=true YALNIZ istifadəçi "
                     "açıq təsdiq etdikdə.", {"gonder": {"type": "boolean"}, "techizatci": {"type": "string"}}),
    "rehber_telefonu": (rehber_telefonu, "Klinika rəhbərinin WhatsApp nömrəsini yadda saxlayır.", {"telefon_nomresi": {"type": "string"}}),
}
MECBURI = {"rehber_telefonu": ["telefon_nomresi"]}

MODUL = {
    "kod": "stok",
    "ad": "Stok nəzarəti",
    "tesvir": "Diş klinikası: azalan məhsullar, rəhbərə bildiriş, təchizatçılara WhatsApp sifarişi",
    "status": "aktiv",
    "teklifler": [
        {"s": "Stok vəziyyəti", "b": "Stok", "a": "qalıq, minimum, bitən və azalan məhsullar"},
        {"s": "Azalan məhsulları göstər", "b": "Azalan məhsullar", "a": "sifariş təklifi ilə"},
        {"s": "Rəhbərə stok bildirişi hazırla", "b": "Rəhbəri məlumatlandır", "a": "WhatsApp: nə azalıb"},
        {"s": "Təchizatçılara sifariş hazırla", "b": "Təchizatçıya sifariş", "a": "WhatsApp: hər təchizatçıya öz məhsulları"},
    ],
}
