"""Mühasibat modulu: bank, ƏDV depozit və e-qaimə uzlaşması, borclar, məktublar.

Hər alət {"metn": ..., "ui": ...} qaytarır:
  metn - AI modelinə gedən qısa mətn (minimal məlumat);
  ui   - interfeysdə göstərilən cədvəl, düymə, məktub və ya fayl.
Rəqəmləri yalnız mühərrik hesablayır; AI onları uydurmur, yalnız izah edir.
"""
import datetime as dt
from collections import Counter

from uzlasma import config as C
from uzlasma import report as R
from uzlasma.config import az

from ..state import HESABATLAR

TR = str.maketrans("əüöığşçı", "euoigsci")


def norm(s):
    return str(s).replace("İ", "i").replace("I", "ı").lower().translate(TR)


def _kontragentler(s, sorgu):
    """Ada (hissəsinə) və ya VÖEN-ə görə kontragent tapır."""
    if not sorgu:
        return []
    n = norm(sorgu)
    return [(r.voen, r.ad) for r in s.k.itertuples() if n == r.voen or n in norm(r.ad)]


def _borc_df(s):
    return R.borclar(s.m, s.dovr_sonu)


# ---------------- alətlər ----------------

def xulase(s):
    b = _borc_df(s)
    bagli = len({x["odenis_id"] for x in s.m.baglantilar})
    cem = len(s.m.o)
    gozleyen = sum(1 for x in s.m.baglantilar if not s.tesdiq(x["odenis_id"], x["qaime_id"]))
    deb = float(b[b["nov"] == "Satış"]["qaliq"].sum()) if len(b) else 0.0
    kre = float(b[b["nov"] == "Alış"]["qaliq"].sum()) if len(b) else 0.0
    yuksek = sum(1 for i in s.ist if i["risk"] == "Yüksək")
    novler = Counter(i["nov"] for i in s.ist).most_common(5)
    metn = (f"Dövr sonu {s.dovr_sonu:%d.%m.%Y}. Ödəniş sətirləri: {cem}, avtomatik bağlanan: {bagli} "
            f"({bagli / cem:.0%}). Təsdiq gözləyən bağlantı: {gozleyen}. İstisnalar: {len(s.ist)} "
            f"(yüksək riskli: {yuksek}). Debitor borc: {az(deb)} AZN, kreditor borc: {az(kre)} AZN. "
            f"Ən çox istisna növləri: " + "; ".join(f"{n} ({c})" for n, c in novler))
    ui = {"tip": "kartlar", "kartlar": [
        {"ad": "Avtomatik bağlanıb", "deyer": f"{bagli}/{cem}", "alt": f"{bagli / cem:.0%}"},
        {"ad": "Təsdiq gözləyir", "deyer": str(gozleyen)},
        {"ad": "İstisnalar", "deyer": str(len(s.ist)), "alt": f"{yuksek} yüksək risk"},
        {"ad": "Debitor borc", "deyer": f"{az(deb)} ₼"},
        {"ad": "Kreditor borc", "deyer": f"{az(kre)} ₼"},
    ]}
    return {"metn": metn, "ui": ui}


def istisnalar(s, risk=None, nov=None, kontragent=None, limit=20):
    ad = dict(zip(s.k["voen"], s.k["ad"]))
    voenler = {v for v, _ in _kontragentler(s, kontragent)} if kontragent else None
    sec = [i for i in s.ist
           if (not risk or norm(i["risk"]) == norm(risk))
           and (not nov or norm(nov) in norm(i["nov"]))
           and (voenler is None or i["voen"] in voenler)]
    sec.sort(key=lambda i: R.RISK_SIRA.get(i["risk"], 9))
    setirler = [[i["risk"], i["nov"], i["istinad"], ad.get(i["voen"], i["voen"]), i["izah"]] for i in sec[:limit]]
    metn = f"{len(sec)} istisna tapıldı" + (f", ilk {limit} göstərilir" if len(sec) > limit else "") + ":\n" + \
        "\n".join(f"- [{r[0]}] {r[1]} | {r[2]} | {r[4]}" for r in setirler)
    return {"metn": metn, "ui": {"tip": "cedvel", "basliq": "İstisnalar",
                                 "sutunlar": ["Risk", "Növ", "İstinad", "Kontragent", "İzah"], "setirler": setirler}}


def baglantilar(s, status="gozleyen", etibar=None, kontragent=None, limit=30):
    voenler = {v for v, _ in _kontragentler(s, kontragent)} if kontragent else None
    sec = []
    for x in s.m.baglantilar:
        t = s.tesdiq(x["odenis_id"], x["qaime_id"])
        if status == "gozleyen" and t:
            continue
        if etibar and norm(x["etibar"]) != norm(etibar):
            continue
        if voenler is not None and x["voen"] not in voenler:
            continue
        sec.append({**x, "tesdiq": t})
    sec.sort(key=lambda x: (x["etibar"] != "Orta", x["odenis_tarixi"]))  # orta etibarlılar əvvəl
    setirler = [{"odenis_id": x["odenis_id"], "qaime_id": x["qaime_id"], "tarix": f"{x['odenis_tarixi']:%d.%m.%Y}",
                 "axin": "Bank" if x["axin"] == "bank" else "ƏDV depozit", "kontragent": x["kontragent"],
                 "odenis": az(x["odenis_meblegi"]), "baglanan": az(x["baglanan_mebleg"]), "metn": x["metn"],
                 "qayda": x["qayda"], "etibar": x["etibar"], "tesdiq": x["tesdiq"]} for x in sec[:limit]]
    metn = f"{len(sec)} bağlantı" + (" təsdiq gözləyir" if status == "gozleyen" else "") + ":\n" + "\n".join(
        f"- {r['odenis_id']}→{r['qaime_id']} {r['kontragent']} {r['baglanan']} AZN ({r['qayda']}, etibar {r['etibar']})"
        for r in setirler)
    return {"metn": metn, "ui": {"tip": "baglantilar", "setirler": setirler}}


def baglantini_tesdiqle(s, odenis_id, qaime_id, qerar):
    s.tesdiq_et(odenis_id, qaime_id, qerar)
    return {"metn": f"{odenis_id} → {qaime_id}: '{qerar}' qeyd edildi.", "ui": None}


def borclar(s, nov=None, kontragent=None):
    b = _borc_df(s)
    if not len(b):
        return {"metn": "Açıq borc yoxdur.", "ui": None}
    if nov:
        b = b[b["nov"] == ("Satış" if norm(nov).startswith("deb") or norm(nov) == "satis" else "Alış")]
    if kontragent:
        voenler = {v for v, _ in _kontragentler(s, kontragent)}
        b = b[b["voen"].isin(voenler)]
    b = b.sort_values(["nov", "qaliq"], ascending=[False, False])
    setirler = [["Debitor" if r.nov == "Satış" else "Kreditor", r.ad, r.qaime, f"{r.tarix:%d.%m.%Y}", az(r.cem),
                 az(r.odenilib), az(r.qaliq), (s.dovr_sonu - r.tarix).days] for r in b.itertuples()]
    cemler = b.groupby("ad")["qaliq"].sum().sort_values(ascending=False)
    metn = f"Cəmi qalıq: {az(b['qaliq'].sum())} AZN. Kontragentlər üzrə: " + "; ".join(
        f"{a}: {az(v)}" for a, v in cemler.items())
    return {"metn": metn, "ui": {"tip": "cedvel", "basliq": "Açıq borclar",
                                 "sutunlar": ["Növ", "Kontragent", "Qaimə", "Tarix", "Cəmi", "Ödənilib", "Qalıq", "Gün"],
                                 "setirler": setirler}}


def qaime_axtar(s, sorgu):
    n = norm(sorgu).replace(" ", "").upper()
    q = s.m.q
    sec = q[(q["id"].str.upper() == n)]
    if sec.empty:
        voenler = {v for v, _ in _kontragentler(s, sorgu)}
        sec = q[q["voen"].isin(voenler)]
    if sec.empty:
        return {"metn": f"'{sorgu}' üzrə qaimə tapılmadı.", "ui": None}
    setirler, setir_metn = [], []
    for r in sec.itertuples():
        legv = r.legv
        qn = 0 if legv else s.m.qalan.get(("bank", r.id), 0)
        qe = 0 if legv else s.m.qalan.get(("depozit", r.id), 0)
        odenisler = [f"{b['odenis_id']} {az(b['baglanan_mebleg'])}" for b in s.m.baglantilar if b["qaime_id"] == r.id]
        status = "Ləğv edilib" if legv else ("Dublikat" if r.dublikat else ("Bağlanıb" if qn + qe <= C.DOZUM else "Açıq"))
        setirler.append([r.id, r.nov, f"{r.tarix:%d.%m.%Y}", r.ad, az(r.netto), az(r.edv), az(max(qn, 0) + max(qe, 0)),
                         status, ", ".join(odenisler) or "—"])
        setir_metn.append(f"- {r.id} {r.nov} {r.ad} {r.tarix:%d.%m.%Y}: ƏDV-siz {az(r.netto)}, ƏDV {az(r.edv)}, "
                          f"qalıq {az(max(qn, 0) + max(qe, 0))}, status {status}")
    return {"metn": "\n".join(setir_metn), "ui": {"tip": "cedvel", "basliq": "Qaimələr",
            "sutunlar": ["Qaimə", "Növ", "Tarix", "Kontragent", "ƏDV-siz", "ƏDV", "Qalıq", "Status", "Ödənişlər"],
            "setirler": setirler}}


def borc_mektubu(s, kontragent):
    tap = _kontragentler(s, kontragent)
    if len(tap) != 1:
        adlar = ", ".join(a for _, a in tap) or "heç biri"
        return {"metn": f"Kontragent dəqiq müəyyən edilmədi ({adlar}). Dəqiq adı soruşun.", "ui": None}
    voen, adi = tap[0]
    b = _borc_df(s)
    g = b[(b["voen"] == voen) & (b["nov"] == "Satış")] if len(b) else b
    if not len(g):
        return {"metn": f"{adi} üzrə debitor borc yoxdur, məktuba ehtiyac yoxdur.", "ui": None}
    maddeler = "\n".join(f"  - {r.qaime}, {r.tarix:%d.%m.%Y}: qaimə {az(r.cem)} AZN, ödənilib {az(r.odenilib)} AZN, "
                         f"qalıq {az(r.qaliq)} AZN" for r in g.itertuples())
    movzu = f"{C.SIRKET_ADI}: {s.dovr_sonu:%d.%m.%Y} tarixinə borc üzləşdirməsi"
    metn = (f"Hörmətli {adi} mühasibatlığı,\n\n{s.dovr_sonu:%d.%m.%Y} tarixinə {C.SIRKET_ADI} qarşısında borcunuz "
            f"{az(g['qaliq'].sum())} AZN təşkil edir:\n{maddeler}\n\nMəlumatlar Sizin uçotunuzla uyğun gəlmirsə, "
            f"{C.CAVAB_MUDDETI_IS_GUNU} iş günü ərzində bizə bildirməyinizi xahiş edirik.\n\nHörmətlə,\n{C.SIRKET_ADI} mühasibatlığı")
    email = dict(zip(s.k["voen"], s.k["email"])).get(voen, "")
    return {"metn": f"{adi} üçün məktub qaralaması hazırlandı (borc {az(g['qaliq'].sum())} AZN). Göndərilməyib.",
            "ui": {"tip": "mektub", "kime": email, "movzu": movzu, "metn": metn}}


def hesabat_yarat(s):
    ad = f"uzlasma_{s.dovr_sonu:%Y-%m}_{dt.datetime.now():%H%M%S}.xlsx"
    R.yaz(str(HESABATLAR / ad), s.m, s.ist, s.o, s.k, s.dovr_sonu, tesdiq=s.tesdiq)
    return {"metn": f"Excel hesabatı hazırdır: {ad}", "ui": {"tip": "fayl", "ad": ad, "url": f"/api/fayl/{ad}"}}


# ---------------- sxemlər ----------------
_kontr = {"type": "string", "description": "Kontragentin adı (hissəsi) və ya VÖEN"}
ALETLER = {
    "xulase": (xulase, "Uzlaşmanın ümumi vəziyyəti: bağlanma faizi, istisnalar, borclar.", {}),
    "istisnalar": (istisnalar, "Mühasibin baxmalı olduğu uyğunsuzluqlar. Riskə, növə və ya kontragentə görə süzülə bilər.", {
        "risk": {"type": "string", "enum": ["Yüksək", "Orta", "Aşağı", "Məlumat"]},
        "nov": {"type": "string", "description": "İstisna növünün hissəsi, məs. 'ƏDV', 'müqavilə', 'ödənilməyib'"},
        "kontragent": _kontr, "limit": {"type": "integer"}}),
    "baglantilar": (baglantilar, "Sistemin təklif etdiyi ödəniş→qaimə bağlantıları. Mühasib onları interfeysdə təsdiqləyir.", {
        "status": {"type": "string", "enum": ["gozleyen", "hamisi"]},
        "etibar": {"type": "string", "enum": ["Yüksək", "Orta"]}, "kontragent": _kontr, "limit": {"type": "integer"}}),
    "baglantini_tesdiqle": (baglantini_tesdiqle, "Bağlantını təsdiqləyir və ya rədd edir. YALNIZ mühasib konkret bağlantı üçün "
                            "açıq şəkildə 'təsdiqlə' və ya 'rədd et' dedikdə çağır.", {
        "odenis_id": {"type": "string"}, "qaime_id": {"type": "string"},
        "qerar": {"type": "string", "enum": ["Bəli", "Xeyr"]}}),
    "borclar": (borclar, "Dövrün sonuna açıq debitor (bizə borclu) və kreditor (bizim borcumuz) qalıqları.", {
        "nov": {"type": "string", "enum": ["Debitor", "Kreditor"]}, "kontragent": _kontr}),
    "qaime_axtar": (qaime_axtar, "Qaiməni nömrəyə (məs. AA1005), VÖEN-ə və ya kontragent adına görə tapır.", {
        "sorgu": {"type": "string"}}),
    "borc_mektubu": (borc_mektubu, "Kontragentə borc üzləşdirmə məktubunun qaralamasını hazırlayır. Göndərmir.", {
        "kontragent": _kontr}),
    "hesabat_yarat": (hesabat_yarat, "Bütün nəticələrlə Excel hesabatı yaradır və yükləmə linki verir.", {}),
}
MECBURI = {"baglantini_tesdiqle": ["odenis_id", "qaime_id", "qerar"], "qaime_axtar": ["sorgu"], "borc_mektubu": ["kontragent"]}


MODUL = {
    "kod": "muhasibat",
    "ad": "Mühasibat",
    "tesvir": "Bank, ƏDV depozit və e-qaimə uzlaşması, açıq borclar, borc məktubları, Excel hesabatı",
    "status": "aktiv",
    "teklifler": [
        {"s": "Vəziyyət necədir?", "b": "Ümumi vəziyyət", "a": "bağlanma faizi, istisnalar, borclar"},
        {"s": "Yüksək riskli istisnaları göstər", "b": "Yüksək riskli istisnalar", "a": "ƏDV, qaiməsiz mədaxil və s."},
        {"s": "Təsdiq gözləyən bağlantıları göstər", "b": "Təsdiq gözləyən bağlantılar", "a": "ödəniş → qaimə təklifləri"},
        {"s": "Debitor borcları göstər", "b": "Debitor borclar", "a": "bizə kim nə qədər borcludur"},
    ],
}
