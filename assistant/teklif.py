"""Yüklənmiş fayla görə hazır sorğular (prompt-lar): istifadəçi faylın içində nə soruşa biləcəyini görür.

Hər sorğu faylın öz məlumatından qurulur (sütun adları, kontragentlər, hesablar, illər) və əlavə olunmadan
əvvəl lokal qaydalarla yoxlanılır: sistemin cavab verə bilmədiyi sorğu təklif olunmur.
"""
import re

from .modullar import bank as BANK
from .modullar import excel as EX

MAKS = 8


def _bank(s):
    b = s.bank
    ss = b["setirler"]
    say = {k["kod"]: sum(1 for r in ss if r["kateqoriya"] == k["kod"]) for k in BANK.KATEQORIYALAR + [BANK.YOXLANMALI]}
    out = ["Bank daxilolmalarının xülasəsi"]
    sual = len(BANK._sual_setirleri(b))
    if sual:
        out.insert(0, f"Sualları ver ({sual})")
    if say["yoxlanmali"]:
        out.append(f"Yoxlanmalıları göstər ({say['yoxlanmali']})")
    if b["xeberdarliqlar"]:
        out.append(f"Xəbərdarlıqları göstər ({len(b['xeberdarliqlar'])})")
    out.append("Şablonu doldur" if getattr(s, "sablon", None) else "1C üçün Excel hazırla")
    musz = sum(1 for r in ss if r["kateqoriya"] == "sigorta_haqqi" and not r.get("muqavile"))
    if musz:
        out.append(f"Müqaviləsizləri də şablona sal ({musz})")
    for k in BANK.KATEQORIYALAR:
        if say[k["kod"]]:
            out.append(f"{k['ad']} daxilolmalarını göstər")
    return out


def _uzlasma(s):
    from uzlasma import report as R
    out = ["Vəziyyət necədir?", "Yüksək riskli istisnalar", "Təsdiq gözləyən bağlantılar", "Debitor borclar"]
    try:
        b = R.borclar(s.m, s.dovr_sonu)
        if len(b):
            ad = dict(zip(s.k["voen"], s.k["ad"]))
            top = b.groupby("voen")["qaliq"].sum().sort_values(ascending=False)
            if len(top):
                kontr = ad.get(top.index[0], top.index[0])
                out.append(f"{kontr} üçün borc məktubu")
    except Exception:
        pass
    out.append("Excel hesabatı hazırla")
    return out


def _qrup_sutunu(s, v, maks_ferq=60):
    """Qruplaşdırmaq üçün mətn sütunu: 2..maks_ferq fərqli dəyəri olan, ən doluu."""
    en = None
    for c in EX._sutun_sec(v, maks=10):
        ad = v["sutunlar"][v["sql"].index(c)] if c in v["sql"] else c
        if v["tipler"].get(ad) != "mətn" or EX.TEXNIKI_REQEM.search(EX.norm(ad)):
            continue
        r = EX._sorgu(s, f"SELECT count(DISTINCT {c}), count({c}) FROM (SELECT {c} FROM {v['cedvel']} LIMIT 20000)").iloc[0]
        ferq, dolu = int(r.iloc[0] or 0), int(r.iloc[1] or 0)
        if 2 <= ferq <= maks_ferq and dolu and (en is None or dolu > en[1]):
            en = (ad, dolu, c)
    return en


def _excel(s):
    vv = s.excel["vereqler"]
    out = []
    j = EX._jurnal(s)
    if j:  # mühasibat yazılışları var (1C): xərc/gəlir sualları
        _, v, *_ = j
        out += ["Şirkətin ümumi xərci nə qədərdir?", "Şirkətin ümumi gəliri nə qədərdir?"]
        tarix = next((x for x in v["sutunlar"] if v["tipler"].get(x) == "tarix"), None)
        if tarix:
            il = EX._sorgu(s, f"SELECT max(substr({v['sql'][v['sutunlar'].index(tarix)]}, 1, 4)) FROM {v['cedvel']}").iloc[0, 0]
            if il and re.fullmatch(r"20\d\d", str(il)):
                out.append(f"{il}-ci ildə ümumi xərc")
        adlar = " ".join(EX._hesab_adlari(s).values()).lower()
        if "ekspert" in adlar:
            out.append("Ekspert xərcləri nə qədərdir?")

    # əsas vərəq: məbləğ sütunu olan ən böyük vərəq (1C-də texniki «-yekun» vərəqləri olmadan)
    namized = sorted((a for a in vv if EX._mebleg_sutunu(vv[a]) and not re.search(r"yekun|subk|›", a)),
                     key=lambda a: -vv[a]["setir"])
    if namized and not j:
        ad = namized[0]
        v = vv[ad]
        c = EX._mebleg_sutunu(v)
        q = _qrup_sutunu(s, v)
        if q:
            out.append(f"{c} cəmi {q[0]} üzrə")
            top = EX._sorgu(s, f"SELECT {q[2]} AS g, sum({v['sql'][v['sutunlar'].index(c)]}) AS m FROM {v['cedvel']} "
                               f"WHERE {q[2]} IS NOT NULL GROUP BY g ORDER BY m DESC LIMIT 1")
            if len(top) and str(top.iloc[0, 0]).strip():
                out.append(f"{top.iloc[0, 0]} üzrə {c} nə qədərdir?")
        out.append(f"Ən böyük 10 {c}")
    elif namized and j:
        # 1C: mövzu vərəqlərindən biri (avans, ödəniş, maaş...)
        for movzu, sorgu in [("avans", "Ümumi avans məbləği"), ("odenis tapsir", "Ümumi ödəniş cəmi"), ("emek haqq", "Ümumi əmək haqqı cəmi")]:
            if any(movzu in EX.norm(a) for a in namized):
                out.append(sorgu)
                break
    out.append("Fayl haqqında ümumi məlumat ver")
    if len(vv) <= 10:  # çox vərəqli faylda (1C) "hansı vərəq?" deyə soruşur: təklif etmirik
        out.append("Fayldakı problemləri tap")
    return out


def _yoxla(s, sorgu):
    """Sorğu lokal qaydalarla başa düşülürmü (cavabsız təklif olmasın)."""
    try:
        if getattr(s, "bank", None):
            return BANK.qayda(s, sorgu) is not None
        if s.excel:
            return EX.qayda(s, sorgu) is not None
    except Exception:
        return False
    return True


def hazirla(s):
    """Yüklənmiş fayla görə 4–8 hazır sorğu."""
    try:
        if getattr(s, "bank", None):
            out = _bank(s)
        elif s.hazir:
            out = _uzlasma(s)
        elif s.excel:
            out = _excel(s)
        else:
            return []
    except Exception:
        return []
    niyyet = getattr(s, "son_niyyet", None)
    netice = []
    for x in out:
        x = re.sub(r"\s+", " ", str(x)).strip()
        if x and x not in netice and (s.hazir and not getattr(s, "bank", None) or _yoxla(s, re.sub(r" \(\d+\)$", "", x))):
            netice.append(x)
    try:  # yoxlama zamanı qaydalar "gözləyən niyyət" yarada bilər: əvvəlki vəziyyəti qaytar
        s.son_niyyet = niyyet
    except AttributeError:
        pass
    return netice[:MAKS]
