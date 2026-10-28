"""Uyğunlaşdırmadan sonra yoxlamalar: ƏDV, açıq qaimələr, dövr fərqi, dublikatlar və müqavilələr."""
import datetime as dt

from . import config as C
from .config import az

TOL = C.DOZUM


def yoxla(m, qaimeler, odenisler, muqavileler, dovr_sonu):
    ist = m.istisnalar

    def elave(nov, risk, istinad, voen, izah):
        ist.append({"nov": nov, "risk": risk, "istinad": istinad, "voen": voen, "izah": izah})

    # Dublikatlar
    for _, r in qaimeler[qaimeler["dublikat"]].iterrows():
        elave("Dublikat qaimə", "Orta", r["id"], r["voen"], f"{r['ad']}: {r['id']} siyahıda təkrar qeyd olunub.")
    for _, r in odenisler[odenisler["dublikat"]].iterrows():
        yer = "Bank" if r["axin"] == "bank" else "ƏDV depozit"
        elave("Ödəniş sətri təkrar", "Orta", f"{yer} {r['tarix']:%d.%m.%Y}", r["voen"],
              f"{r['ad']}: {az(r['mebleg'])} AZN ({r['metn']}) çıxarışda iki dəfə görünür. Təkrar sətir nəzərə alınmayıb.")

    aktiv = m.q[~m.q["legv"]]
    for _, r in aktiv.iterrows():
        qid, ad = r["id"], r["ad"]
        # ƏDV dərəcəsi
        gozlenen = round(r["netto"] * C.EDV_DERECESI, 2)
        if abs(r["edv"] - gozlenen) > TOL:
            elave("ƏDV hesablanması səhvdir", "Yüksək", qid, r["voen"],
                  f"{ad}: qaimədə ƏDV {az(r['edv'])} AZN ({r['edv'] / r['netto']:.0%}), {C.EDV_DERECESI:.0%} ilə {az(gozlenen)} olmalıdır.")
        qn = m.qalan[("bank", qid)]
        qe = m.qalan[("depozit", qid)]
        if qn > TOL and qe > TOL and abs(qn - r["netto"]) <= TOL and abs(qe - r["edv"]) <= TOL:
            gun = (dovr_sonu - r["tarix"]).days
            risk = "Yüksək" if gun > C.BORC_GECIKME_GUN else "Orta"
            nov = "Satış qaiməsi ödənilməyib" if r["nov"] == "Satış" else "Alış qaiməsi ödənilməyib"
            elave(nov, risk, qid, r["voen"], f"{ad}: {az(r['netto'] + r['edv'])} AZN, {gun} gündür açıqdır.")
        elif qn > TOL:
            elave("Qismən ödəniş", "Orta", qid, r["voen"],
                  f"{ad}: ƏDV-siz hissədən {az(qn)} AZN, ƏDV-dən {az(max(qe, 0))} AZN qalıq var.")
        elif qe > TOL:
            if r["nov"] == "Satış":
                elave("ƏDV depozitə gəlməyib", "Yüksək", qid, r["voen"],
                      f"{ad}: ƏDV-siz hissə tam ödənilib, ƏDV depozitə {az(qe)} AZN gəlməyib.")
            else:
                elave("Alış ƏDV-si depozitdən ödənilməyib", "Yüksək", qid, r["voen"],
                      f"{ad}: ƏDV-siz hissə ödənilib, ƏDV {az(qe)} AZN depozitdən köçürülməyib. Əvəzləşdirmə riski.")

    # Dövr fərqi
    for b in m.baglantilar:
        if b["odenis_tarixi"] > dovr_sonu:
            elave("Dövr fərqi", "Aşağı", b["qaime_id"], b["voen"],
                  f"{b['kontragent']}: ödəniş {b['odenis_tarixi']:%d.%m.%Y} tarixində ({b['axin']}), hesabat dövründən sonra.")

    # Müqavilələr
    for _, r in aktiv.iterrows():
        mq = muqavileler[(muqavileler["voen"] == r["voen"]) & (muqavileler["nov"] == r["nov"])]
        if mq.empty:
            elave("Müqaviləsiz kontragent", "Orta", r["id"], r["voen"], f"{r['ad']}: {r['nov'].lower()} müqaviləsi tapılmadı.")
            continue
        aktiv_mq = mq[(mq["baslama"] <= r["tarix"]) & (mq["bitme"] >= r["tarix"])]
        if aktiv_mq.empty:
            son = mq.sort_values("bitme").iloc[-1]
            elave("Bitmiş müqavilə üzrə qaimə", "Orta", r["id"], r["voen"],
                  f"{r['ad']}: {son['nomre']} {son['bitme']:%d.%m.%Y} tarixində bitib, qaimə {r['tarix']:%d.%m.%Y}.")
    for _, c in muqavileler.iterrows():
        q = aktiv[(aktiv["voen"] == c["voen"]) & (aktiv["nov"] == c["nov"]) &
                  (aktiv["tarix"] >= c["baslama"]) & (aktiv["tarix"] <= c["bitme"])]
        cem = c["evvelki"] + q["netto"].sum()
        if cem > c["mebleg"] + TOL:
            elave("Müqavilə məbləği aşılıb", "Orta", ", ".join(q["id"]), c["voen"],
                  f"{c['ad']}: {c['nomre']} limiti {az(c['mebleg'])}, qaimələrin cəmi {az(cem)} AZN.")
        qalan_gun = (c["bitme"] - dovr_sonu).days
        if 0 <= qalan_gun <= C.MUQAVILE_XEBERDARLIQ_GUN:
            elave("Müddəti bitmək üzrə müqavilə", "Məlumat", c["nomre"], c["voen"],
                  f"{c['ad']}: {c['bitme']:%d.%m.%Y} tarixində bitir ({qalan_gun} gün qalıb).")
    return ist
