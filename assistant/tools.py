"""Bütün modulların alətləri üçün vahid giriş nöqtəsi.

Hər alət {"metn": ..., "ui": ...} qaytarır:
  metn - AI modelinə gedən qısa mətn (minimal məlumat);
  ui   - interfeysdə göstərilən cədvəl, düymə, məktub və ya fayl.
Rəqəmləri yalnız mühərriklər hesablayır; AI onları uydurmur, yalnız izah edir.
"""
from .modullar import ALETLER, MECBURI, SAHIB, modullar  # noqa: F401
from .modullar.muhasibat import norm, xulase  # noqa: F401


def sxem(ad):
    _, tesvir, xususiyyetler = ALETLER[ad]
    return tesvir, {"type": "object", "properties": xususiyyetler, "required": MECBURI.get(ad, [])}


def cagir(s, ad, args):
    if ad not in ALETLER:
        return {"metn": f"Naməlum alət: {ad}", "ui": None}
    if SAHIB.get(ad) == "muhasibat" and not s.hazir:
        return {"metn": "Yüklənmiş fayl uzlaşma şablonunda deyil (qaimə, bank, ƏDV depozit vərəqləri yoxdur), "
                        "ona görə Mühasibat alətləri işləmir. Excel analizi alətlərindən istifadə et.", "ui": None}
    if SAHIB.get(ad) == "bank" and not getattr(s, "bank", None):
        return {"metn": "Yüklənmiş fayl bank çıxarışı kimi tanınmadı (tarix, mədaxil/kredit və təyinat sütunları tapılmadı).", "ui": None}
    try:
        return ALETLER[ad][0](s, **(args or {}))
    except Exception as e:  # alət xətası modelə qaytarılır ki, istifadəçiyə izah etsin
        return {"metn": f"Xəta: {e}", "ui": None}
