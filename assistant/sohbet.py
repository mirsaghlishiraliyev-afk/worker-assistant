"""Şablon (gündəlik) suallara cavablar: salamlaşma, hal-əhval, "sən kimsən", təşəkkür və s.

Bu cavablar AI modelinə getmədən, istənilən rejimdə və fayl yüklənməmiş də işləyir.
Mesajda iş sualı da varsa (məs. "Salam, Alfa borcu nə qədərdir?"), şablon cavab verilmir: sual alətlərə ötürülür.
"""
import random
import re

from . import tools as T

# Hər niyyət: (nümunələr - normallaşdırılmış mətndə axtarılır, cavab funksiyası)
_SALAM = r"\b(salam|salamlar|salam aleykum|salam eleykum|hey|hi|hello|sabahiniz xeyir|sabah xeyir|axsaminiz xeyir|gunortaniz xeyir)\b"
_HAL = r"\b(necesen|necesiz|nece sen|nece siz|ne var ne yox|isler necedir|keyfin necedir|halin necedir|nə var)\b"
_KIM = r"\b(sen kimsen|siz kimsiniz|kimsen|adin nedir|adiniz nedir|sen nesen|ozunu teqdim et|ozun haqqinda)\b"
_BACARIQ = r"\b(ne ede bilersen|ne edirsen|nece komek|neye komek|ne bacarirsan|imkanlarin|funksiyalarin|kömək|komek et|help)\b"
_TESEKKUR = r"\b(tesekkur|tesekkurler|sag ol|sagol|sag olun|cox sag ol|minnetdaram|thanks|thank you|ela|super|afərin|aferin)\b"
_SAGOL = r"\b(helelik|gorusenedek|sabaha qeder|xudahafiz|bye|gorusmek uzre)\b"
_YARADAN = r"\b(seni kim yaradib|kim yaradib|kim duzeldib|hansi ai|hansi model|sen robotsan|sen sun.? intellektsen)\b"

# Bu sözlərdən başqa söz varsa, mesaj iş sualı sayılır
_DOLGU = set("""salam salamlar aleykum eleykum hey hi hello sabahiniz sabah axsaminiz gunortaniz xeyir
necesen necesiz nece sen siz ne var yox isler necedir keyfin halin kimsen kimsiniz adin adiniz nedir nesen ozunu
ozun teqdim et haqqinda ede bilersen edirsen komek neye bacarirsan imkanlarin funksiyalarin help tesekkur tesekkurler
sag ol olun sagol cox minnetdaram thanks thank you ela super aferin hellik helelik gorusenedek sabaha qeder xudahafiz
bye gorusmek uzre seni kim yaradib duzeldib hansi ai model robotsan sun intellektsen men biz de da mene bize
bir yaxsi pis eh oke ok aha he beli xeyr yaxsiyam qardas dostum xanim bey muellim""".split())


def _norm(m):
    m = T.norm(m)
    return re.sub(r"[^\w\s]", " ", m)


def _is_suali(n):
    """Mesajda şablon sözlərdən kənar məzmun varsa, True."""
    sozler = [s for s in n.split() if s]
    artiq = [s for s in sozler if s not in _DOLGU]
    return len(artiq) >= 2


def _aktiv():
    return [m for m in T.modullar() if m["status"] == "aktiv"]


def _planda():
    return [m for m in T.modullar() if m["status"] != "aktiv"]


def _fayl_ipucu(s):
    return "" if s.fayl_var else "\n\nBaşlamaq üçün 📎 düyməsi ilə Excel faylını əlavə edin."


def _kim(s, rejim):
    akt = "\n".join(f"- **{m['ad']}**: {m['tesvir']}" for m in _aktiv())
    plan = ", ".join(m["ad"] for m in _planda())
    return (f"Mən **Worker's Assistant**-am: işçilərin əl ilə gördüyü təkrarlanan işləri avtomatlaşdıran köməkçiyəm.\n\n"
            f"Hazırda aktiv modullar:\n{akt}\n\nTezliklə: {plan}.\n\n"
            f"Mən yalnız təklif edirəm, yekun qərarı siz verirsiniz." + _fayl_ipucu(s))


def _bacariq(s, rejim):
    misallar = []
    for m in _aktiv():
        misallar += [f"- «{t['s']}»" for t in m.get("teklifler", [])]
    misallar += ["- «Məbləğ cəmi müştəri üzrə»", "- «Ən böyük 10 məbləğ»", "- «'Alfa' axtar»",
                 "- «AA1005 qaiməsi»", "- «Alfa üçün borc məktubu»"]
    return "Bu cür sorğularla kömək edə bilərəm:\n" + "\n".join(misallar) + _fayl_ipucu(s)


def _ad(s):
    """Girişdə tanınan istifadəçinin adı (məs. «Mirsaghli Shiraliyev» -> «Mirsaghli»)."""
    ad = (getattr(s, "istifadeci_adi", "") or "").split()
    return ad[0] if ad and "@" not in ad[0] else ""


def _salam(s, rejim):
    ad = _ad(s)
    giris = random.choice([f"Salam, {ad}! 👋", f"Salam, {ad}, xoş gördük!", f"Salam, {ad}!"] if ad else
                          ["Salam! 👋", "Salam, xoş gördük!", "Salam!"])
    if s.fayl_var:
        return f"{giris} «{s.excel['ad']}» faylı yüklənib. Ümumi vəziyyətə baxaq, yoxsa konkret sualınız var?"
    return f"{giris} Mən Worker's Assistant-am. Bu gün hansı işi həll edək?" + _fayl_ipucu(s)


def _hal(s, rejim):
    return random.choice([
        "Yaxşıyam, təşəkkür edirəm! İşə hazıram. Siz necəsiniz?",
        "Əla, sağ olun! Sizin üçün nə edə bilərəm?",
    ]) + _fayl_ipucu(s)


def _tesekkur(s, rejim):
    return random.choice(["Buyurun, həmişə xidmətinizdəyəm!", "Dəyməz! Başqa nə ilə kömək edim?", "Xoşdur! Növbəti iş nədir?"])


def _sagol(s, rejim):
    return random.choice(["Sağ olun, uğurlu iş günü!", "Görüşənədək! Təsdiqləriniz yadda saxlanılıb."])


def _yaradan(s, rejim):
    return (f"Mən Worker's Assistant-am, şirkətinizin daxili işləri üçün hazırlanmış köməkçiyəm. "
            f"Hazırkı iş rejimim: {rejim}. Hesablamaları öz modullarım aparır, AI yalnız sualı başa düşmək və izah etmək üçündür.")


_MEN_KIM = r"\bmen kim\w*|\bmeni tanı\w*|\bmeni tani\w*|kim kimi daxil|hansi hesabla"


def _men_kim(s, rejim):
    ad = getattr(s, "istifadeci_adi", "")
    if not ad:
        return "Sistem lokal rejimdə işləyir — giriş olmadan, ona görə sizi adla tanımıram."
    return (f"Siz **{ad}**-siniz — şirkət e-poçtunuzla daxil olmusunuz. Söhbətləriniz və yüklədiyiniz fayllar yalnız "
            f"sizə görünür; istənilən kompüterdən daxil olsanız, söhbətləriniz sol paneldə olacaq.")


NIYYETLER = [  # sıra önəmlidir: daha dəqiq olanlar əvvəl
    (_MEN_KIM, _men_kim), (_YARADAN, _yaradan), (_KIM, _kim), (_BACARIQ, _bacariq), (_HAL, _hal),
    (_TESEKKUR, _tesekkur), (_SAGOL, _sagol), (_SALAM, _salam),
]


def cavab(s, mesaj, rejim=""):
    """Şablon sual olarsa cavab mətni qaytarır, əks halda None."""
    n = _norm(mesaj)
    if not n.strip() or _is_suali(n):
        return None
    tapilan = [f for p, f in NIYYETLER if re.search(p, n)]
    if not tapilan:
        return None
    # "Salam, necəsən?" kimi birləşmələr: salamı qısa saxla, əsas niyyətə cavab ver
    esas = tapilan[0]
    if esas is _salam or len(tapilan) == 1:
        return esas(s, rejim)
    ad = _ad(s)
    return (f"Salam, {ad}! " if ad else "Salam! ") + esas(s, rejim) if _salam in tapilan else esas(s, rejim)
