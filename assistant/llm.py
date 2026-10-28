"""Söhbət mühərrikləri. WA_LLM mühit dəyişəni ilə seçilir:

  qayda     - AI-sız, açar sözlərə əsaslanan lokal rejim (default). Məlumat kompüterdən çıxmır.
  anthropic - Claude API (ANTHROPIC_API_KEY, WA_MODEL). Alət nəticələri xarici serverə gedir: rəhbərliyin icazəsi lazımdır.
  openai    - OpenAI-uyğun API, məs. lokal Ollama (WA_OPENAI_BASE=http://localhost:11434/v1, WA_MODEL=qwen2.5:14b).
"""
import json
import os
import re

from . import tools as T
from .modullar import bank as BANK
from .modullar import excel as EX

EX_CAVAB = {
    "excel_xulase": "Faylın strukturu aşağıdadır. Hansı vərəq və ya sütunla işləyək?",
    "sutunlar": "Sütunlar üzrə statistika aşağıdadır.",
    "setirler": "Sətirlər aşağıdadır.",
    "qrupla": "Hesablama nəticəsi aşağıdadır.",
    "en_boyuk": "Seçilmiş sətirlər aşağıdadır.",
    "axtar": "Axtarış nəticəsi aşağıdadır.",
    "problemler": "Faylda tapılan problemlər aşağıdadır.",
}

SISTEM = """Sən Worker's Assistant-san: {sirket} şirkətində işçilərin əl ilə gördüyü təkrarlanan işləri avtomatlaşdıran
AI köməkçi. Həmişə Azərbaycan dilində, qısa və dəqiq cavab ver.

Aktiv modullar:
{modullar}

Qaydalar:
- Rəqəmləri, sənəd nömrələrini və adları YALNIZ alətlərin nəticəsindən götür. Heç vaxt uydurma.
- Alət nəticəsi interfeysdə cədvəl və ya düymə kimi artıq göstərilir; onu təkrar sadalama, əsas nəticəni 1-3 cümlə ilə izah et
  və növbəti addımı təklif et.
- Sən yalnız TƏKLİF edirsən. Təsdiq tələb edən əməliyyatı yalnız işçi konkret obyekt üçün açıq şəkildə istədikdə et.
- Məktub və sənədləri göndərmirsən, yalnız qaralama hazırlayırsan.
- Aktiv modullarda olmayan iş istənərsə, bunu hələ edə bilmədiyini de və hansı modulun lazım olduğunu qeyd et.
- Vergi və hüquq məsələlərində əmin deyilsənsə, "auditor və ya hüquqşünasla yoxlayın" de.
- Fayl yüklənməyibsə, 📎 düyməsi ilə faylı əlavə etməyi xahiş et."""


def sistem():
    from uzlasma import config as C
    mods = "\n".join(f"- {m['ad']}: {m['tesvir']}" for m in T.modullar() if m["status"] == "aktiv")
    return SISTEM.format(sirket=C.SIRKET_ADI, modullar=mods)


MAKS_ADDIM = 6


def _hadise(ad, args, netice):
    return {"alet": ad, "args": args or {}, "ui": netice.get("ui")}


# ---------------- 1. Lokal qaydalar ----------------
class QaydaMuherriki:
    kod = "qayda"
    ad = "Lokal rejim (AI-sız). Məlumat kompüterdən çıxmır"

    def cavab(self, s, mesaj):
        n = T.norm(mesaj)
        hadiseler = []

        def et(ad, **args):
            netice = T.cagir(s, ad, args)
            hadiseler.append(_hadise(ad, args, netice))
            return netice

        # Bank daxilolmaları: bank çıxarışı yüklənibsə əvvəl o
        if getattr(s, "bank", None):
            plan = BANK.qayda(s, mesaj)
            if plan:
                netice = et(plan[0], **plan[1])
                return netice["metn"], hadiseler
            if not re.search(r"sutun|struktur|vereq|problem|axtar|en boyuk|en kicik", n):
                et("bank_xulase")
                return ("Bank çıxarışı üçün bu sorğuları başa düşürəm: «Xülasə», «Sığorta haqlarını göstər», «Yoxlanmalıları göstər», "
                        "«Xəbərdarlıqlar», «12-ci sətir subroqasiyadır», «Subroqasiya maddəsi \"...\" olsun», «1C üçün Excel hazırla»."), hadiseler

        # bank sorğusu, amma yüklənmiş fayl bank çıxarışı deyil: başqa sorğuya çevirmə, düz de
        if not getattr(s, "bank", None) and s.excel and re.search(
                r"yoxlanmali|taninmayan|tanimadig|daxilolma|subroq|ezamiyye|telim haqq|1c ucun|1c-ye", n):
            vv = s.excel["vereqler"]
            uygun = [a for a in vv if re.search(r"yoxlan|istisna|tesdiq gozle|taninmayan", T.norm(a))]
            if uygun:
                netice = et("setirler", vereq=uygun[0], n=30)
                return (f"«Yoxlanmalı daxilolmalar» bank çıxarışı üçündür, bu fayl isə çıxarış deyil. Bu faylda yoxlanmalı "
                        f"məlumat «{uygun[0]}» vərəqindədir: " + netice["metn"].split("\n")[0]), hadiseler
            return (f"«Yoxlanmalı daxilolmalar» bank çıxarışı üçündür, «{s.excel['ad']}» isə bank çıxarışı deyil "
                    f"(Kredit/Mədaxil və Təyinat sütunları olan vərəq tapılmadı). Bank çıxarışını «⟳ Faylı dəyiş» ilə yükləyin — "
                    f"sual avtomatik cavablandırılacaq. Bu faylda səhvləri axtarmaq üçün «Problemləri tap» yazın."), []

        # Excel analizi: uzlaşma faylı deyilsə həmişə, uzlaşma faylıdırsa yalnız ümumi Excel sualları üçün
        ex_suali = re.search(r"sutun|struktur|\bcem|toplam|\borta|\bsay|en boyuk|en kicik|en cox|axtar|setir|problem tap|fayl haqqinda|goster|siyahi|umumi|cemi|ne qeder", n)
        if not s.hazir or ex_suali or (s.excel and not s.hazir):
            plan = EX.qayda(s, mesaj)
            if plan:
                netice = et(plan[0], **plan[1])
                if plan[0] in ("obyekt_gostericisi", "setirler", "problemler", "movzu_goster", "cem_hesabla", "hesab_ayarla") or (plan[0] == "qrupla" and not plan[1].get("qrup_sutunu")):
                    return netice["metn"].split("\n")[0], hadiseler
                return EX_CAVAB.get(plan[0], "Nəticə aşağıdadır."), hadiseler
            if not s.hazir:
                return self._komek_excel(s), []

        kontr = None
        for r in s.k.itertuples():
            ilk = T.norm(r.ad).split()[0]
            if re.search(rf"\b{re.escape(ilk)}", n) or r.voen in n:
                kontr = r.ad
                break
        qaime = re.search(r"\b([a-z]{2})\s?-?(\d{3,})\b", n)
        risk = ("Yüksək" if "yuksek" in n else "Orta" if re.search(r"\borta\b", n) else
                "Aşağı" if "asagi" in n else "Məlumat" if "melumat" in n else None)

        if re.search(r"komek|nece istifade|ne ede bilersen", n):
            return self._komek(), []
        if qaime:
            et("qaime_axtar", sorgu=qaime.group(1).upper() + qaime.group(2))
            return "Qaimə üzrə məlumat aşağıdadır.", hadiseler
        if "mektub" in n or ("yaz" in n and "borc" in n):
            if not kontr:
                return "Hansı kontragentə məktub hazırlayım? Adını yazın, məsələn: 'Alfa üçün borc məktubu'.", []
            et("borc_mektubu", kontragent=kontr)
            return "Məktub qaralaması hazırdır. Yoxlayıb e-poçtla göndərə bilərsiniz.", hadiseler
        if re.search(r"hesabat|excel|yukle", n):
            et("hesabat_yarat")
            return "Excel hesabatı hazırdır.", hadiseler
        if re.search(r"baglanti|tesdiq|gozleyen", n):
            et("baglantilar", status="gozleyen", kontragent=kontr)
            return "Təsdiq gözləyən bağlantılar aşağıdadır. Əvvəlcə 'Orta' etibarlıları yoxlayın.", hadiseler
        if "borc" in n:
            nov = "Kreditor" if re.search(r"kreditor|bizim borc|biz borclu", n) else ("Debitor" if "debitor" in n else None)
            et("borclar", nov=nov, kontragent=kontr)
            return "Açıq borclar aşağıdadır.", hadiseler
        if re.search(r"istisna|problem|uygunsuz|risk|sehv|edv|muqavile", n) or kontr:
            novs = "ƏDV" if "edv" in n else ("müqavilə" if "muqavile" in n else None)
            et("istisnalar", risk=risk, nov=novs, kontragent=kontr)
            return "Baxılmalı olan istisnalar aşağıdadır.", hadiseler
        if re.search(r"xulase|veziyyet|umumi|nece gedir", n):
            et("xulase")
            return "Uzlaşmanın ümumi vəziyyəti belədir.", hadiseler
        return self._komek(), []

    @staticmethod
    def _komek_excel(s):
        cols = []
        if s.excel:
            _, v = EX._v(s, None)
            cols = [str(c) for c in v["sutunlar"][:8]]
        return ("Lokal rejimdə Excel üçün bu sorğuları başa düşürəm:\n"
                "• «Kür Avto-nun borcu nə qədərdir?» / «Aynurun maaşı»\n• «Fayl haqqında ümumi məlumat»\n• «Sütunları göstər»\n• «Problemləri tap»\n"
                "• «<rəqəm sütunu> cəmi <sütun> üzrə»\n• «Ən böyük 10 <sütun>»\n• «'mətn' axtar»\n• «İlk 20 sətir»"
                + (f"\n\nBu faylın sütunları: {', '.join(cols)}" if cols else ""))

    @staticmethod
    def _komek():
        return ("Lokal rejimdə hələlik Mühasibat modulunun bu sorğularını başa düşürəm:\n"
                "• «Vəziyyət necədir?»\n• «Yüksək riskli istisnalar»\n• «Təsdiq gözləyən bağlantılar»\n"
                "• «Debitor borclar» / «Alfa borcu»\n• «AA1005 qaiməsi»\n• «Alfa üçün borc məktubu»\n• «Excel hesabatı»")


# ---------------- 2. Claude (Anthropic) ----------------
class AnthropicMuherriki:
    def __init__(self):
        import anthropic  # istəyə bağlı asılılıq
        self.client = anthropic.Anthropic()
        self.model = os.environ.get("WA_MODEL", "claude-sonnet-4-5")
        self.kod = "anthropic"
        self.ad = f"Claude ({self.model}). Məlumat Anthropic serverinə gedir"
        self.aletler = []
        for a in T.ALETLER:
            tesvir, sx = T.sxem(a)
            self.aletler.append({"name": a, "description": tesvir, "input_schema": sx})

    def cavab(self, s, mesaj):
        mesajlar = list(s.tarixce[-12:]) + [{"role": "user", "content": mesaj}]
        hadiseler = []
        for _ in range(MAKS_ADDIM):
            r = self.client.messages.create(model=self.model, max_tokens=1500, system=sistem(),
                                            tools=self.aletler, messages=mesajlar)
            if r.stop_reason != "tool_use":
                return "".join(b.text for b in r.content if b.type == "text"), hadiseler
            mesajlar.append({"role": "assistant", "content": r.content})
            neticeler = []
            for b in r.content:
                if b.type == "tool_use":
                    netice = T.cagir(s, b.name, b.input)
                    hadiseler.append(_hadise(b.name, b.input, netice))
                    neticeler.append({"type": "tool_result", "tool_use_id": b.id, "content": netice["metn"]})
            mesajlar.append({"role": "user", "content": neticeler})
        return "Sorğu çox addım tələb etdi. Zəhmət olmasa, daha konkret soruşun.", hadiseler


# ---------------- 3. OpenAI-uyğun (lokal Ollama və s.) ----------------
class OpenAIMuherriki:
    def __init__(self):
        import httpx
        self.http = httpx.Client(timeout=120)
        self.base = os.environ.get("WA_OPENAI_BASE", "http://localhost:11434/v1").rstrip("/")
        self.key = os.environ.get("WA_OPENAI_KEY", "ollama")
        self.model = os.environ.get("WA_MODEL", "qwen2.5:14b")
        self.kod = "openai"
        self.ad = f"Lokal AI modeli ({self.model}). Məlumat kompüterdən çıxmır"
        self.aletler = []
        for a in T.ALETLER:
            tesvir, sx = T.sxem(a)
            self.aletler.append({"type": "function", "function": {"name": a, "description": tesvir, "parameters": sx}})

    def cavab(self, s, mesaj):
        mesajlar = [{"role": "system", "content": sistem()}] + list(s.tarixce[-12:]) + \
                   [{"role": "user", "content": mesaj}]
        hadiseler = []
        for _ in range(MAKS_ADDIM):
            r = self.http.post(f"{self.base}/chat/completions", headers={"Authorization": f"Bearer {self.key}"},
                               json={"model": self.model, "messages": mesajlar, "tools": self.aletler})
            r.raise_for_status()
            m = r.json()["choices"][0]["message"]
            cagirislar = m.get("tool_calls") or []
            if not cagirislar:
                return m.get("content") or "", hadiseler
            mesajlar.append({"role": "assistant", "content": m.get("content") or "", "tool_calls": cagirislar})
            for c in cagirislar:
                ad = c["function"]["name"]
                args = c["function"].get("arguments") or "{}"
                args = json.loads(args) if isinstance(args, str) else args
                netice = T.cagir(s, ad, args)
                hadiseler.append(_hadise(ad, args, netice))
                mesajlar.append({"role": "tool", "tool_call_id": c.get("id", ad), "content": netice["metn"]})
        return "Sorğu çox addım tələb etdi. Zəhmət olmasa, daha konkret soruşun.", hadiseler


def sec():
    rejim = os.environ.get("WA_LLM", "qayda").lower()
    if rejim == "anthropic":
        return AnthropicMuherriki()
    if rejim == "openai":
        return OpenAIMuherriki()
    return QaydaMuherriki()
