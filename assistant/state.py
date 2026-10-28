"""Sessiyanın vəziyyəti: yüklənmiş fayl, mühərrikin nəticələri və mühasibin təsdiqləri."""
import datetime as dt
import json
import os
import pathlib

from uzlasma.cli import isle
from uzlasma.loaders import FaylXetasi

from .modullar import bank as BANK
from .modullar import excel as EX

KOK = pathlib.Path(__file__).resolve().parents[1]
DATA = pathlib.Path(os.environ.get("WA_DATA", KOK / "data"))
YUKLEMELER = DATA / "yuklemeler"
HESABATLAR = DATA / "hesabatlar"
TESDIQ_FAYLI = DATA / "tesdiqler.json"
EXCEL_KES = DATA / "excel_kes"  # böyük Excel faylların SQLite keşi (fayl hash-inə görə)
EX.AYAR_FAYLI = DATA / "hesab_ayarlari.json"  # mühasibin öyrətdiyi hesab qaydaları
BANK.QAYDA_FAYLI = DATA / "bank_qaydalari.json"  # mühasibin bank düzəlişləri (ödəyici -> kateqoriya, maddələr)
BANK.HESABAT_QOVLUGU = HESABATLAR


class Sessiya:
    """Pilot üçün bir istifadəçili sessiya. Çox istifadəçi üçün sonra verilənlər bazasına keçiriləcək."""

    def __init__(self):
        for p in (YUKLEMELER, HESABATLAR):
            p.mkdir(parents=True, exist_ok=True)
        self.fayl = None
        self.dovr_sonu = None
        self.m = None
        self.ist = []
        self.o = None
        self.k = None
        self.excel = None  # istənilən Excel: {"ad", "db", "vereqler"}
        self.teklifler = []  # fayla görə hazır sorğular (teklif.py)
        self.sablon = None  # mühasibin inteqrasiya şablonu: {"yol", "ad"} — 1C faylı onun üzərində doldurulur
        self.bank = None  # bank çıxarışı: {"ad", "setirler", "xeberdarliqlar", ...}
        self.son_niyyet = None  # cavab gözləyən sual (məs. "hansı vərəqi cəmləyim?")
        self.tarixce = []  # yalnız mətn növbələri: [{"role": "user"|"assistant", "content": str}]
        self.tesdiqler = self._tesdiq_oxu()

    @property
    def hazir(self):
        """Uzlaşma (Mühasibat) nəticələri hazırdır."""
        return self.m is not None

    @property
    def fayl_var(self):
        return self.excel is not None

    def yukle(self, fayl, dovr_sonu: dt.date, orijinal_ad=None, progress=None):
        """İstənilən Excel-i oxuyur; uzlaşma şablonuna uyğundursa, Mühasibat mühərrikini də işə salır.
        Qaytarır: "uzlasma" və ya "excel"."""
        fayl = pathlib.Path(fayl)
        self.excel = EX.ac(fayl, EXCEL_KES, ad=orijinal_ad or fayl.name, progress=progress)
        self.fayl, self.dovr_sonu = fayl, dovr_sonu
        self.tarixce = []
        self.son_niyyet = None
        self.bank = BANK.oxu(fayl, ad=orijinal_ad or fayl.name)  # bank çıxarışıdırsa, daxilolmalar təsnif olunur
        if self.bank:
            self.m, self.ist, self.o, self.k = None, [], None, None
            return "bank"
        try:
            self.m, self.ist, self.o, self.k = isle(str(fayl), dovr_sonu)
            return "uzlasma"
        except FaylXetasi:
            self.m, self.ist, self.o, self.k = None, [], None, None
            return "excel"

    def dovr_deyis(self, dovr_sonu: dt.date):
        self.m, self.ist, self.o, self.k = isle(str(self.fayl), dovr_sonu)
        self.dovr_sonu = dovr_sonu

    # ---------- təsdiqlər ----------
    def _tesdiq_oxu(self):
        try:
            return json.loads(TESDIQ_FAYLI.read_text(encoding="utf-8"))
        except (FileNotFoundError, json.JSONDecodeError):
            return {}

    def _acar(self, odenis_id, qaime_id):
        h = self.o.loc[self.o["id"] == odenis_id, "hash"]
        return f"{h.iloc[0] if len(h) else odenis_id}|{qaime_id}"

    def tesdiq(self, odenis_id, qaime_id):
        return self.tesdiqler.get(self._acar(odenis_id, qaime_id), {}).get("qerar", "")

    def tesdiq_et(self, odenis_id, qaime_id, qerar, kim="mühasib"):
        movcud = any(b["odenis_id"] == odenis_id and b["qaime_id"] == qaime_id for b in self.m.baglantilar)
        if not movcud:
            raise ValueError(f"{odenis_id} → {qaime_id} bağlantısı tapılmadı")
        if qerar not in ("Bəli", "Xeyr"):
            raise ValueError("Qərar 'Bəli' və ya 'Xeyr' olmalıdır")
        self.tesdiqler[self._acar(odenis_id, qaime_id)] = {
            "qerar": qerar, "kim": kim, "vaxt": dt.datetime.now().isoformat(timespec="seconds")}
        TESDIQ_FAYLI.write_text(json.dumps(self.tesdiqler, ensure_ascii=False, indent=2), encoding="utf-8")
