"""İstifadə:  python -m uzlasma giris.xlsx --dovr-sonu 2026-09-30 --cixis hesabat.xlsx"""
import argparse
import datetime as dt

from . import checks
from . import loaders
from . import report
from .matcher import Muherrik


def isle(giris, dovr_sonu):
    loaders.yoxla(giris)
    q = loaders.qaimeler(giris)
    o = loaders.odenisler(giris)
    mq = loaders.muqavileler(giris)
    k = loaders.kontragentler(giris)
    m = Muherrik(q, o).isle()
    ist = checks.yoxla(m, q, o, mq, dovr_sonu)
    return m, ist, o, k


def main():
    ap = argparse.ArgumentParser(description="Bank, ƏDV depozit və e-qaimə uyğunlaşdırması")
    ap.add_argument("giris")
    ap.add_argument("--dovr-sonu", required=True, help="YYYY-AA-GG")
    ap.add_argument("--cixis", default="uzlasma_hesabati.xlsx")
    a = ap.parse_args()
    dovr_sonu = dt.date.fromisoformat(a.dovr_sonu)
    m, ist, o, k = isle(a.giris, dovr_sonu)
    report.yaz(a.cixis, m, ist, o, k, dovr_sonu)
    bagli = len({b["odenis_id"] for b in m.baglantilar})
    print(f"Ödənişlər: {len(m.o)}, avtomatik bağlanan: {bagli}, istisnalar: {len(ist)} -> {a.cixis}")


if __name__ == "__main__":
    main()
