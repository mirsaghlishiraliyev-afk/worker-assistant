"""Ödənişləri qaimələrə bağlayan mühərrik.

Qaydalar ən etibarlıdan aşağıya doğru, ayrı keçidlərlə işləyir:
  1. Təyinatda qaimə nömrəsi          -> Yüksək
  2. VÖEN + yeganə eyni məbləğli qaimə  -> Orta
  3. VÖEN + qaimələrin cəmi = ödəniş    -> Orta
  4. Qalanlar -> istisna (sistem qərar vermir, mühasibə göstərir)
Bank (ƏDV-siz məbləğ) və ƏDV depozit axınları ayrıca uyğunlaşdırılır.
Sistem yalnız TƏKLİF edir; yekun qərar mühasibindir.
"""
import re
from itertools import combinations

from . import config as C
from .config import az

TOL = C.DOZUM


def _beraber(a, b):
    return abs(a - b) <= TOL + 1e-9


class Muherrik:
    def __init__(self, qaimeler, odenisler):
        self.q = qaimeler[~qaimeler["dublikat"]].copy()
        self.o = odenisler[~odenisler["dublikat"]].sort_values(["tarix", "id"]).copy()
        self.qalan = {}  # (axin, qaime_id) -> qalıq
        for _, r in self.q[~self.q["legv"]].iterrows():
            self.qalan[("bank", r["id"])] = r["netto"]
            self.qalan[("depozit", r["id"])] = r["edv"]
        self.qaime = {r["id"]: r for _, r in self.q.iterrows()}
        self.baglantilar = []
        self.istisnalar = []
        self.odenis_qalan = dict(zip(self.o["id"], self.o["mebleg"]))
        self.islenib = set()

    # ---------- köməkçi ----------
    def _aciq(self, axin, voen, nov):
        res = []
        for (a, qid), qal in self.qalan.items():
            r = self.qaime[qid]
            if a == axin and r["voen"] == voen and r["nov"] == nov and qal > TOL:
                res.append(qid)
        return sorted(res, key=lambda i: (self.qaime[i]["tarix"], i))

    def _bagla(self, p, qid, mebleg, qayda, etibar):
        self.baglantilar.append({
            "odenis_id": p["id"], "axin": p["axin"], "odenis_tarixi": p["tarix"], "voen": p["voen"], "kontragent": p["ad"],
            "odenis_meblegi": p["mebleg"], "metn": p["metn"], "qaime_id": qid, "baglanan_mebleg": round(mebleg, 2),
            "qayda": qayda, "etibar": etibar,
        })
        self.qalan[(p["axin"], qid)] = round(self.qalan[(p["axin"], qid)] - mebleg, 2)
        self.odenis_qalan[p["id"]] = round(self.odenis_qalan[p["id"]] - mebleg, 2)

    def _bolustur(self, p, qids, qayda, etibar):
        """Ödənişi qaimələrə sıra ilə paylayır; dözüm daxilindəki fərqi bağlayır."""
        for qid in qids:
            qal = self.qalan[(p["axin"], qid)]
            pq = self.odenis_qalan[p["id"]]
            if pq <= TOL:
                break
            if _beraber(pq, qal) or pq > qal:
                mebleg = qal if pq > qal + TOL else pq
                self._bagla(p, qid, mebleg, qayda, etibar)
                if _beraber(self.qalan[(p["axin"], qid)], 0):
                    self.qalan[(p["axin"], qid)] = 0.0
            else:
                self._bagla(p, qid, pq, qayda, etibar)
        if _beraber(self.odenis_qalan[p["id"]], 0):
            self.odenis_qalan[p["id"]] = 0.0
        self.islenib.add(p["id"])

    def _istisna(self, nov, risk, istinad, voen, izah):
        self.istisnalar.append({"nov": nov, "risk": risk, "istinad": istinad, "voen": voen, "izah": izah})

    # ---------- keçidlər ----------
    def kecid_nomre(self):
        for _, p in self.o.iterrows():
            tapilan = ["".join(m) for m in re.findall(C.QAIME_REGEX, p["metn"].upper())]
            if not tapilan:
                continue
            etibarli = []
            for qid in dict.fromkeys(tapilan):
                r = self.qaime.get(qid)
                if r is None or r["nov"] != p["nov"]:
                    continue
                if r["legv"]:
                    self._istisna("Ləğv edilmiş qaimə üzrə ödəniş", "Orta", qid, p["voen"],
                                  f"{p['ad']}: {qid} ləğv edilib, amma {p['tarix']:%d.%m.%Y} tarixində "
                                  f"{az(p['mebleg'])} AZN ödəniş ({p['axin']}) aparılıb.")
                    self.islenib.add(p["id"])
                    continue
                if r["voen"] != p["voen"]:
                    self._istisna("Qaimə başqa kontragentə aiddir", "Yüksək", qid, p["voen"],
                                  f"Təyinatda {qid} göstərilib, amma ödəyən VÖEN {p['voen']}, qaimədəki VÖEN {r['voen']}.")
                    continue
                etibarli.append(qid)
            if etibarli:
                self._bolustur(p, etibarli, "Təyinatda qaimə nömrəsi", "Yüksək")

    def kecid_mebleg(self):
        for _, p in self.o.iterrows():
            if p["id"] in self.islenib:
                continue
            aciq = self._aciq(p["axin"], p["voen"], p["nov"])
            eyni = [q for q in aciq if _beraber(self.qalan[(p["axin"], q)], p["mebleg"])]
            if len(eyni) == 1:
                self._bolustur(p, eyni, "VÖEN + yeganə eyni məbləğ", "Orta")
            elif len(eyni) > 1:
                self._istisna("Qeyri-müəyyən uyğunluq", "Orta", " / ".join(eyni), p["voen"],
                              f"{p['ad']}: {p['tarix']:%d.%m.%Y} tarixli {az(p['mebleg'])} AZN ödəniş ({p['axin']}) "
                              f"eyni məbləğli {len(eyni)} qaiməyə uyğun gəlir. Mühasib seçməlidir.")
                self.islenib.add(p["id"])

    def kecid_kombinasiya(self):
        for _, p in self.o.iterrows():
            if p["id"] in self.islenib:
                continue
            aciq = self._aciq(p["axin"], p["voen"], p["nov"])[:10]
            uygun = []
            for k in range(2, min(C.KOMBINASIYA_MAKS, len(aciq)) + 1):
                for komb in combinations(aciq, k):
                    if _beraber(sum(self.qalan[(p["axin"], q)] for q in komb), p["mebleg"]):
                        uygun.append(komb)
            if len(uygun) == 1:
                self._bolustur(p, list(uygun[0]), "VÖEN + qaimələrin cəmi", "Orta")
            elif len(uygun) > 1:
                self._istisna("Qeyri-müəyyən uyğunluq", "Orta", "; ".join("+".join(k) for k in uygun), p["voen"],
                              f"{p['ad']}: {az(p['mebleg'])} AZN ödəniş bir neçə qaimə kombinasiyasına uyğun gəlir.")
                self.islenib.add(p["id"])

    def kecid_qalanlar(self):
        for _, p in self.o.iterrows():
            if p["id"] in self.islenib:
                continue
            aciq = self._aciq(p["axin"], p["voen"], p["nov"])
            yer = "Bank" if p["axin"] == "bank" else "ƏDV depozit"
            if aciq:
                self._istisna("Uyğunlaşdırılmayıb", "Orta", f"{yer} {p['tarix']:%d.%m.%Y}", p["voen"],
                              f"{p['ad']}: {az(p['mebleg'])} AZN ödəniş heç bir açıq qaimə və ya cəmlə üst-üstə düşmür. "
                              f"Açıq qaimələr: {', '.join(aciq)}.")
            else:
                ad = "Qaiməsiz mədaxil" if p["nov"] == "Satış" else "Qaiməsiz məxaric"
                self._istisna(ad, "Yüksək", f"{yer} {p['tarix']:%d.%m.%Y}", p["voen"],
                              f"{p['ad']}: {az(p['mebleg'])} AZN, uyğun açıq qaimə yoxdur (avans və ya qeydə alınmamış qaimə ola bilər).")
            self.islenib.add(p["id"])

    def isle(self):
        self.kecid_nomre()
        self.kecid_mebleg()
        self.kecid_kombinasiya()
        self.kecid_qalanlar()
        # Bağlandıqdan sonra artıq qalan ödənişlər
        for _, p in self.o.iterrows():
            qal = self.odenis_qalan[p["id"]]
            if qal > TOL and any(b["odenis_id"] == p["id"] for b in self.baglantilar):
                qids = sorted({b["qaime_id"] for b in self.baglantilar if b["odenis_id"] == p["id"]})
                self._istisna("Artıq ödəniş", "Orta", ", ".join(qids), p["voen"],
                              f"{p['ad']}: {p['axin']} üzrə {az(qal)} AZN qaimə məbləğindən artıq ödənilib.")
        return self
