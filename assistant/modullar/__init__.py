"""Worker's Assistant modulları.

Hər modul bir iş sahəsidir (mühasibat, kadrlar, sənəd dövriyyəsi...) və bunları təqdim edir:
  MODUL    - ad, təsvir, status ("aktiv" / "planlaşdırılır") və interfeys üçün hazır sorğular
  ALETLER  - {alət_adı: (funksiya, təsvir, parametrlər)}
  MECBURI  - {alət_adı: [mütləq parametrlər]}

Yeni modul əlavə etmək: bu qovluqda fayl yaradın və aşağıdakı AKTIV siyahısına əlavə edin.
"""
from . import bank, excel, muhasibat

AKTIV = [bank, muhasibat, excel]

# Hələ qurulmayan, amma yol xəritəsində olan sahələr (interfeysdə "tezliklə" kimi görünür)
PLANLASDIRILIR = [
    {"kod": "muqavile", "ad": "Müqavilələr", "tesvir": "Reyestr, müddət bildirişləri, müqavilə müqayisəsi", "status": "planlaşdırılır"},
    {"kod": "senedler", "ad": "Sənəd dövriyyəsi", "tesvir": "Məktubların qeydiyyatı, cavab qaralamaları, axtarış", "status": "planlaşdırılır"},
    {"kod": "kadr", "ad": "Kadrlar", "tesvir": "Məzuniyyət, əmrlər, tabel və arayışlar", "status": "planlaşdırılır"},
]

ALETLER, MECBURI, SAHIB = {}, {}, {}
for _m in AKTIV:
    for _ad in _m.ALETLER:
        if _ad in ALETLER:
            raise RuntimeError(f"Alət adı təkrarlanır: {_ad}")
        SAHIB[_ad] = _m.MODUL["kod"]
    ALETLER.update(_m.ALETLER)
    MECBURI.update(_m.MECBURI)


def modullar():
    return [m.MODUL for m in AKTIV] + PLANLASDIRILIR
