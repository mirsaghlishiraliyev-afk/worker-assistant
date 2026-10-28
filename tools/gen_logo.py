"""Worker's Assistant loqosunu yaradır: python tools/gen_logo.py

Loqo: W (sol yarı — əl ilə, kağız) -> A -> W (sağ yarı — rəqəmsal piksellər və dövrə)."""
import pathlib

OUT = pathlib.Path(__file__).resolve().parents[1] / "assistant" / "static"


def lerp(a, b, t): return (a[0] + (b[0]-a[0])*t, a[1] + (b[1]-a[1])*t)
P = [(4.5, 8), (10, 24), (16, 10.5), (22, 24), (27.5, 8)]

KECID = (23.9, 18.5)  # bütöv xəttin bitdiyi, piksellərin başladığı nöqtə


def pikseller(sw, olcu):
    """Son xətt: bütöv xətt piksellərə çevrilir, sonda piksellər ayrılıb yuxarı-sağa uçur."""
    out = []
    # xətt üzərində piksellər (yuxarıya doğru seyrəkləşir)
    for i, t in enumerate([0.12, 0.36, 0.6, 0.84]):
        x, y = lerp(KECID, P[4], t)
        s = olcu * (1 - 0.08 * i)
        out.append(f'<rect x="{x-s/2:.2f}" y="{y-s/2:.2f}" width="{s:.2f}" height="{s:.2f}" rx="{s*0.15:.2f}" fill="#fff" opacity="{1-0.1*i:.2f}"/>')
    # uçan piksellər
    for i, (dx, dy) in enumerate([(2.6, -2.2), (4.2, -0.6), (3.4, 1.6)]):
        x, y = P[4][0] + dx * 0.55, P[4][1] + dy * 0.9
        s = olcu * (0.55 - 0.1 * i)
        out.append(f'<rect x="{x-s/2:.2f}" y="{y-s/2:.2f}" width="{s:.2f}" height="{s:.2f}" rx="{s*0.15:.2f}" fill="#fff" opacity="{0.75-0.15*i:.2f}"/>')
    return "\n    ".join(out)

def isare(detalli=True, sw=2.6):
    sol = f'<polyline points="4.5,8 10,24 16,10.5 22,24 {KECID[0]},{KECID[1]}" fill="none" stroke="#fff" stroke-width="{sw}" stroke-linecap="round" stroke-linejoin="round"/>'
    # A-nın üfüqi xətti: sol tərəfdə bütöv, sağa doğru rəqəmsallaşır
    xett = f'<line x1="12.6" y1="18.6" x2="19.4" y2="18.6" stroke="#fff" stroke-width="{sw}" stroke-linecap="round"/>'
    pix = pikseller(sw, sw*0.95)
    # son nöqtə: dövrə düyünü
    duyun = ""
    elave = ""
    if detalli:
        # sol yarıda kağız xətləri, sağ yarıda dövrə izləri
        elave = '''<g stroke="#fff" stroke-opacity=".18" stroke-width=".6" stroke-linecap="round">
      <line x1="3" y1="27.5" x2="12" y2="27.5"/><line x1="3" y1="29.2" x2="9.5" y2="29.2"/>
    </g>
    <g fill="none" stroke="#fff" stroke-opacity=".28" stroke-width=".7" stroke-linecap="round" stroke-linejoin="round">
      <polyline points="25.5,20 28,20 29.5,21.5"/><polyline points="23,27.5 26.5,27.5 28.5,25.5"/>
    </g>
    <g fill="#fff" fill-opacity=".5"><circle cx="29.6" cy="21.6" r=".8"/><circle cx="28.6" cy="25.4" r=".8"/></g>'''
    return f"{elave}\n    {sol}\n    {xett}\n    {pix}\n    {duyun}"

GRAD = '''<defs><linearGradient id="g" x1="0" y1="0" x2="0" y2="1">
    <stop offset="0" stop-color="#5a5a60"/><stop offset="1" stop-color="#36363b"/>
  </linearGradient></defs>'''

def tam(detalli=True):
    return f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32">
  {GRAD}
  <rect width="32" height="32" rx="8" fill="url(#g)"/>
  <!-- Sol: əl ilə görülən iş (bütöv xətt). Sağ: rəqəmsal (piksellər və dövrə düyünü). Ortada A: Assistant -->
  <g>
    {isare(detalli)}
  </g>
</svg>'''

def ikon():  # fonsuz, interfeysdə gradient kvadratın içində
    return f'<svg width="20" height="20" viewBox="0 0 32 32">{isare(False, 2.9)}</svg>'

open(OUT / "logo.svg", "w", encoding="utf-8").write(tam(True))
open(OUT / "favicon.svg", "w", encoding="utf-8").write(tam(False))
print("İnterfeys ikonu (index.html-də .logo və WA_IKON, bir sətirdə):\n" + " ".join(ikon().split()))
