"""Veb server: söhbət, fayl yükləmə, təsdiq və hesabat endpoint-ləri."""
import base64
import datetime as dt
import os
import pathlib
import secrets
import re
import threading
import uuid

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, Response
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from . import __version__, llm
from .giris import parol
from . import sohbet as SB
from . import teklif as TK
from .modullar import bank as BANK
from . import tools as T
from .state import HESABATLAR, YUKLEMELER, Sessiya

STATIC = pathlib.Path(__file__).parent / "static"
app = FastAPI(title="Worker's Assistant")
S = Sessiya()
BASLAMA = dt.datetime.now()
AI = llm.sec()

# Giriş: WA_PAROL verilibsə, bütün səhifə və API HTTP Basic ilə qorunur (serverdə mütləq).
ISTIFADECI = os.environ.get("WA_ISTIFADECI", "muhasib")
PAROL = parol()  # WA_PAROL və ya Docker secret (/run/secrets/wa_parol)


@app.middleware("http")
async def giris(request: Request, call_next):
    if PAROL and request.url.path != "/saglamliq":
        h = request.headers.get("authorization", "")
        ok = False
        if h.lower().startswith("basic "):
            try:
                ad, _, p = base64.b64decode(h[6:]).decode("utf-8").partition(":")
                ok = secrets.compare_digest(ad, ISTIFADECI) & secrets.compare_digest(p, PAROL)
            except Exception:
                ok = False
        if not ok:
            return Response("Giriş tələb olunur", status_code=401,
                            headers={"WWW-Authenticate": 'Basic realm="Worker\'s Assistant", charset="UTF-8"'})
    return await call_next(request)


@app.get("/saglamliq")
def saglamliq():
    """Docker/TeamCity yoxlaması üçün (girişsiz, məlumat qaytarmır)."""
    return {"ok": True, "versiya": __version__}


class Mesaj(BaseModel):
    mesaj: str


class Tesdiq(BaseModel):
    odenis_id: str
    qaime_id: str
    qerar: str


@app.get("/")
def index():
    return FileResponse(STATIC / "index.html")


@app.get("/api/veziyyet")
def veziyyet():
    ex = None
    if S.excel:
        vv = S.excel["vereqler"]
        ex = {"ad": S.excel["ad"], "vereq": len(vv), "setir": sum(v["setir"] for v in vv.values())}
    return {"hazir": S.fayl_var, "uzlasma": S.hazir, "excel": ex, "fayl": S.excel["ad"] if S.excel else None,
            "dovr_sonu": S.dovr_sonu.isoformat() if S.dovr_sonu else None, "rejim": AI.ad, "rejim_kod": getattr(AI, "kod", "qayda"),
            "modullar": T.modullar(), "versiya": __version__, "baslama": f"{BASLAMA:%H:%M}",
            "xulase": T.xulase(S)["ui"] if S.hazir else None,
            "teklifler": getattr(S, "teklifler", []),
            "sablon": S.sablon["ad"] if getattr(S, "sablon", None) else None,
            "bank": ({"daxilolma": len(S.bank["setirler"]), "yoxlanmali": sum(r["kateqoriya"] == "yoxlanmali" for r in S.bank["setirler"]),
                      "sual": len(BANK._sual_setirleri(S.bank)),
                      "cem": round(sum(r["mebleg"] for r in S.bank["setirler"]), 2)} if S.bank else None)}


ISLER = {}  # yükləmə işləri: böyük fayllar arxa planda çevrilir
KILID = threading.Lock()


@app.post("/api/yukle")
async def yukle(fayl: list[UploadFile] = File(...), dovr_sonu: str = Form("")):
    """Bir və ya bir neçə Excel: bank çıxarışı (və ya istənilən Excel) + istəyə görə inteqrasiya şablonu."""
    if len(fayl) > 5:
        raise HTTPException(400, "Bir dəfəyə ən çox 5 fayl yükləyin")
    yollar = []
    for f in fayl:
        if not f.filename.lower().endswith((".xlsx", ".xlsm")):
            raise HTTPException(400, f"«{f.filename}»: yalnız Excel faylı (.xlsx) qəbul olunur")
        ad = re.sub(r"[^\w.\-]", "_", pathlib.Path(f.filename).name)
        yol = YUKLEMELER / f"{dt.datetime.now():%Y%m%d_%H%M%S}_{uuid.uuid4().hex[:4]}_{ad}"
        with open(yol, "wb") as h:
            while blok := await f.read(1 << 20):
                h.write(blok)
        yollar.append((yol, pathlib.Path(f.filename).name))
    try:
        # verilməyibsə: keçən ayın son günü (uzlaşma faylıdırsa, istifadəçi sonra dəyişə bilər)
        dovr = dt.date.fromisoformat(dovr_sonu) if dovr_sonu else dt.date.today().replace(day=1) - dt.timedelta(days=1)
    except ValueError:
        raise HTTPException(400, "Dövrün sonu düzgün tarix deyil")
    is_id = uuid.uuid4().hex[:10]
    ISLER[is_id] = {"status": "isleyir", "i": 0, "n": 0, "vereq": ""}

    def progress(i, n, vereq):
        ISLER[is_id].update(i=i + 1, n=n, vereq=vereq)

    def isci():
        try:
            with KILID:
                sablonlar = [(y, a) for y, a in yollar if BANK.sablon_mi(y)]
                esas = [(y, a) for y, a in yollar if (y, a) not in sablonlar]
                if sablonlar:  # inteqrasiya şablonu: yadda saxlanır, 1C faylı onun üzərində doldurulur
                    S.sablon = {"yol": sablonlar[0][0], "ad": sablonlar[0][1]}
                if esas:
                    sablon = S.sablon
                    S.yukle(esas[0][0], dovr, orijinal_ad=esas[0][1], progress=progress)
                    S.sablon = sablon
                elif not S.fayl_var:
                    pass  # yalnız şablon yüklənib: bank çıxarışı gözlənilir
                S.teklifler = TK.hazirla(S)  # fayla görə hazır sorğular
            ISLER[is_id].update(status="hazir", netice={**veziyyet(), "yalniz_sablon": not esas,
                                                         "fayllar": [a for _, a in yollar]})
        except Exception as e:
            ISLER[is_id].update(status="xeta", xeta=f"Fayl oxunmadı: {e}")

    threading.Thread(target=isci, daemon=True).start()
    return {"is": is_id}


@app.get("/api/yukle/{is_id}")
def yukle_veziyyeti(is_id: str):
    if is_id not in ISLER:
        raise HTTPException(404, "İş tapılmadı")
    return ISLER[is_id]


class Dovr(BaseModel):
    dovr_sonu: str


@app.post("/api/dovr")
def dovr_deyis(d: Dovr):
    """Uzlaşma faylının dövrünü dəyişir: fayl yenidən yüklənmir, yalnız uzlaşma yenidən hesablanır."""
    if not S.hazir:
        raise HTTPException(400, "Yüklənmiş fayl uzlaşma faylı deyil, dövrün sonu lazım deyil")
    try:
        S.dovr_deyis(dt.date.fromisoformat(d.dovr_sonu))
        S.teklifler = TK.hazirla(S)
    except ValueError:
        raise HTTPException(400, "Dövrün sonu düzgün tarix deyil")
    return veziyyet()


@app.post("/api/sohbet")
def sohbet(m: Mesaj):
    sablon = SB.cavab(S, m.mesaj, AI.ad)  # salam, necəsən, sən kimsən...
    if sablon:
        S.tarixce += [{"role": "user", "content": m.mesaj}, {"role": "assistant", "content": sablon}]
        return {"cavab": sablon, "hadiseler": []}
    if not S.fayl_var:
        # sual yadda qalır: fayl yüklənən kimi interfeys onu təkrar soruşmadan cavablandırır
        return {"cavab": "Bu sual üçün fayl lazımdır. 📎 düyməsi ilə Excel faylını əlavə edin — yüklənən kimi sualınızı cavablandıracağam.",
                "hadiseler": [], "fayl_lazim": True}
    try:
        cavab, hadiseler = AI.cavab(S, m.mesaj)
    except Exception as e:
        return {"cavab": f"AI xidmətinə qoşulmaq alınmadı: {e}", "hadiseler": []}
    S.tarixce += [{"role": "user", "content": m.mesaj}, {"role": "assistant", "content": cavab or "…"}]
    if any(h.get("alet") in ("bank_teyin", "bank_kec", "hesab_ayarla") for h in hadiseler):
        S.teklifler = TK.hazirla(S)  # say dəyişdi (məs. yoxlanmalılar azaldı)
    return {"cavab": cavab, "hadiseler": hadiseler}


@app.post("/api/yeni")
def yeni():
    """Yeni söhbət: AI-ın yaddaşını sıfırlayır, yüklənmiş fayl və təsdiqlər qalır."""
    S.tarixce = []
    return {"ok": True}


@app.post("/api/tesdiq")
def tesdiq(t: Tesdiq):
    if not S.hazir:
        raise HTTPException(400, "Fayl yüklənməyib")
    try:
        S.tesdiq_et(t.odenis_id, t.qaime_id, t.qerar)
    except ValueError as e:
        raise HTTPException(400, str(e))
    return {"ok": True, "xulase": T.xulase(S)["ui"]}


@app.get("/api/fayl/{ad}")
def fayl(ad: str):
    yol = (HESABATLAR / ad).resolve()
    if yol.parent != HESABATLAR.resolve() or not yol.exists():
        raise HTTPException(404, "Fayl tapılmadı")
    return FileResponse(yol, filename=ad)


app.mount("/static", StaticFiles(directory=STATIC), name="static")
