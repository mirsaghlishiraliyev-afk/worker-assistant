"""Veb server: söhbət, fayl yükləmə, təsdiq və hesabat endpoint-ləri."""
import contextvars
import datetime as dt
import pathlib
import re
import threading
import uuid

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from . import __version__, llm
from . import giris as G
from . import sohbet as SB
from . import tarixce as TR
from . import teklif as TK
from .modullar import bank as BANK
from . import tools as T
from .state import HESABATLAR, YUKLEMELER, Sessiya

STATIC = pathlib.Path(__file__).parent / "static"
app = FastAPI(title="Worker's Assistant")
BASLAMA = dt.datetime.now()
AI = llm.sec()

# ---------------- hər istifadəçinin öz sessiyası ----------------
# Giriş aktivdirsə, hər domen istifadəçisinin öz faylı, söhbəti və nəticələri olur (bir-birini görmürlər).
_CARI = contextvars.ContextVar("istifadeci", default="")
_SESSIYALAR, _SK = {}, threading.Lock()


def sessiya(login=None):
    login = _CARI.get() if login is None else login
    with _SK:
        if login not in _SESSIYALAR:
            _SESSIYALAR[login] = Sessiya()
            _SESSIYALAR[login].icazeli_fayllar = set()
        return _SESSIYALAR[login]


class _Proksi:
    """Kod «S.excel», «S.yukle(...)» yazır — proksi bunu cari istifadəçinin sessiyasına yönləndirir."""
    def __getattr__(self, k):
        return getattr(sessiya(), k)

    def __setattr__(self, k, v):
        setattr(sessiya(), k, v)


S = _Proksi()

# ---------------- giriş (şirkət e-poçtu: IMAP mail.gov.az və ya domen) ----------------
ACIQ = {"/saglamliq", "/giris", "/static/logo.svg", "/static/favicon.svg"}


@app.middleware("http")
async def giris_yoxla(request: Request, call_next):
    if G.aktiv():
        u = G.kuki_oxu(request.cookies.get(G.KUKI))
        if not u and request.url.path not in ACIQ:
            if request.url.path.startswith("/api/"):
                return JSONResponse({"detail": "Giriş tələb olunur", "giris": True}, status_code=401)
            return RedirectResponse("/giris", status_code=303)
        if u:
            _CARI.set(u["l"])
            request.state.istifadeci = u
            sessiya(u["l"]).istifadeci_adi = u.get("a") or u["l"]  # salamlaşmada adla müraciət üçün
    return await call_next(request)


def _https(request):
    return request.url.scheme == "https" or request.headers.get("x-forwarded-proto", "").lower() == "https"


@app.get("/giris")
def giris_sehifesi(request: Request):
    if not G.aktiv() or G.kuki_oxu(request.cookies.get(G.KUKI)):
        return RedirectResponse("/", status_code=303)
    return HTMLResponse(G.sehife())


@app.post("/giris")
def giris_et(request: Request, login: str = Form(""), parol: str = Form("")):
    if not G.aktiv():
        return RedirectResponse("/", status_code=303)
    acar = f"{(login or '').strip().lower()}|{request.client.host if request.client else ''}"
    if G.bloklanib(acar):
        return HTMLResponse(G.sehife("Çox sayda uğursuz cəhd. 10 dəqiqə sonra yenidən yoxlayın.", login), status_code=429)
    try:
        u = G.YOXLAYICI(login, parol)
    except G.GirisXetasi as x:
        return HTMLResponse(G.sehife(str(x), login), status_code=503)
    if not u:
        G.ugursuz(acar)
        return HTMLResponse(G.sehife("İstifadəçi adı və ya parol səhvdir.", login), status_code=401)
    G.ugurlu(acar)
    r = RedirectResponse("/", status_code=303)
    r.set_cookie(G.KUKI, G.kuki_yarat(u), max_age=int(G.SESSIYA_SAAT * 3600), httponly=True,
                 samesite="lax", secure=_https(request), path="/")
    return r


@app.get("/cixis")
def cixis():
    r = RedirectResponse("/giris" if G.aktiv() else "/", status_code=303)
    r.delete_cookie(G.KUKI, path="/")
    return r


@app.get("/saglamliq")
def saglamliq():
    """Docker/TeamCity yoxlaması üçün (girişsiz, məlumat qaytarmır)."""
    return {"ok": True, "versiya": __version__}


class Mesaj(BaseModel):
    mesaj: str


class Sohbet(BaseModel):
    ad: str = ""
    mesajlar: list = []


# ---------------- söhbət tarixçəsi (hər istifadəçinin özü) ----------------
@app.get("/api/sohbetler")
def sohbetler():
    return TR.siyahi(_CARI.get())


@app.get("/api/sohbetler/{sid}")
def sohbet_al(sid: str):
    s = TR.al(_CARI.get(), sid)
    if not s:
        raise HTTPException(404, "Söhbət tapılmadı")
    return s


@app.put("/api/sohbetler/{sid}")
def sohbet_saxla(sid: str, s: Sohbet):
    try:
        TR.saxla(_CARI.get(), sid, s.ad, s.mesajlar)
    except ValueError as e:
        raise HTTPException(400, str(e))
    return {"ok": True}


@app.delete("/api/sohbetler/{sid}")
def sohbet_sil(sid: str):
    TR.sil(_CARI.get(), sid)
    return {"ok": True}


class Tesdiq(BaseModel):
    odenis_id: str
    qaime_id: str
    qerar: str


@app.get("/")
def index():
    return FileResponse(STATIC / "index.html")


@app.get("/api/veziyyet")
def veziyyet(request: Request = None):
    u = getattr(request.state, "istifadeci", None) if request is not None else None
    ex = None
    if S.excel:
        vv = S.excel["vereqler"]
        ex = {"ad": S.excel["ad"], "vereq": len(vv), "setir": sum(v["setir"] for v in vv.values())}
    return {"hazir": S.fayl_var, "uzlasma": S.hazir, "excel": ex, "fayl": S.excel["ad"] if S.excel else None,
            "dovr_sonu": S.dovr_sonu.isoformat() if S.dovr_sonu else None, "rejim": AI.ad, "rejim_kod": getattr(AI, "kod", "qayda"),
            "modullar": T.modullar(), "versiya": __version__, "baslama": f"{BASLAMA:%H:%M}",
            "xulase": T.xulase(S)["ui"] if S.hazir else None,
            "teklifler": getattr(S, "teklifler", []),
            "istifadeci": {"ad": u.get("a") or u["l"], "login": u["l"]} if u else None,
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
    login = _CARI.get()  # arxa plan axınında istifadəçi kontekstdən gəlmir: əvvəlcədən götürülür
    ISLER[is_id] = {"status": "isleyir", "i": 0, "n": 0, "vereq": "", "_sahib": login}

    def progress(i, n, vereq):
        ISLER[is_id].update(i=i + 1, n=n, vereq=vereq)

    def isci():
        _CARI.set(login)
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
    if is_id not in ISLER or ISLER[is_id].get("_sahib") != _CARI.get():
        raise HTTPException(404, "İş tapılmadı")
    return {k: v for k, v in ISLER[is_id].items() if not k.startswith("_")}


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
    for h in hadiseler:  # istifadəçi yalnız özünə hazırlanmış faylları yükləyə bilər
        ui = h.get("ui") or {}
        for f in ([ui] if ui.get("tip") == "fayl" else ui.get("fayllar", []) if ui.get("tip") == "fayllar" else []):
            S.icazeli_fayllar.add(f["ad"])
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
    if G.aktiv() and ad not in S.icazeli_fayllar:
        raise HTTPException(404, "Fayl tapılmadı")
    if yol.parent != HESABATLAR.resolve() or not yol.exists():
        raise HTTPException(404, "Fayl tapılmadı")
    return FileResponse(yol, filename=ad)


app.mount("/static", StaticFiles(directory=STATIC), name="static")
