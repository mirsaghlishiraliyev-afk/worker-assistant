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
from . import hesab as H
from . import profil as PR
from . import sohbet as SB
from . import tarixce as TR
from . import kitabxana as KB
from . import teklif as TK
from .modullar import bank as BANK
from .modullar import qebul as QB
from .modullar import stok as STOK
from .modullar import tapsiriq as TP
from .modullar import teleb as TELEB
from . import tools as T
from .state import HESABATLAR, YUKLEMELER, Sessiya

STATIC = pathlib.Path(__file__).parent / "static"
app = FastAPI(title="Worker's Assistant")
BASLAMA = dt.datetime.now()
AI = llm.sec()
QB.planlayici_baslat()  # WhatsApp API qurulubsa: qəbula yaxın xatırlatmalar

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
ACIQ = {"/saglamliq", "/giris", "/giris/google", "/giris/google/geri", "/giris/kod", "/giris/mobil", "/qeydiyyat",
        "/giris/bagla", "/giris/bagla/kod",
        "/static/logo.svg", "/static/favicon.svg"}
QEYD_KUKI = "wa_qeyd"  # telefonla ilk giriş: təsdiqlənmiş nömrə, qeydiyyat tamamlanana qədər
BAGLA_KUKI = "wa_bagla"  # poçtla giriş telefonla yaradılmış hesaba: o hesabın nömrəsinə kodla təsdiq
G.MOBIL_MUMKUN = H.mobil_mumkun


@app.middleware("http")
async def giris_yoxla(request: Request, call_next):
    if G.aktiv():
        u = G.kuki_oxu(request.cookies.get(G.KUKI))
        if not u and request.url.path not in ACIQ:
            if request.url.path.startswith("/api/"):
                return JSONResponse({"detail": "Giriş tələb olunur", "giris": True}, status_code=401)
            return RedirectResponse("/giris", status_code=303)
        if u and H.MECBURI and not H.tam_mi(u["l"]) and not request.url.path.startswith(("/qeydiyyat", "/cixis", "/static", "/giris")):
            # qeydiyyat tamamlanmayıb (telefon və ya peşə yoxdur): sistem açılmır
            if request.url.path.startswith("/api/"):
                return JSONResponse({"detail": "Qeydiyyat tamamlanmayıb", "qeydiyyat": True}, status_code=403)
            return RedirectResponse("/qeydiyyat", status_code=303)
        if u:
            _CARI.set(u["l"])
            request.state.istifadeci = u
            sessiya(u["l"]).istifadeci_adi = u.get("a") or u["l"]  # salamlaşmada adla müraciət üçün
            sessiya(u["l"]).istifadeci_login = u["l"]
            TP.qeyd_et(u["l"], u.get("a") or u["l"])  # rəhbər işçini adı ilə tapa bilsin
    return await call_next(request)


def _https(request):
    return request.url.scheme == "https" or request.headers.get("x-forwarded-proto", "").lower() == "https"


@app.get("/giris")
def giris_sehifesi(request: Request):
    if not G.aktiv() or G.kuki_oxu(request.cookies.get(G.KUKI)):
        return RedirectResponse("/", status_code=303)
    return HTMLResponse(G.sehife())


@app.post("/giris")
def giris_et(request: Request, login: str = Form(""), parol: str = Form(""), usul: str = Form("outlook")):
    if not G.aktiv():
        return RedirectResponse("/", status_code=303)
    if G.YERLI:  # telefon rejimi: parolsuz, ünvan istifadəçini tanımaq üçündür
        try:
            return _poct_ile_daxil(request, G.yerli_yoxla(login, usul if usul in ("outlook", "gmail") else "outlook"))
        except G.GirisXetasi as x:
            return HTMLResponse(G.sehife(str(x), login, usul), status_code=400)
    acar = f"{(login or '').strip().lower()}|{request.client.host if request.client else ''}"
    if G.bloklanib(acar):
        return HTMLResponse(G.sehife("Çox sayda uğursuz cəhd. 10 dəqiqə sonra yenidən yoxlayın.", login), status_code=429)
    if not G.outlook_aktiv():
        return RedirectResponse("/giris", status_code=303)
    try:
        u = G.YOXLAYICI(login, parol)
    except G.GirisXetasi as x:
        return HTMLResponse(G.sehife(str(x), login), status_code=503)
    if not u:
        G.ugursuz(acar)
        return HTMLResponse(G.sehife("İstifadəçi adı və ya parol səhvdir.", login), status_code=401)
    G.ugurlu(acar)
    return _poct_ile_daxil(request, u)


def _daxil_et(request, u):
    r = RedirectResponse("/", status_code=303)
    r.set_cookie(G.KUKI, G.kuki_yarat(u), max_age=int(G.SESSIYA_SAAT * 3600), httponly=True,
                 samesite="lax", secure=_https(request), path="/")
    return r


# ---------------- mobil giriş və qeydiyyat ----------------
@app.post("/giris/kod")
def giris_kod(request: Request, tel: str = Form(""), kanal: str = Form("sms")):
    """Telefon nömrəsinə SMS / WhatsApp təsdiq kodu (girişdə və qeydiyyatda)."""
    acar = f"kod|{request.client.host if request.client else ''}"
    if G.bloklanib(acar):
        return {"ok": False, "metn": "Çox sayda sorğu. 10 dəqiqə sonra yenidən cəhd edin."}
    n = H.kod_gonder(tel, kanal)
    if not n["ok"]:
        G.ugursuz(acar)
    return n


def _qeyd_kukisi(request):
    import time as _t
    d = G._imza_oxu(request.cookies.get(QEYD_KUKI))
    return d if isinstance(d, dict) and d.get("exp", 0) > _t.time() and d.get("tel") else None


@app.post("/giris/mobil")
def giris_mobil(request: Request, tel: str = Form(""), kod: str = Form("")):
    import time as _t
    acar = f"mobil|{QB.telefon(tel)}|{request.client.host if request.client else ''}"
    if G.bloklanib(acar):
        return HTMLResponse(G.sehife("Çox sayda uğursuz cəhd. 10 dəqiqə sonra yenidən yoxlayın.", usul="mobil", tel=tel), status_code=429)
    ok, xeta = H.kod_yoxla(tel, kod)
    if not ok:
        G.ugursuz(acar)
        return HTMLResponse(G.sehife(xeta, usul="mobil", tel=tel), status_code=401)
    G.ugurlu(acar)
    h = H.tel_ile(tel)
    if h:  # sistem istifadəçini nömrəsi ilə tanıyır (poçtun parolunu unutsa da öz hesabına girir)
        return _daxil_et(request, {"login": h["login"], "ad": h.get("ad") or h["login"]})
    r = RedirectResponse("/qeydiyyat", status_code=303)  # ilk dəfə: qeydiyyat
    r.set_cookie(QEYD_KUKI, G._imzala({"tel": QB.telefon(tel), "exp": int(_t.time() + 1800)}), max_age=1800, httponly=True,
                 samesite="lax", secure=_https(request), path="/")
    return r


def _poct_ile_daxil(request, u):
    """Poçtla (Outlook/Gmail) giriş: ünvan telefonla yaradılmış hesaba bağlıdırsa — həmin hesaba (qaldığı yerdən davam).
    Serverdə bunun üçün o hesabın nömrəsinə gələn kod lazımdır (poçt ünvanını yazmaq hesabı «tutmağa» yetməsin)."""
    import time as _t
    mail = H.poct(u.get("mail")) or H.poct(u["login"])
    hesab = H.mail_ile(mail) if mail else None
    if not hesab or hesab["login"] == u["login"]:
        return _daxil_et(request, u)
    if G.YERLI or hesab.get("mail_tesdiq"):  # telefonda (parolsuz rejim) və ya bağlantı əvvəl təsdiqlənib
        H.mail_tesdiqle(hesab["login"])
        return _daxil_et(request, {"login": hesab["login"], "ad": hesab.get("ad") or u.get("ad", "")})
    r = RedirectResponse("/giris/bagla", status_code=303)
    r.set_cookie(BAGLA_KUKI, G._imzala({"hedef": hesab["login"], "mail": mail, "exp": int(_t.time() + 900)}), max_age=900,
                 httponly=True, samesite="lax", secure=_https(request), path="/")
    return r


def _bagla_kukisi(request):
    import time as _t
    d = G._imza_oxu(request.cookies.get(BAGLA_KUKI))
    h = H.al(d.get("hedef")) if isinstance(d, dict) and d.get("exp", 0) > _t.time() else None
    return (d, h) if h else (None, None)


@app.get("/giris/bagla")
def bagla_sehifesi(request: Request, xeta: str = ""):
    d, h = _bagla_kukisi(request)
    if not h:
        return RedirectResponse("/giris", status_code=303)
    return HTMLResponse(H.bagla_sehifesi(d["mail"], h, xeta))


@app.post("/giris/bagla/kod")
def bagla_kod(request: Request, kanal: str = Form("sms")):
    d, h = _bagla_kukisi(request)
    if not h:
        return {"ok": False, "metn": "Vaxt keçib — yenidən daxil olun."}
    return H.kod_gonder(h["tel"], kanal)  # nömrə brauzerə göndərilmir: hesabdan götürülür


@app.post("/giris/bagla")
def bagla(request: Request, kod: str = Form("")):
    d, h = _bagla_kukisi(request)
    if not h:
        return RedirectResponse("/giris", status_code=303)
    ok, xeta = H.kod_yoxla(h["tel"], kod)
    if not ok:
        return HTMLResponse(H.bagla_sehifesi(d["mail"], h, xeta), status_code=401)
    H.mail_tesdiqle(h["login"])
    r = _daxil_et(request, {"login": h["login"], "ad": h.get("ad") or h["login"]})
    r.delete_cookie(BAGLA_KUKI, path="/")
    return r


def _qeyd_kim(request):
    """Qeydiyyatı kim tamamlayır: daxil olmuş istifadəçi (telefonu/poçtu/peşəsi yoxdur) və ya telefonla ilk dəfə gələn.
    Qaytarır: (login, ad, tel, tel təsdiqli, poçt, poçt dəyişməz)."""
    u = G.kuki_oxu(request.cookies.get(G.KUKI))
    if u:
        h = H.al(u["l"]) or {}
        mail = h.get("mail") or H.poct(u["l"])
        return (u["l"], h.get("ad") or (u.get("a") if u.get("a") != u["l"] else ""), h.get("tel", ""), bool(h.get("tel_tesdiq")),
                mail, bool(H.poct(u["l"])))
    q = _qeyd_kukisi(request)
    if q:
        return q["tel"], H.ad_tap(tel=q["tel"]), q["tel"], True, "", False
    return None


@app.get("/qeydiyyat")
def qeydiyyat(request: Request):
    k = _qeyd_kim(request)
    if not k:
        return RedirectResponse("/giris", status_code=303)
    login, ad, tel, tesdiqli, mail, sabit = k
    prf = PR.al(login)
    return HTMLResponse(H.qeydiyyat_sehifesi(ad=ad, tel=tel, tesdiqli=tesdiqli, secilmis=(prf or {}).get("peseler", []), mail=mail,
                                             mail_sabit=sabit, izah=None if ad else "Adınızı tapa bilmədim — ad və soyadınızı yazın."))


@app.post("/qeydiyyat")
def qeydiyyat_saxla(request: Request, ad: str = Form(""), tel: str = Form(""), kod: str = Form(""), mail: str = Form(""),
                    peseler: list[str] = Form([])):
    k = _qeyd_kim(request)
    if not k:
        return RedirectResponse("/giris", status_code=303)
    login, _, kohne_tel, tesdiqli, kohne_mail, sabit = k
    ad, t = " ".join(ad.split()), QB.telefon(tel) or (QB.telefon(kohne_tel) if tesdiqli else "")
    m = kohne_mail if sabit else H.poct(mail)

    def xeta(x):
        return HTMLResponse(H.qeydiyyat_sehifesi(x, ad=ad, tel=tel or kohne_tel, tesdiqli=tesdiqli, secilmis=peseler,
                                                 mail=mail or kohne_mail, mail_sabit=sabit), status_code=400)
    if len(ad) < 3:
        return xeta("Ad soyadı yazın.")
    if not t:
        return xeta("Telefon nömrəsi məcburidir (məs.: 050 123 45 67).")
    if not m:
        return xeta("Poçt ünvanı məcburidir (məs.: ad.soyad@asf.gov.az və ya ad@gmail.com).")
    basqa = H.mail_ile(m)
    if basqa and basqa["login"] != login:
        return xeta("Bu poçt ünvanı başqa hesaba bağlıdır — həmin poçtla daxil olun.")
    if not [p for p in peseler if p in PR.KODLAR]:
        return xeta("Peşə məcburidir — siyahıdan ən azı birini seçin.")
    tesdiq = tesdiqli and t == QB.telefon(kohne_tel)
    if not tesdiq and kod:
        ok, kod_xetasi = H.kod_yoxla(t, kod)
        if not ok:
            return xeta(kod_xetasi)
        tesdiq = True
    if not tesdiq and H.mobil_mumkun():
        return xeta("Nömrəni təsdiqləyin: «Təsdiq kodu göndər», sonra gələn kodu yazın.")
    H.saxla(login, ad, t, tesdiq, m)
    PR.saxla(login, peseler)
    TP.qeyd_et(login, ad)
    r = _daxil_et(request, {"login": login, "ad": ad})
    r.delete_cookie(QEYD_KUKI, path="/")
    return r


def _google_geri_unvan(request):
    if G.GOOGLE_REDIRECT:
        return G.GOOGLE_REDIRECT
    sxem = "https" if _https(request) else request.url.scheme
    return f"{sxem}://{request.headers.get('host', request.url.netloc)}/giris/google/geri"


@app.get("/giris/google")
def google_giris(request: Request):
    """Gmail: Google-un giriş səhifəsinə yönləndirir (parol bu sistemə gəlmir)."""
    if not (G.GOOGLE_ID and G.GOOGLE_SECRET):  # telefon rejimində Gmail — formla, OAuth yox
        return RedirectResponse("/giris", status_code=303)
    url, kuki = G.google_basla(_google_geri_unvan(request))
    r = RedirectResponse(url, status_code=303)
    r.set_cookie(G.OAUTH_KUKI, kuki, max_age=600, httponly=True, samesite="lax", secure=_https(request), path="/giris/google")
    return r


@app.get("/giris/google/geri")
def google_geri(request: Request, code: str = "", state: str = "", error: str = ""):
    if not (G.GOOGLE_ID and G.GOOGLE_SECRET):  # telefon rejimində Gmail — formla, OAuth yox
        return RedirectResponse("/giris", status_code=303)

    def xeta(m, kod=400):
        r = HTMLResponse(G.sehife(m, usul="gmail"), status_code=kod)
        r.delete_cookie(G.OAUTH_KUKI, path="/giris/google")
        return r
    if error:
        return xeta("Google ilə giriş ləğv olundu." if error == "access_denied" else f"Google girişi alınmadı ({error}).")
    d = G.google_kuki_oxu(request.cookies.get(G.OAUTH_KUKI), state)
    if not d or not code:
        return xeta("Giriş sessiyası köhnəlib və ya yanlışdır. «Google ilə daxil ol»-u yenidən basın.")
    acar = f"google|{request.client.host if request.client else ''}"
    if G.bloklanib(acar):
        return xeta("Çox sayda uğursuz cəhd. 10 dəqiqə sonra yenidən yoxlayın.", 429)
    try:
        u = G.GOOGLE_YOXLAYICI(code, *d)
    except G.GirisXetasi as x:
        G.ugursuz(acar)
        return xeta(str(x), 403)
    G.ugurlu(acar)
    r = _poct_ile_daxil(request, u)
    r.delete_cookie(G.OAUTH_KUKI, path="/giris/google")
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
    fayllar: list = []  # söhbətdə istifadə olunan fayllar: [{"id", "ad", "nov"}] («Fayllarım»a istinad)


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
        TR.saxla(_CARI.get(), sid, s.ad, s.mesajlar, s.fayllar)
    except ValueError as e:
        raise HTTPException(400, str(e))
    return {"ok": True}


@app.delete("/api/sohbetler/{sid}")
def sohbet_sil(sid: str):
    TR.sil(_CARI.get(), sid)
    return {"ok": True}


# ---------------- profil: istifadəçinin peşələri ----------------
class Profil(BaseModel):
    peseler: list[str] = []


@app.get("/api/profil")
def profil_al():
    return {"profil": PR.al(_CARI.get()), "peseler": PR.PESELER}


@app.put("/api/profil")
def profil_saxla(p: Profil):
    try:
        prf = PR.saxla(_CARI.get(), p.peseler)
        return {"profil": prf, "telimat": PR.telimat(prf)}
    except ValueError as e:
        raise HTTPException(400, str(e))


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
            "modullar": PR.modul_siyahisi(T.modullar(), PR.al(_CARI.get())),  # yalnız peşəyə aid modullar
            "versiya": __version__, "baslama": f"{BASLAMA:%H:%M}",
            "xulase": T.xulase(S)["ui"] if S.hazir else None,
            "teklifler": getattr(S, "teklifler", []),
            "istifadeci": {"ad": u.get("a") or u["l"], "login": u["l"]} if u else None,
            "profil": PR.al(_CARI.get()), "peseler": PR.PESELER, "heller": PR.heller(PR.al(_CARI.get())),
            "pese_effektiv": PR.effektiv(PR.al(_CARI.get())),
            "bildiris": TP.bildirisler(_CARI.get()),  # yeni tapşırıq, gecikən, cavab gözləyən sorğu
            "sablon": S.sablon["ad"] if getattr(S, "sablon", None) else None,
            "muqavile": ({"ad": S.muqavileler["ad"], "say": len(S.muqavileler["siyahi"]),
                          "uygun": sum(m["uygun"] for m in S.muqavileler["siyahi"])} if getattr(S, "muqavileler", None) else None),
            "bank": ({"daxilolma": len(S.bank["setirler"]), "yoxlanmali": sum(r["kateqoriya"] == "yoxlanmali" for r in S.bank["setirler"]),
                      "sual": len(BANK._sual_setirleri(S.bank)),
                      "cem": round(sum(r["mebleg"] for r in S.bank["setirler"]), 2)} if S.bank else None)}


ISLER = {}  # yükləmə işləri: böyük fayllar arxa planda çevrilir
KILID = threading.Lock()


def _dovr(dovr_sonu):
    try:
        # verilməyibsə: keçən ayın son günü (uzlaşma faylıdırsa, istifadəçi sonra dəyişə bilər)
        return dt.date.fromisoformat(dovr_sonu) if dovr_sonu else dt.date.today().replace(day=1) - dt.timedelta(days=1)
    except ValueError:
        raise HTTPException(400, "Dövrün sonu düzgün tarix deyil")


def _is_baslat(yollar, dovr, rejim="", kitabxanaya=True, ref=None):
    """Faylları arxa planda oxuyur (yeni yüklənən və ya «Fayllarım»dan açılan). yollar: [(yol, orijinal ad)]."""
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
                muqler = [(y, a) for y, a in yollar if (y, a) not in sablonlar and BANK.muqavile_fayli_mi(y)]
                esas = [(y, a) for y, a in yollar if (y, a) not in sablonlar and (y, a) not in muqler]
                refler = list(ref or [])  # söhbətdə fayla istinad: «Fayllarım»dakı id ilə (sonra yenidən açmaq üçün)
                if kitabxanaya:  # «Fayllarım»: kim nəyi yükləyib — yadda qalır
                    def _elave(y, a, nov):
                        q, yol = KB.elave(login, a, y, nov, rejim)
                        refler.append({"id": q["id"], "ad": a, "nov": nov})
                        return pathlib.Path(yol), a
                    sablonlar = [_elave(y, a, "sablon") for y, a in sablonlar]
                    muqler = [_elave(y, a, "muqavile") for y, a in muqler]
                    esas = [_elave(y, a, "fayl") for y, a in esas]
                if sablonlar:  # inteqrasiya şablonu: yadda saxlanır, 1C faylı onun üzərində doldurulur
                    S.sablon = {"yol": sablonlar[0][0], "ad": sablonlar[0][1]}
                if muqler:  # müqavilələr: nömrəsi olmayan mədaxillər FİN ilə bunlara tutuşdurulur
                    S.muqavileler = {"ad": muqler[0][1], "siyahi": BANK.muqavile_oxu(muqler[0][0])}
                if esas:
                    sablon, muq = S.sablon, getattr(S, "muqavileler", None)
                    S.yukle(esas[0][0], dovr, orijinal_ad=esas[0][1], progress=progress)
                    S.sablon, S.muqavileler = sablon, muq
                S.teklifler = TK.hazirla(S)  # fayla görə hazır sorğular
            ISLER[is_id].update(status="hazir", netice={**veziyyet(), "yalniz_sablon": not esas and bool(sablonlar), "fayl_ref": refler,
                                                         "yalniz_muqavile": not esas and not sablonlar and bool(muqler),
                                                         "fayllar": [a for _, a in yollar]})
        except Exception as e:
            ISLER[is_id].update(status="xeta", xeta=f"Fayl oxunmadı: {e}")

    threading.Thread(target=isci, daemon=True).start()
    return {"is": is_id}


@app.post("/api/yukle")
async def yukle(fayl: list[UploadFile] = File(...), dovr_sonu: str = Form(""), rejim: str = Form("")):
    """Bir və ya bir neçə Excel: bank çıxarışı (və ya istənilən Excel) + istəyə görə inteqrasiya şablonu.
    rejim="uzerine": eyni adlı köhnə faylın üstünə yazılır; əks halda ayrıca saxlanır."""
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
    return _is_baslat(yollar, _dovr(dovr_sonu), rejim=rejim)


# ---------------- «Fayllarım»: istifadəçinin yüklədiyi fayllar ----------------
class Yoxla(BaseModel):
    fayllar: list = []


@app.get("/api/fayllar")
def fayllarim():
    return KB.siyahi(_CARI.get())


@app.post("/api/fayllar/yoxla")
def fayllari_yoxla(y: Yoxla):
    """Yükləmədən əvvəl: bu adda fayl artıq yüklənibmi (interfeys «üstünə yazım?» soruşur)."""
    return KB.yoxla(_CARI.get(), y.fayllar)


@app.post("/api/fayllar/{fid}/ac")
def fayli_ac(fid: str, dovr_sonu: str = ""):
    e = KB.al(_CARI.get(), fid)
    if not e or not pathlib.Path(e["yol"]).exists():
        raise HTTPException(404, "Fayl tapılmadı (silinib ola bilər) — yenidən yükləyin")
    return _is_baslat([(pathlib.Path(e["yol"]), e["ad"])], _dovr(dovr_sonu), kitabxanaya=False,
                      ref=[{"id": e["id"], "ad": e["ad"], "nov": e.get("nov", "fayl")}])


@app.delete("/api/fayllar/{fid}")
def fayli_sil(fid: str):
    if not KB.sil(_CARI.get(), fid):
        raise HTTPException(404, "Fayl tapılmadı")
    return {"ok": True}


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
    if re.fullmatch(r"\s*(telimat|telimati goster|nece istifade edim|istifade telimati)\s*[?!.]*\s*", T.norm(m.mesaj)):
        t = PR.telimat(PR.al(_CARI.get()))  # peşəyə görə təlimat (fayl lazım deyil)
        cavab = "Peşənizə görə istifadə təlimatı:" if t else "Əvvəlcə peşənizi seçin — yan paneldə «Peşələrim»."
        return {"cavab": cavab, "hadiseler": [{"alet": "telimat", "args": {}, "ui": t}] if t else []}
    plan = TP.qayda(S, m.mesaj) or TELEB.qayda(S, m.mesaj)  # tapşırıqlar, sorğular, stoka tələb: fayl lazım deyil
    if not plan and not S.fayl_var:  # təsdiqlənmiş tələblər üzrə sifariş: stok faylı olmadan da
        plan = STOK.qayda(S, m.mesaj)
    if plan:
        netice = T.cagir(S, plan[0], plan[1])
        S.tarixce += [{"role": "user", "content": m.mesaj}, {"role": "assistant", "content": netice["metn"]}]
        return {"cavab": netice["metn"], "hadiseler": [{"alet": plan[0], "args": plan[1], "ui": netice.get("ui")}]}
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
