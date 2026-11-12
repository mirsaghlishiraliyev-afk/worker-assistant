"""Telefonda serveri qaldırır: GitHub-dan son versiyanı çəkir (alınmasa — əvvəlki və ya APK-dakı), uvicorn-u başladır.

MainActivity.java bunu çağırır: basla(<tətbiqin qovluğu>, 8000). Server yalnız 127.0.0.1-də dinləyir (giriş tələb olunmur),
məlumat (yüklənən fayllar, qaydalar, söhbətlər) <tətbiqin qovluğu>/data-dadır.

Kodun yeniləməsi: GitHub API-dən main-in son commit-i alınır; yenidirsə, zip yüklənir və yalnız assistant/ və uzlasma/
çıxarılır (<qovluq>/kod/<sha>). Yeni kod idxal olunmasa (məs. APK-da olmayan yeni paket tələb edir) — APK-dakı
versiyaya qayıdılır. Ən çox iki versiya saxlanır.
"""
import io
import json
import os
import pathlib
import shutil
import sys
import threading
import time
import urllib.request
import zipfile

REPO = "mirsaghlishiraliyev-afk/worker-assistant"
BUDAQ = "main"
VAXT_LIMITI = 15
_VEZIYYET = {"isleyir": False}


def _al(url, limit=VAXT_LIMITI):
    req = urllib.request.Request(url, headers={"User-Agent": "WorkerAssistant-Android", "Accept": "application/vnd.github+json"})
    with urllib.request.urlopen(req, timeout=limit) as r:
        return r.read()


def kodu_cixar(zip_baytlari, hedef):
    """GitHub zip-indən yalnız assistant/ və uzlasma/ qovluqlarını hədəfə çıxarır."""
    hedef = pathlib.Path(hedef)
    with zipfile.ZipFile(io.BytesIO(zip_baytlari)) as z:
        for ad in z.namelist():
            hisseler = ad.split("/", 1)  # «worker-assistant-<sha>/assistant/...»
            if len(hisseler) < 2 or not hisseler[1].startswith(("assistant/", "uzlasma/")) or ad.endswith("/"):
                continue
            nisbi = pathlib.PurePosixPath(hisseler[1])
            if ".." in nisbi.parts or "__pycache__" in nisbi.parts:
                continue
            yol = hedef.joinpath(*nisbi.parts)
            yol.parent.mkdir(parents=True, exist_ok=True)
            yol.write_bytes(z.read(ad))
    return (hedef / "assistant" / "__init__.py").exists()


def son_kod(kok):
    """GitHub-dan son versiya; alınmasa — əvvəl yüklənmiş ən yenisi; heç biri yoxdursa None (APK-dakı işlənir)."""
    kod_q = pathlib.Path(kok) / "kod"
    kod_q.mkdir(parents=True, exist_ok=True)
    movcud = sorted((p for p in kod_q.iterdir() if (p / "assistant" / "__init__.py").exists()), key=lambda p: p.stat().st_mtime)
    if os.environ.get("WA_ANDROID_YENILEME", "1") == "0":
        return (movcud[-1] if movcud else None), "yeniləmə söndürülüb"
    try:
        sha = json.loads(_al(f"https://api.github.com/repos/{REPO}/commits/{BUDAQ}"))["sha"]
        hedef = kod_q / sha[:12]
        if not (hedef / "assistant" / "__init__.py").exists():
            mvq = kod_q / (sha[:12] + ".yuklenir")
            shutil.rmtree(mvq, ignore_errors=True)
            if not kodu_cixar(_al(f"https://codeload.github.com/{REPO}/zip/refs/heads/{BUDAQ}", 60), mvq):
                raise RuntimeError("zip-də assistant/ yoxdur")
            mvq.rename(hedef)
        os.utime(hedef)
        for kohne in sorted((p for p in kod_q.iterdir() if p != hedef), key=lambda p: p.stat().st_mtime)[:-1]:
            shutil.rmtree(kohne, ignore_errors=True)  # cari + bir əvvəlki saxlanır
        return hedef, f"GitHub ({sha[:7]})"
    except Exception as x:  # internet yoxdur, GitHub əlçatmazdır və s.
        return (movcud[-1] if movcud else None), f"GitHub əlçatmaz ({type(x).__name__}); " + ("əvvəlki versiya" if movcud else "APK-dakı versiya")


def _idxal(kod):
    if kod:
        sys.path.insert(0, str(kod))
    try:
        from assistant import app as A  # noqa: F401
        return A
    except Exception:
        if not kod:
            raise
        sys.path.remove(str(kod))  # yeni kod bu APK-da işləmir: APK-dakı versiyaya qayıt
        for m in [m for m in sys.modules if m.split(".")[0] in ("assistant", "uzlasma")]:
            del sys.modules[m]
        from assistant import app as A
        return A


def basla(kok, port=8000):
    """Serveri qaldırır (artıq işləyirsə — yenidən başlatmır). Qaytarır: JSON mətn (versiya, mənbə)."""
    if _VEZIYYET["isleyir"]:
        return json.dumps(_VEZIYYET, ensure_ascii=False)
    kok = pathlib.Path(kok)
    os.environ.setdefault("WA_DATA", str(kok / "data"))
    os.environ.setdefault("WA_HOST", "127.0.0.1")
    os.environ.setdefault("WA_PORT", str(port))
    os.environ.setdefault("WA_LLM", "qayda")
    pathlib.Path(os.environ["WA_DATA"]).mkdir(parents=True, exist_ok=True)
    kod, menbe = son_kod(kok)
    A = _idxal(kod)
    if kod and str(kod) not in sys.path:
        menbe += " — yeni kod bu APK-da işləmədi, APK-dakı versiya"
    import uvicorn

    import contextlib

    class Server(uvicorn.Server):  # arxa plan axınında siqnal tutulmur (Android-də yalnız əsas axın tuta bilər)
        def install_signal_handlers(self):  # uvicorn < 0.29
            pass

        @contextlib.contextmanager
        def capture_signals(self):  # uvicorn >= 0.29
            yield

    server = Server(uvicorn.Config(A.app, host="127.0.0.1", port=int(port), log_level="warning"))
    threading.Thread(target=server.run, name="wa-server", daemon=True).start()
    for _ in range(200):  # hazır olana qədər gözlə (ən çox ~20 san.)
        if getattr(server, "started", False):
            break
        time.sleep(0.1)
    from assistant import __version__
    _VEZIYYET.update(isleyir=True, versiya=__version__, menbe=menbe, port=int(port))
    return json.dumps(_VEZIYYET, ensure_ascii=False)
