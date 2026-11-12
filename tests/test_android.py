"""Android başladıcısı (android/app/src/main/python/wa_android.py): GitHub zip-indən kodun çıxarılması və serverin qalxması."""
import io
import json
import pathlib
import socket
import sys
import tempfile
import urllib.request
import zipfile

import pytest  # noqa: E402

KOK = pathlib.Path(__file__).resolve().parents[1]
ANDROID = KOK / "android" / "app" / "src" / "main" / "python"
if not (ANDROID / "wa_android.py").exists():  # Docker obrazında android/ yoxdur (server üçün lazım deyil)
    pytest.skip("android/ qovluğu yoxdur", allow_module_level=True)
sys.path.insert(0, str(KOK))
sys.path.insert(0, str(ANDROID))
import wa_android as W  # noqa: E402


def _zip(fayllar):
    b = io.BytesIO()
    with zipfile.ZipFile(b, "w") as z:
        for ad, metn in fayllar.items():
            z.writestr(ad, metn)
    return b.getvalue()


def test_yalniz_kod_qovluqlari_cixarilir():
    h = pathlib.Path(tempfile.mkdtemp())
    z = _zip({"worker-assistant-abc/assistant/__init__.py": "__version__ = '9'", "worker-assistant-abc/assistant/static/a.html": "x",
              "worker-assistant-abc/uzlasma/__init__.py": "", "worker-assistant-abc/tests/test_x.py": "pis",
              "worker-assistant-abc/assistant/../../kenar.py": "pis", "worker-assistant-abc/README.md": "x"})
    assert W.kodu_cixar(z, h)
    fayllar = sorted(str(p.relative_to(h)) for p in h.rglob("*") if p.is_file())
    assert fayllar == ["assistant/__init__.py", "assistant/static/a.html", "uzlasma/__init__.py"]   # testlər, ../ yolları yox


def test_github_yeni_versiya_ve_internetsiz(monkeypatch):
    kok = pathlib.Path(tempfile.mkdtemp())
    z = _zip({"r-1/assistant/__init__.py": "", "r-1/uzlasma/__init__.py": ""})
    cavab = {"commits": json.dumps({"sha": "abcdef1234567890"}).encode(), "codeload": z}
    monkeypatch.setattr(W, "_al", lambda url, limit=15: cavab["commits" if "api.github" in url else "codeload"])
    kod, menbe = W.son_kod(kok)
    assert kod.name == "abcdef123456" and menbe == "GitHub (abcdef1)"

    def internet_yoxdur(url, limit=15):
        raise OSError("şəbəkə yoxdur")
    monkeypatch.setattr(W, "_al", internet_yoxdur)
    kod2, menbe2 = W.son_kod(kok)
    assert kod2 == kod and "əvvəlki versiya" in menbe2                              # internet yoxdursa — əvvəlki


def test_server_telefondaki_kimi_qalxir(monkeypatch):
    monkeypatch.setenv("WA_ANDROID_YENILEME", "0")
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        port = s.getsockname()[1]
    info = json.loads(W.basla(tempfile.mkdtemp(), port))
    assert info["isleyir"] and info["versiya"]
    with urllib.request.urlopen(f"http://127.0.0.1:{port}/saglamliq", timeout=5) as r:
        assert json.loads(r.read())["ok"] is True
    assert json.loads(W.basla("/yox", port))["port"] == port                        # ikinci çağırış: yenidən başlatmır
