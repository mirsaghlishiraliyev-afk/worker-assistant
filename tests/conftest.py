"""Ümumi test ayarları."""
import os
import tempfile

import pytest

# Testlər heç vaxt reponun data/ qovluğuna yazmasın: assistant idxal olunmazdan əvvəl müvəqqəti qovluq
os.environ.setdefault("WA_DATA", tempfile.mkdtemp(prefix="wa_test_"))


@pytest.fixture(autouse=True)
def _qeydiyyat_qapisi(request, monkeypatch):
    """Qeydiyyat qapısı (telefon + peşə məcburi) yalnız «qeydiyyat» işarəli testlərdə açıqdır:
    köhnə testlərin istifadəçilərinin telefonu yoxdur."""
    try:
        from assistant import hesab as H
    except Exception:  # Android başladıcısı testləri assistant-ı ayrıca idxal edir
        yield
        return
    monkeypatch.setattr(H, "MECBURI", bool(request.node.get_closest_marker("qeydiyyat")))
    yield


def pytest_configure(config):
    config.addinivalue_line("markers", "qeydiyyat: qeydiyyat qapısı (telefon + peşə) açıq olan testlər")
