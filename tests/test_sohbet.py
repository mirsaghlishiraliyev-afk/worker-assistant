"""Şablon suallar: salam, necəsən, sən kimsən... və iş suallarının şablona düşməməsi."""
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
from assistant import sohbet  # noqa: E402


class SaxtaSessiya:
    hazir = False
    fayl_var = False
    fayl = None
    excel = None


S = SaxtaSessiya()


def test_sablon_suallar():
    for q in ["Salam", "Necəsən?", "Salam, necəsən?", "Sən kimsən?", "Nə edə bilərsən?", "Təşəkkür edirəm", "Hələlik"]:
        assert sohbet.cavab(S, q), q


def test_is_suallari_sablona_dusmur():
    for q in ["Salam, Alfa borcu nə qədərdir?", "Vəziyyət necədir?", "Yüksək riskli istisnaları göstər", "AA1005"]:
        assert sohbet.cavab(S, q) is None, q
