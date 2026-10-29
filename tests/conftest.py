"""Testlər real data/ qovluğuna (yüklənən fayllar, söhbətlər, qaydalar) yazmasın: hər işə salınmada boş müvəqqəti qovluq."""
import os
import tempfile

os.environ["WA_DATA"] = tempfile.mkdtemp(prefix="wa_test_")
