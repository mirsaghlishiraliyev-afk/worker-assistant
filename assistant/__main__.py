"""İşə salmaq:  python -m assistant   (sonra brauzerdə http://127.0.0.1:8000)

Mühit dəyişənləri:
  WA_HOST   - dinlənən ünvan (susmaya görə 127.0.0.1 — yalnız bu kompüter; Docker-də 0.0.0.0)
  WA_PORT   - port (8000)
  WA_RELOAD - 1: kod dəyişəndə server özü yenidən başlayır (yalnız inkişaf zamanı)
  WA_IMAP_HOST, WA_IMAP_DOMEN (və ya WA_LDAP_*) - şəbəkəyə açılanda e-poçt hesabı ilə giriş; bax: assistant/giris.py.
                0.0.0.0-da giriş qurulmadan server işə düşmür.
"""
import os
import sys

import uvicorn

from assistant import giris

if __name__ == "__main__":
    host = os.environ.get("WA_HOST", "127.0.0.1")
    port = int(os.environ.get("WA_PORT", "8000"))
    if host not in ("127.0.0.1", "localhost") and not giris.aktiv():
        sys.exit("Təhlükəsizlik: server şəbəkəyə açılır (WA_HOST=%s), amma giriş qurulmayıb. Real maliyyə məlumatı "
                 "girişsiz açıla bilməz — WA_IMAP_HOST (məs. mail.gov.az) və WA_IMAP_DOMEN (məs. asf.gov.az) təyin edin." % host)
    print(f"Worker's Assistant: http://{host}:{port}")
    reload = os.environ.get("WA_RELOAD") == "1"
    uvicorn.run("assistant.app:app", host=host, port=port, reload=reload)
