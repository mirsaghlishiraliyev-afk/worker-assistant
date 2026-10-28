"""İşə salmaq:  python -m assistant   (sonra brauzerdə http://127.0.0.1:8000)"""
import os

import uvicorn

if __name__ == "__main__":
    port = int(os.environ.get("WA_PORT", "8000"))
    print(f"Worker's Assistant: http://127.0.0.1:{port}")
    # Yalnız bu kompüterdən açılır. Şəbəkəyə açmaq üçün autentifikasiya əlavə olunmalıdır.
    uvicorn.run("assistant.app:app", host="127.0.0.1", port=port)
