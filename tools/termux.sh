#!/data/data/com.termux/files/usr/bin/bash
# Worker's Assistant-ı Android telefonda (Termux) işə salır.
#
# Bir dəfəlik:
#   1) Termux-u F-Droid-dən qurun (Google Play-dəki köhnədir): https://f-droid.org/packages/com.termux/
#   2) Termux-da:  termux-setup-storage   (Yükləmələr qovluğundakı Excel-lərə çıxış üçün)
#   3) Proyekti telefona gətirin (git clone və ya ZIP) və bu skripti işə salın:
#        bash worker-assistant/tools/termux.sh
# Sonrakı dəfələr eyni əmr serveri dərhal başladır. Brauzerdə açın: http://127.0.0.1:8000
#
# Server yalnız telefonun özündə (127.0.0.1) açılır: başqası ona qoşula bilmir, ona görə giriş tələb olunmur.
set -e
KOK="$(cd "$(dirname "$0")/.." && pwd)"
cd "$KOK"

if ! python -c "import pandas, openpyxl, fastapi, uvicorn, multipart, ldap3" 2>/dev/null; then
    echo ">> İlk quraşdırma (10-20 dəqiqə çəkə bilər)…"
    pkg update -y
    pkg install -y python
    # pandas/numpy Termux paketindən gəlir (pip ilə telefonda kompilyasiya çox uzun çəkir)
    pkg install -y python-numpy python-pandas 2>/dev/null || { pkg install -y tur-repo && pkg install -y python-numpy python-pandas; }
    # fastapi-nin pydantic-core hissəsi Rust tələb edə bilər
    pkg install -y python-pydantic-core 2>/dev/null || pkg install -y rust binutils
    # pydantic-core mənbədən qurulanda maturin Android API səviyyəsini tələb edir
    export ANDROID_API_LEVEL="${ANDROID_API_LEVEL:-$(getprop ro.build.version.sdk)}"
    pip install openpyxl fastapi uvicorn python-multipart ldap3 httpx
fi

export WA_DATA="${WA_DATA:-$HOME/wa_data}"   # yaddaş (qaydalar, hesabatlar) telefonun Termux qovluğunda
mkdir -p "$WA_DATA"
command -v termux-wake-lock >/dev/null && termux-wake-lock   # ekran sönəndə server dayanmasın
( sleep 3; command -v termux-open-url >/dev/null && termux-open-url http://127.0.0.1:8000 ) &
echo ">> Server: http://127.0.0.1:8000  (dayandırmaq: Ctrl+C)"
exec python -m assistant
