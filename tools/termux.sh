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

var() { python -c "import $1" 2>/dev/null; }
# pydantic-core və pandas mənbədən qurulanda maturin/meson Android API səviyyəsini tələb edir
export ANDROID_API_LEVEL="${ANDROID_API_LEVEL:-$(getprop ro.build.version.sdk)}"

if ! var "pandas, openpyxl, fastapi, uvicorn, multipart, ldap3"; then
    echo ">> İlk quraşdırma (20-40 dəqiqə çəkə bilər, Termux-u bağlamayın)…"
    pkg update -y
    command -v python >/dev/null || pkg install -y python
    # numpy Termux paketindən gəlir (telefonda kompilyasiyası çox uzun çəkir)
    var numpy || pkg install -y python-numpy
    # pandas: əvvəl Termux paketi; bu Python versiyası üçün yoxdursa — mənbədən (hazır numpy ilə)
    if ! var pandas; then
        pkg install -y python-pandas 2>/dev/null || true
    fi
    if ! var pandas; then
        { pkg install -y tur-repo && pkg update -y && pkg install -y python-pandas; } 2>/dev/null || true
    fi
    if ! var pandas; then
        echo ">> pandas mənbədən qurulur (uzun çəkir)…"
        pkg install -y build-essential binutils ninja patchelf
        pip install meson meson-python "cython>=3" versioneer
        pip install --no-build-isolation pandas
    fi
    # fastapi-nin pydantic-core hissəsi: Termux paketi yoxdursa Rust ilə qurulur
    if ! var pydantic_core; then
        pkg install -y python-pydantic-core 2>/dev/null || pkg install -y rust binutils
    fi
    pip install openpyxl fastapi uvicorn python-multipart ldap3 httpx
fi
for m in numpy pandas openpyxl fastapi uvicorn multipart ldap3; do
    var "$m" || { echo "!! «$m» qurulmadı — yuxarıdakı xətanın son 30 sətrini göndərin."; exit 1; }
done

export WA_DATA="${WA_DATA:-$HOME/wa_data}"   # yaddaş (qaydalar, hesabatlar) telefonun Termux qovluğunda
mkdir -p "$WA_DATA"
command -v termux-wake-lock >/dev/null && termux-wake-lock   # ekran sönəndə server dayanmasın
( sleep 3; command -v termux-open-url >/dev/null && termux-open-url http://127.0.0.1:8000 ) &
echo ">> Server: http://127.0.0.1:8000  (dayandırmaq: Ctrl+C)"
exec python -m assistant
