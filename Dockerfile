# Worker's Assistant — test.server üçün Docker obrazı
FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    WA_HOST=0.0.0.0 \
    WA_PORT=8000 \
    WA_DATA=/data \
    WA_LLM=qayda

WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt

COPY assistant ./assistant
COPY uzlasma ./uzlasma
COPY tools ./tools
COPY tests ./tests

# Build zamanı testlər: biri sınsa, obraz yaranmır (TeamCity build-i dayanır)
RUN python -m pytest -q tests && rm -rf tests .pytest_cache

# Real məlumat (yüklənən fayllar, 1C keşi, hesabatlar, mühasibin qaydaları) obrazda deyil, volume-dadır
RUN useradd --create-home --uid 10001 wa && mkdir -p /data && chown wa:wa /data
USER wa
VOLUME ["/data"]
EXPOSE 8000

HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
  CMD python -c "import os,urllib.request; urllib.request.urlopen(f'http://127.0.0.1:{os.environ[\"WA_PORT\"]}/saglamliq', timeout=4)"

CMD ["python", "-m", "assistant"]
