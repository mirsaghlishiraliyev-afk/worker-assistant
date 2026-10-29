"""Giriş parolu: WA_PAROL mühit dəyişənindən və ya Docker secret faylından.

Serverdə tövsiyə olunan: Docker secret (parol compose faylında, TeamCity log-unda və `docker inspect`-də görünmür):
    printf '%s' 'PAROL' | docker secret create wa_parol -
Konteyner onu /run/secrets/wa_parol faylından oxuyur (və ya WA_PAROL_FILE ilə verilən yoldan).
"""
import os
import pathlib


def parol():
    if os.environ.get("WA_PAROL"):
        return os.environ["WA_PAROL"]
    yol = pathlib.Path(os.environ.get("WA_PAROL_FILE", "/run/secrets/wa_parol"))
    try:
        return yol.read_text(encoding="utf-8").strip()
    except OSError:
        return ""
