# -*- coding: utf-8 -*-
"""Lecture sécurisée des clés API : variable d'environnement, fichier `.env` (local, jamais dans GitHub) ou Secrets Streamlit."""

from __future__ import annotations

import os
from pathlib import Path

_ENV_LOADED = False


def load_dotenv(path: str | Path = ".env") -> None:
    """Charge un fichier `.env` (KEY=valeur) dans os.environ, sans écraser les variables déjà définies."""
    global _ENV_LOADED
    if _ENV_LOADED:
        return
    _ENV_LOADED = True
    for base in (Path.cwd(), Path(__file__).resolve().parent.parent):
        p = base / path
        if not p.is_file():
            continue
        for line in p.read_text(encoding="utf-8", errors="ignore").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))
        return


def get_secret(name: str, default: str = "") -> str:
    """Ordre : variable d'environnement / .env, puis Secrets Streamlit."""
    load_dotenv()
    val = os.environ.get(name, "")
    if val:
        return val
    try:
        import streamlit as st
        return str(st.secrets.get(name, default) or default)
    except Exception:  # noqa: BLE001  (pas de secrets.toml)
        return default
