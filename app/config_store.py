# -*- coding: utf-8 -*-
"""Centralized persistent config: always under %APPDATA%, with migration.

Previously config.json lived next to the exe / project root, which fails
under Program Files (no write permission) and mixes user data with binaries.
Now: %APPDATA%/PDFBatchPrinterPro/config.json on Windows,
~/.pdfbatchprinterpro/config.json elsewhere.
Old locations are migrated automatically on first load.
"""
from __future__ import annotations

import json
import os
import sys

APP_DIR_NAME = "PDFBatchPrinterPro"
CONFIG_FILENAME = "config.json"


def get_app_data_dir() -> str:
    """User-writable dir for config/logs (never Program Files)."""
    if sys.platform == "win32":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
        return os.path.join(base, APP_DIR_NAME)
    return os.path.join(os.path.expanduser("~"), ".pdfbatchprinterpro")


def _legacy_candidates() -> list[str]:
    cands: list[str] = []
    try:
        if getattr(sys, "frozen", False):
            cands.append(os.path.join(os.path.dirname(sys.executable), CONFIG_FILENAME))
    except Exception:
        pass
    cands.append(os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))), CONFIG_FILENAME))
    return cands


CONFIG_FILE = os.path.join(get_app_data_dir(), CONFIG_FILENAME)


def load_config() -> dict:
    path = CONFIG_FILE
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data if isinstance(data, dict) else {}
        except Exception as exc:
            print(f"Error loading config: {exc}")
            return {}
    # Migrate from legacy location once
    for legacy in _legacy_candidates():
        if legacy != path and os.path.exists(legacy):
            try:
                with open(legacy, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict) and data:
                    save_config(data)
                    return data
            except Exception:
                continue
    return {}


def save_config(cfg: dict) -> None:
    try:
        os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
        tmp = CONFIG_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
        os.replace(tmp, CONFIG_FILE)
    except Exception as exc:
        print(f"Error saving config: {exc}")
