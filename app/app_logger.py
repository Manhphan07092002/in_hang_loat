# -*- coding: utf-8 -*-
"""File logging with rotation: %APPDATA%/PDFBatchPrinterPro/logs/app.log.

Keeps 3 files x 1 MB. GUI mirrors _log() lines here so print history
survives restarts and bug reports can attach the log.
"""
from __future__ import annotations

import logging
import os
import sys
from logging.handlers import RotatingFileHandler

_LOGGER_NAME = "pdfbatchprinter"
_handler_added = False


def get_log_dir() -> str:
    if sys.platform == "win32":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
        return os.path.join(base, "PDFBatchPrinterPro", "logs")
    return os.path.join(os.path.expanduser("~"), ".pdfbatchprinterpro", "logs")


def get_logger() -> logging.Logger:
    global _handler_added
    logger = logging.getLogger(_LOGGER_NAME)
    logger.setLevel(logging.INFO)
    if not _handler_added:
        try:
            os.makedirs(get_log_dir(), exist_ok=True)
            fh = RotatingFileHandler(
                os.path.join(get_log_dir(), "app.log"),
                maxBytes=1 * 1024 * 1024, backupCount=3, encoding="utf-8",
            )
            fh.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s", "%H:%M:%S"))
            logger.addHandler(fh)
        except Exception as exc:
            print(f"Cannot init file log: {exc}")
        _handler_added = True
    return logger


def log_info(msg: str) -> None:
    try:
        get_logger().info(msg)
    except Exception:
        pass


def log_error(msg: str) -> None:
    try:
        get_logger().error(msg)
    except Exception:
        pass
