# -*- coding: utf-8 -*-
"""Resource path helper shared by GUI modules."""
from __future__ import annotations

import os
import sys


def get_resource_path(relative_path: str) -> str:
    """Get absolute path to resource, works for dev and for PyInstaller bundle."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        # __file__ = app/gui/paths.py → lên 3 cấp mới tới project root
        base_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return os.path.join(base_path, relative_path)
