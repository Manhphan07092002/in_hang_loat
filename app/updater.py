# -*- coding: utf-8 -*-
"""Self-update checker against GitHub Releases (stdlib only, no new deps).

Flow: check_for_updates() → dialog → download_asset() → user runs installer.
Network failures are swallowed (return has_update=False + error text) so the
app never breaks offline.
"""
from __future__ import annotations

import datetime
import json
import os
import re
import tempfile
import urllib.request
from typing import Callable, Optional

GITHUB_REPO = "Manhphan07092002/in_hang_loat"
API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
CHECK_INTERVAL_HOURS = 24


def parse_version(text: str) -> tuple[int, ...]:
    """'v1.0.0.3' / '1.0.0' → (1, 0, 0, 3). Non-numeric tail is ignored."""
    nums = re.findall(r"\d+", (text or "").strip())
    parts = [int(x) for x in nums[:4]]
    while len(parts) < 4:
        parts.append(0)
    return tuple(parts)


def is_newer(current: str, latest: str) -> bool:
    return parse_version(latest) > parse_version(current)


def get_current_version() -> str:
    try:
        from app.settings import APP_VERSION
        return APP_VERSION
    except Exception:
        return "1.0.0.3"


def should_auto_check(last_check_iso: str = "") -> bool:
    """True nếu chưa từng check hoặc đã quá CHECK_INTERVAL_HOURS."""
    if not last_check_iso:
        return True
    try:
        last = datetime.datetime.fromisoformat(last_check_iso)
        delta = datetime.datetime.now() - last
        return delta.total_seconds() >= CHECK_INTERVAL_HOURS * 3600
    except Exception:
        return True


def check_for_updates(timeout: float = 10.0) -> dict:
    """Return {has_update, current, latest, url, asset_name, notes, error}."""
    current = get_current_version()
    result = {"has_update": False, "current": current, "latest": current,
              "url": "", "asset_name": "", "notes": "", "error": ""}
    try:
        req = urllib.request.Request(API_URL, headers={"User-Agent": "PDFBatchPrinterPro-updater"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.load(resp)
    except Exception as exc:
        result["error"] = f"Không kiểm tra được cập nhật: {exc}"
        return result

    tag = str(data.get("tag_name", "")).strip()
    if not tag:
        return result
    result["latest"] = tag.lstrip("v")
    result["notes"] = str(data.get("body", "") or "")[:2000]
    result["url"] = str(data.get("html_url", "") or "")

    if not is_newer(current, tag):
        return result

    # Ưu tiên file Setup, rồi exe chính
    asset_url, asset_name = "", ""
    for asset in data.get("assets", []) or []:
        name = str(asset.get("name", ""))
        dl = str(asset.get("browser_download_url", ""))
        if not dl:
            continue
        low = name.lower()
        if low.endswith("_setup.exe") or "setup" in low:
            asset_url, asset_name = dl, name
            break
    if not asset_url:
        for asset in data.get("assets", []) or []:
            name = str(asset.get("name", ""))
            dl = str(asset.get("browser_download_url", ""))
            if dl and name.lower().endswith(".exe"):
                asset_url, asset_name = dl, name
                break
    result.update(has_update=True, url=asset_url or result["url"],
                  asset_name=asset_name)
    return result


def download_asset(url: str, dest: Optional[str] = None,
                   progress_cb: Optional[Callable[[int, int], None]] = None,
                   timeout: float = 30.0) -> str:
    """Download *url* to *dest* (default %TEMP%). Return final path."""
    if not url:
        raise ValueError("URL tải về trống.")
    if not dest:
        fname = url.split("?")[0].rstrip("/").split("/")[-1] or "update_setup.exe"
        dest = os.path.join(tempfile.gettempdir(), fname)
    req = urllib.request.Request(url, headers={"User-Agent": "PDFBatchPrinterPro-updater"})
    with urllib.request.urlopen(req, timeout=timeout) as resp, open(dest, "wb") as f:
        total = 0
        try:
            total = int(resp.headers.get("Content-Length", 0) or 0)
        except Exception:
            total = 0
        done = 0
        while True:
            chunk = resp.read(256 * 1024)
            if not chunk:
                break
            f.write(chunk)
            done += len(chunk)
            if progress_cb:
                try:
                    progress_cb(done, total)
                except Exception:
                    pass
    return dest
