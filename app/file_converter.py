"""
File converter — converts Word, PowerPoint, Excel, and images to PDF
using Win32 COM automation (for Office) and Pillow (for images).

Converted files are stored in a temporary directory that is automatically
cleaned up on program exit.
"""

import os
import tempfile
import uuid
import atexit
import shutil
from typing import Optional

# Temp directory for converted files
_TEMP_DIR: Optional[str] = None


def _get_temp_dir() -> str:
    """Lazily create and return the temp directory."""
    global _TEMP_DIR
    if _TEMP_DIR is None or not os.path.exists(_TEMP_DIR):
        _TEMP_DIR = tempfile.mkdtemp(prefix="pdf_batch_printer_")
    return _TEMP_DIR


def _unique_pdf_path(original_path: str) -> str:
    """Generate a unique temporary PDF file path."""
    base = os.path.splitext(os.path.basename(original_path))[0]
    uid = uuid.uuid4().hex[:8]
    return os.path.join(_get_temp_dir(), f"{base}_{uid}.pdf")


# ═════════════════════════════════════════════════════════════════════════
#  Public API
# ═════════════════════════════════════════════════════════════════════════

def convert_to_pdf(file_path: str) -> tuple[str, bool]:
    """
    Convert *file_path* to PDF.

    Returns
    -------
    (pdf_path, is_temporary)
        For PDF files: ``(original_path, False)``.
        For other types: ``(temp_pdf_path, True)``.

    Raises
    ------
    ValueError
        If the file extension is not supported.
    RuntimeError
        If the conversion fails.
    """
    ext = os.path.splitext(file_path)[1].lower()
    abs_path = os.path.abspath(file_path)

    if ext == ".pdf":
        return abs_path, False
    elif ext in (".doc", ".docx", ".rtf"):
        return _convert_word(abs_path), True
    elif ext in (".ppt", ".pptx"):
        return _convert_powerpoint(abs_path), True
    elif ext in (".xls", ".xlsx"):
        return _convert_excel(abs_path), True
    elif ext in (".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".gif", ".webp"):
        return _convert_image(abs_path), True
    else:
        raise ValueError(f"Định dạng không được hỗ trợ: {ext}")


# ═════════════════════════════════════════════════════════════════════════
#  Office converters (COM automation — requires Office installed)
#  Shared app instances are reused across files in one session (guarded by
#  a lock, since COM STA objects are not thread-safe). Each instance quits
#  at exit or after CONVERTER_TTL_SECONDS of inactivity. Every conversion
#  also runs under a timeout so a hung Office dialog can't freeze the app.
# ═════════════════════════════════════════════════════════════════════════

import threading as _threading
import time as _time

CONVERT_TIMEOUT_SECONDS = 180
CONVERTER_TTL_SECONDS = 300

_office_lock = _threading.Lock()
_office_apps: dict[str, tuple] = {}  # kind -> (app, last_used_monotonic)


def _touch(kind: str, app) -> None:
    _office_apps[kind] = (app, _time.monotonic())


def _get_shared_app(kind: str):
    """Return a cached COM app instance, creating it on first use."""
    entry = _office_apps.get(kind)
    if entry is not None:
        app, last_used = entry
        if _time.monotonic() - last_used > CONVERTER_TTL_SECONDS:
            _drop_shared_app(kind)
        else:
            try:
                # Liveness probe (cheap property access); dead object raises.
                _ = app.Visible if kind != "PowerPoint" else app.Presentations.Count
                _touch(kind, app)
                return app
            except Exception:
                _drop_shared_app(kind)
    import win32com.client
    if kind == "Word":
        app = win32com.client.Dispatch("Word.Application")
        app.Visible = False
        app.DisplayAlerts = False
    elif kind == "PowerPoint":
        app = win32com.client.Dispatch("PowerPoint.Application")
    else:
        app = win32com.client.Dispatch("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
    _touch(kind, app)
    return app


def _drop_shared_app(kind: str) -> None:
    entry = _office_apps.pop(kind, None)
    if entry is not None:
        app, _ = entry
        try:
            app.Quit()
        except Exception:
            pass


def _close_shared_apps() -> None:
    for kind in list(_office_apps.keys()):
        _drop_shared_app(kind)


atexit.register(_close_shared_apps)


def _run_with_timeout(func, timeout: float, what: str):
    """Run *func* in a worker thread; raise RuntimeError on timeout."""
    import concurrent.futures
    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as pool:
        fut = pool.submit(func)
        try:
            return fut.result(timeout=timeout)
        except concurrent.futures.TimeoutError:
            raise RuntimeError(f"Hết thời gian chờ ({int(timeout)}s) khi chuyển đổi {what} — Office có thể đang hiện hộp thoại.")


def _convert_word(word_path: str) -> str:
    import pythoncom

    pdf_path = _unique_pdf_path(word_path)

    def _do():
        pythoncom.CoInitialize()
        try:
            with _office_lock:
                word = _get_shared_app("Word")
                doc = word.Documents.Open(word_path)
                try:
                    doc.SaveAs(pdf_path, FileFormat=17)  # wdFormatPDF
                finally:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
            return pdf_path
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    try:
        return _run_with_timeout(_do, CONVERT_TIMEOUT_SECONDS, os.path.basename(word_path))
    except RuntimeError:
        raise
    except Exception as exc:
        _drop_shared_app("Word")
        raise RuntimeError(f"Lỗi chuyển đổi Word → PDF: {exc}")


def _convert_powerpoint(ppt_path: str) -> str:
    import pythoncom

    pdf_path = _unique_pdf_path(ppt_path)

    def _do():
        pythoncom.CoInitialize()
        try:
            with _office_lock:
                ppt = _get_shared_app("PowerPoint")
                presentation = ppt.Presentations.Open(ppt_path, WithWindow=False)
                try:
                    presentation.SaveAs(pdf_path, FileFormat=32)  # ppSaveAsPDF
                finally:
                    try:
                        presentation.Close()
                    except Exception:
                        pass
            return pdf_path
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    try:
        return _run_with_timeout(_do, CONVERT_TIMEOUT_SECONDS, os.path.basename(ppt_path))
    except RuntimeError:
        raise
    except Exception as exc:
        _drop_shared_app("PowerPoint")
        raise RuntimeError(f"Lỗi chuyển đổi PowerPoint → PDF: {exc}")


def _convert_excel(excel_path: str) -> str:
    import pythoncom

    pdf_path = _unique_pdf_path(excel_path)

    def _do():
        pythoncom.CoInitialize()
        try:
            with _office_lock:
                excel = _get_shared_app("Excel")
                wb = excel.Workbooks.Open(excel_path)
                try:
                    wb.ExportAsFixedFormat(0, pdf_path)  # xlTypePDF
                finally:
                    try:
                        wb.Close(False)
                    except Exception:
                        pass
            return pdf_path
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    try:
        return _run_with_timeout(_do, CONVERT_TIMEOUT_SECONDS, os.path.basename(excel_path))
    except RuntimeError:
        raise
    except Exception as exc:
        _drop_shared_app("Excel")
        raise RuntimeError(f"Lỗi chuyển đổi Excel → PDF: {exc}")


def convert_many_to_pdf(file_paths: list[str]) -> dict[str, str]:
    """Batch-convert Office/image files reusing shared Office instances.

    Returns {original_path: pdf_path}. Raises on first failure.
    """
    out: dict[str, str] = {}
    for p in file_paths:
        pdf_path, _ = convert_to_pdf(p)
        out[p] = pdf_path
    return out


# ═════════════════════════════════════════════════════════════════════════
#  Image converter (Pillow — no Office needed)
# ═════════════════════════════════════════════════════════════════════════

def _convert_image(image_path: str) -> str:
    pdf_path = _unique_pdf_path(image_path)
    try:
        import pymupdf as fitz
        img_doc = fitz.open(image_path)
        try:
            pdf_bytes = img_doc.convert_to_pdf()
        finally:
            img_doc.close()
        pdf_doc = fitz.open("pdf", pdf_bytes)
        try:
            pdf_doc.save(pdf_path)
        finally:
            pdf_doc.close()
        return pdf_path
    except Exception:
        # Fallback to Pillow
        try:
            from PIL import Image
            with Image.open(image_path) as img:
                if img.mode != "RGB":
                    img = img.convert("RGB")
                    img.save(pdf_path, "PDF", resolution=300)
                else:
                    img.save(pdf_path, "PDF", resolution=300)
            return pdf_path
        except Exception as exc:
            raise RuntimeError(f"Lỗi chuyển đổi ảnh → PDF: {exc}")


# ═════════════════════════════════════════════════════════════════════════
#  Cleanup
# ═════════════════════════════════════════════════════════════════════════

def cleanup_temp():
    """Remove all temporary converted files."""
    global _TEMP_DIR
    try:
        _close_shared_apps()
    except Exception:
        pass
    if _TEMP_DIR and os.path.exists(_TEMP_DIR):
        shutil.rmtree(_TEMP_DIR, ignore_errors=True)
        _TEMP_DIR = None


atexit.register(cleanup_temp)
