"""
Windows printer management — enumerate printers, configure DEVMODE, and
send print jobs through the Windows GDI print path.

Flow:
    PDF page  →  PyMuPDF render at printer DPI  →  PIL Image
    →  ImageWin.Dib  →  GDI DC (with DEVMODE)  →  printer driver
"""

import os
from typing import Optional, Callable

import win32print
import win32gui
import win32ui
import win32con
import pymupdf as fitz
from PIL import Image, ImageWin

from app.settings import PAPER_SIZES


class PrinterManager:
    """Manages Windows printers and performs actual printing."""

    # ── Printer enumeration ──────────────────────────────────────────────

    @staticmethod
    def get_printers() -> list[str]:
        """Return a list of available printer names."""
        try:
            flags = (
                win32print.PRINTER_ENUM_LOCAL
                | win32print.PRINTER_ENUM_CONNECTIONS
            )
            printers = win32print.EnumPrinters(flags, None, 2)
            return [p["pPrinterName"] for p in printers]
        except Exception:
            return []

    @staticmethod
    def get_default_printer() -> str:
        try:
            return win32print.GetDefaultPrinter()
        except Exception:
            return ""

    # ── Printer status ───────────────────────────────────────────────────

    @staticmethod
    def get_printer_status(printer_name: str) -> tuple[str, bool]:
        """Return ``(display_text, is_ready)``."""
        try:
            handle = win32print.OpenPrinter(printer_name)
            try:
                info = win32print.GetPrinter(handle, 2)
                status = info.get("Status", 0)
                if status == 0:
                    return ("✓ Máy in sẵn sàng", True)

                _MAP = {
                    getattr(win32print, "PRINTER_STATUS_PAUSED", 0x1): ("⚠ Tạm dừng", False),
                    getattr(win32print, "PRINTER_STATUS_ERROR", 0x2): ("⚠ Lỗi", False),
                    getattr(win32print, "PRINTER_STATUS_PAPER_JAM", 0x8): ("⚠ Kẹt giấy", False),
                    getattr(win32print, "PRINTER_STATUS_PAPER_OUT", 0x10): ("⚠ Hết giấy", False),
                    getattr(win32print, "PRINTER_STATUS_OFFLINE", 0x80): ("⚠ Offline", False),
                    getattr(win32print, "PRINTER_STATUS_IO_ACTIVE", 0x100): ("Đang hoạt động", True),
                    getattr(win32print, "PRINTER_STATUS_BUSY", 0x200): ("Bận", True),
                    getattr(win32print, "PRINTER_STATUS_PRINTING", 0x400): ("Đang in", True),
                    getattr(win32print, "PRINTER_STATUS_WARMING_UP", 0x800): ("Đang khởi động", True),
                }
                for code, (text, ready) in _MAP.items():
                    if status & code:
                        return (text, ready)

                return ("✓ Máy in sẵn sàng", True)
            finally:
                win32print.ClosePrinter(handle)
        except Exception as exc:
            return (f"⚠ Không thể kiểm tra: {exc}", False)

    # ── Duplex capability ────────────────────────────────────────────────

    @staticmethod
    def supports_duplex(printer_name: str) -> bool:
        """Check whether the printer driver advertises duplex support."""
        try:
            dup = win32print.DeviceCapabilities(printer_name, "", win32con.DC_DUPLEX)
            if dup == 1:
                return True
        except Exception:
            pass

        try:
            handle = win32print.OpenPrinter(printer_name)
            try:
                devmode = win32print.GetPrinter(handle, 2).get("pDevMode")
                if devmode is None:
                    return False
                return bool(devmode.Fields & win32con.DM_DUPLEX)
            finally:
                win32print.ClosePrinter(handle)
        except Exception:
            return False

    # ── Actual printing ──────────────────────────────────────────────────

    @staticmethod
    def print_pdf(
        printer_name: str,
        pdf_path: str,
        pages: list[int],
        copies: int,
        paper_size_name: str,
        orientation_value: int,
        duplex_value: int,
        fit_to_page: bool,
        cancel_event,
        page_callback: Optional[Callable[[int, int], None]] = None,
        binding_margin_mm: float = 0.0,
        reverse_order: bool = False,
    ):
        """
        Print selected pages of a PDF to the given Windows printer.

        Parameters
        ----------
        printer_name : str
            The Windows printer name.
        pdf_path : str
            Absolute path to the PDF.
        pages : list[int]
            **0-indexed** page numbers to print.
        copies : int
            Number of copies (set in DEVMODE so the driver handles collation).
        paper_size_name : str
            Key into ``PAPER_SIZES`` (e.g. ``"A4"``).
        orientation_value : int
            ``win32con.DMORIENT_PORTRAIT`` or ``DMORIENT_LANDSCAPE``.
        duplex_value : int
            ``DMDUP_SIMPLEX``, ``DMDUP_VERTICAL``, or ``DMDUP_HORIZONTAL``.
        fit_to_page : bool
            Scale each page to fit the printable area (aspect-ratio preserved).
        cancel_event : threading.Event
            Checked between pages — set it to abort.
        page_callback : callable, optional
            ``callback(current_1based, total)`` invoked after each page is
            sent to the spooler.
        binding_margin_mm : float
            Offset in millimetres for bookbinding margin (punched holes/ring binding).
        reverse_order : bool
            If True, prints from last page to first page.

        Raises
        ------
        Exception
            On any printing failure (caller should catch and report).
        """
        # ── 1. Obtain and configure DEVMODE ──────────────────────────────
        if not printer_name:
            raise ValueError("Chưa chọn máy in.")
        if not pdf_path or not os.path.exists(pdf_path):
            raise FileNotFoundError(f"Không tìm thấy file PDF: {pdf_path}")
        pages = [p for p in (pages or []) if isinstance(p, int) and p >= 0]
        if not pages:
            raise ValueError("Danh sách trang in rỗng.")
        copies = max(1, min(999, int(copies or 1)))

        handle = win32print.OpenPrinter(printer_name)
        devmode = None
        try:
            info = win32print.GetPrinter(handle, 2)
            devmode = info.get("pDevMode")
        finally:
            win32print.ClosePrinter(handle)

        if devmode is not None:
            from app.settings import resolve_paper
            kind, pval = resolve_paper(paper_size_name)
            fields = (
                win32con.DM_ORIENTATION
                | win32con.DM_DUPLEX
                | win32con.DM_COPIES
            )
            if kind == "standard":
                devmode.PaperSize = pval
                fields |= win32con.DM_PAPERSIZE
            else:
                # Khổ tùy chỉnh: DMPAPER_USER + kích thước đơn vị 1/10 mm
                w_mm, h_mm = pval
                devmode.PaperSize = win32con.DMPAPER_USER
                try:
                    devmode.PaperWidth = int(round(w_mm * 10))
                    devmode.PaperLength = int(round(h_mm * 10))
                    fields |= (win32con.DM_PAPERSIZE | win32con.DM_PAPERWIDTH | win32con.DM_PAPERLENGTH)
                except Exception:
                    devmode.PaperSize = win32con.DMPAPER_A4
                    fields |= win32con.DM_PAPERSIZE
            devmode.Orientation = orientation_value
            devmode.Duplex = duplex_value
            # NOTE: nhiều driver bỏ qua DM_COPIES nên luôn để Copies=1
            # và lặp bản in bằng phần mềm bên dưới để đảm bảo đủ số bản.
            try:
                devmode.Copies = 1
            except Exception:
                pass
            devmode.Fields |= fields
            hdc = win32gui.CreateDC("WINSPOOL", printer_name, devmode)
        else:
            hdc = win32gui.CreateDC("WINSPOOL", printer_name, None)

        # ── 2. Create PyCDC from HDC ─────────────────────────────────────
        dc = win32ui.CreateDCFromHandle(hdc)
        try:
            printable_w = dc.GetDeviceCaps(win32con.HORZRES)
            printable_h = dc.GetDeviceCaps(win32con.VERTRES)
            dpi_x = dc.GetDeviceCaps(win32con.LOGPIXELSX)
            dpi_y = dc.GetDeviceCaps(win32con.LOGPIXELSY)

            if not dpi_x or dpi_x <= 0:
                dpi_x = 300
            if not dpi_y or dpi_y <= 0:
                dpi_y = 300
            # Giới hạn DPI render để tránh phình RAM với máy in 600/1200dpi.
            # Chất lượng văn bản vẫn tốt ở 300dpi cho khổ A4/A5.
            render_dpi_x = min(int(dpi_x), 300)
            render_dpi_y = min(int(dpi_y), 300)

            # ── 3. Open PDF and print pages ──────────────────────────────
            doc = fitz.open(pdf_path)
            doc_name = os.path.basename(pdf_path)

            try:
                dc.StartDoc(doc_name)
                
                # Apply reverse page order if requested
                pages_to_print = list(reversed(pages)) if reverse_order else list(pages)
                # Lặp số bản bằng phần mềm để không phụ thuộc driver
                all_pages = pages_to_print * copies
                total = len(all_pages)

                # Binding margin in printer device pixels
                margin_px = int((binding_margin_mm / 25.4) * dpi_x) if binding_margin_mm > 0 else 0

                for idx, page_idx in enumerate(all_pages):
                    if cancel_event is not None and cancel_event.is_set():
                        break

                    if page_idx < 0 or page_idx >= doc.page_count:
                        continue
                    page = doc[page_idx]
                    rect = page.rect
                    pw, ph = rect.width, rect.height  # points

                    # Compute render matrix
                    if fit_to_page:
                        # Allow slight reduction if margin offset is enabled so content doesn't clip
                        eff_w = max(100, printable_w - margin_px) if margin_px > 0 else printable_w
                        sx = eff_w / (pw * render_dpi_x / 72.0)
                        sy = printable_h / (ph * render_dpi_y / 72.0)
                        s = min(sx, sy)
                        zx = s * render_dpi_x / 72.0
                        zy = s * render_dpi_y / 72.0
                    else:
                        zx = render_dpi_x / 72.0
                        zy = render_dpi_y / 72.0

                    mat = fitz.Matrix(zx, zy)
                    pix = page.get_pixmap(matrix=mat, alpha=False)
                    img = Image.frombytes(
                        "RGB", (pix.width, pix.height), pix.samples
                    )

                    # ── Draw onto DC ────────────────────────────────────
                    dc.StartPage()
                    try:
                        dib = ImageWin.Dib(img)

                        # Calculate position with binding margin offset
                        base_x = max(0, (printable_w - pix.width) // 2)
                        base_y = max(0, (printable_h - pix.height) // 2)

                        if margin_px > 0:
                            if duplex_value == win32con.DMDUP_VERTICAL:
                                # Front side shifts right, back side shifts left
                                if idx % 2 == 0:
                                    x = min(max(0, printable_w - pix.width), base_x + margin_px // 2)
                                else:
                                    x = max(0, base_x - margin_px // 2)
                            else:
                                x = min(max(0, printable_w - pix.width), base_x + margin_px)
                        else:
                            x = base_x
                        y = base_y

                        dib.draw(
                            dc.GetHandleOutput(),
                            (
                                x,
                                y,
                                x + min(pix.width, printable_w),
                                y + min(pix.height, printable_h),
                            ),
                        )
                    finally:
                        try:
                            del dib
                        except Exception:
                            pass
                        try:
                            img.close()
                        except Exception:
                            pass
                    dc.EndPage()

                    if page_callback is not None:
                        page_callback(idx + 1, total)

                if cancel_event is not None and cancel_event.is_set():
                    try:
                        dc.AbortDoc()
                    except Exception:
                        pass
                else:
                    dc.EndDoc()

            finally:
                doc.close()
        finally:
            try:
                dc.DeleteDC()
            except Exception:
                pass
            try:
                win32gui.DeleteDC(hdc)
            except Exception:
                pass
