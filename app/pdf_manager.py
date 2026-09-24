"""
File management — reading, validating, converting, and rendering documents.

Supports PDF natively; Word, PowerPoint, Excel, and images are converted to
temporary PDFs via :mod:`app.file_converter` so that preview and printing use
a single code-path.
"""

import os
from typing import Optional

import pymupdf as fitz
from PIL import Image

from app.settings import PREVIEW_DPI, PREVIEW_MAX_CACHE, SUPPORTED_EXTENSIONS
from app.file_converter import convert_to_pdf


# ═════════════════════════════════════════════════════════════════════════
#  FileInfo — one entry in the print queue
# ═════════════════════════════════════════════════════════════════════════

class FileInfo:
    """
    Metadata for a single file in the print queue.

    Attributes
    ----------
    original_path : str
        The path the user selected (may be .docx, .pptx, etc.).
    pdf_path : str
        Path to the PDF used for preview and printing.
        Same as *original_path* for PDFs; a temp file for converted docs.
    is_converted : bool
        ``True`` if *pdf_path* is a temporary conversion result.
    is_converted_ready : bool
        ``True`` if *pdf_path* is generated and ready to read/print.
    """

    __slots__ = (
        "original_path", "pdf_path", "filename", "file_type",
        "is_converted", "is_converted_ready", "page_count", "file_size",
        "copies", "page_range_text", "page_mode", "status",
    )

    def __init__(
        self,
        original_path: str,
        pdf_path: Optional[str] = None,
        page_count: int = 1,
        file_size: int = 0,
        file_type: str = "PDF",
        is_converted: bool = False,
        is_converted_ready: bool = True,
        page_mode: str = "all",
        page_range_text: str = "",
    ):
        self.original_path = original_path
        self.pdf_path = pdf_path if pdf_path else original_path
        self.filename = os.path.basename(original_path)
        self.file_type = file_type
        self.is_converted = is_converted
        self.is_converted_ready = is_converted_ready
        self.page_count = page_count
        self.file_size = file_size
        self.copies: int = 1
        self.page_mode: str = page_mode or "all"  # all | custom | odd | even
        self.page_range_text: str = page_range_text  # custom value, e.g. "1,3,5"
        self.status: str = "Chờ in"

    def resolve_pages(self) -> list[int]:
        """0-based page indices for this file (validates custom ranges)."""
        from app.utils import resolve_page_selection
        return resolve_page_selection(self.page_mode, self.page_range_text, self.page_count)

    def pages_display(self) -> str:
        """Short string for the queue 'Trang in' column."""
        from app.utils import describe_pages
        return describe_pages(self.page_mode, self.page_range_text)


# ═════════════════════════════════════════════════════════════════════════
#  PDFManager
# ═════════════════════════════════════════════════════════════════════════

class PDFManager:
    """
    Manages documents — loading, validation, conversion, rendering, and
    caching with high-performance async capabilities.
    """

    def __init__(self):
        self._thumbnail_cache: dict[tuple, Image.Image] = {}
        self._cache_bytes: int = 0
        self._max_cache: int = PREVIEW_MAX_CACHE
        self._max_cache_bytes: int = 256 * 1024 * 1024  # 256 MB

    # ── Single-file operations ───────────────────────────────────────

    @classmethod
    def quick_inspect(cls, path: str) -> FileInfo:
        """
        Fast non-blocking inspection of file metadata (< 1ms).
        Does not convert non-PDF files immediately.
        """
        abs_path = os.path.abspath(path)
        ext = os.path.splitext(abs_path)[1].lower()

        if ext not in SUPPORTED_EXTENSIONS:
            raise ValueError(f"Định dạng không hỗ trợ: {ext}")

        file_type = SUPPORTED_EXTENSIONS[ext]
        try:
            file_size = os.path.getsize(abs_path)
        except OSError:
            file_size = 0

        if ext == ".pdf":
            page_count = cls.get_page_count(abs_path)
            return FileInfo(
                original_path=abs_path,
                pdf_path=abs_path,
                page_count=page_count,
                file_size=file_size,
                file_type=file_type,
                is_converted=False,
                is_converted_ready=True,
            )
        elif file_type == "Ảnh":
            # Any image type from SUPPORTED_EXTENSIONS that is not PDF/Office
            return FileInfo(
                original_path=abs_path,
                pdf_path=None,
                page_count=1,
                file_size=file_size,
                file_type=file_type,
                is_converted=True,
                is_converted_ready=False,
            )
        else:
            # Office documents (Word, Excel, PowerPoint)
            return FileInfo(
                original_path=abs_path,
                pdf_path=None,
                page_count=1,
                file_size=file_size,
                file_type=file_type,
                is_converted=True,
                is_converted_ready=False,
            )

    @classmethod
    def ensure_pdf(cls, info: FileInfo) -> str:
        """
        Ensure that *info* has a valid converted PDF ready for preview or printing.
        Converts non-PDF documents on demand.
        """
        if info.is_converted_ready and info.pdf_path and os.path.exists(info.pdf_path):
            return info.pdf_path

        ext = os.path.splitext(info.original_path)[1].lower()
        if ext == ".pdf":
            info.pdf_path = info.original_path
            info.is_converted = False
            info.is_converted_ready = True
            info.page_count = cls.get_page_count(info.pdf_path)
            return info.pdf_path

        # Convert to PDF
        pdf_path, is_temp = convert_to_pdf(info.original_path)
        info.pdf_path = pdf_path
        info.is_converted = is_temp
        info.is_converted_ready = True
        info.page_count = cls.get_page_count(pdf_path)
        return info.pdf_path

    def add_file(self, path: str, lazy: bool = True) -> FileInfo:
        """
        Add a single file of any supported type to the queue.
        When *lazy* is True, non-PDF files are inspected in <1ms and converted on-demand.
        """
        if lazy:
            return self.quick_inspect(path)

        abs_path = os.path.abspath(path)
        ext = os.path.splitext(abs_path)[1].lower()

        if ext not in SUPPORTED_EXTENSIONS:
            raise ValueError(f"Định dạng không hỗ trợ: {ext}")

        file_type = SUPPORTED_EXTENSIONS[ext]
        file_size = os.path.getsize(abs_path)

        # Convert to PDF (pass-through for .pdf)
        pdf_path, is_temp = convert_to_pdf(abs_path)
        page_count = self.get_page_count(pdf_path)

        return FileInfo(
            original_path=abs_path,
            pdf_path=pdf_path,
            page_count=page_count,
            file_size=file_size,
            file_type=file_type,
            is_converted=is_temp,
            is_converted_ready=True,
        )

    # ── Batch operations ─────────────────────────────────────────────

    def add_files(self, paths: list[str], lazy: bool = True) -> tuple[list[FileInfo], list[str]]:
        """
        Add multiple files. Returns ``(successes, errors)`` where *errors*
        is a list of human-readable error strings.
        """
        results: list[FileInfo] = []
        errors: list[str] = []
        for p in paths:
            ext = os.path.splitext(p)[1].lower()
            if ext not in SUPPORTED_EXTENSIONS:
                continue
            try:
                info = self.add_file(p, lazy=lazy)
                results.append(info)
            except Exception as exc:
                errors.append(f"{os.path.basename(p)}: {exc}")
        return results, errors

    def scan_folder(
        self, folder: str, recursive: bool = False,
    ) -> list[str]:
        """Return sorted list of supported file paths in *folder* using high-speed os.scandir."""
        found: list[str] = []
        valid_exts = set(SUPPORTED_EXTENSIONS.keys())

        def _scan(target_dir):
            try:
                with os.scandir(target_dir) as entries:
                    for entry in entries:
                        try:
                            if entry.is_file(follow_symlinks=False):
                                ext = os.path.splitext(entry.name)[1].lower()
                                if ext in valid_exts:
                                    found.append(entry.path)
                            elif recursive and entry.is_dir(follow_symlinks=False):
                                _scan(entry.path)
                        except OSError:
                            continue
            except OSError:
                pass

        _scan(folder)
        found.sort(key=str.lower)
        return found

    # ── Page queries ─────────────────────────────────────────────────

    @staticmethod
    def get_page_count(path: str) -> int:
        doc = None
        try:
            doc = fitz.open(path)
            return doc.page_count
        except Exception:
            return 0
        finally:
            if doc is not None:
                try:
                    doc.close()
                except Exception:
                    pass

    @staticmethod
    def get_page_size(path: str, page_num: int = 0) -> tuple[float, float]:
        """``(width, height)`` in points."""
        doc = None
        try:
            doc = fitz.open(path)
            if page_num < 0 or page_num >= doc.page_count:
                return (595.0, 842.0)
            rect = doc[page_num].rect
            return (rect.width, rect.height)
        except Exception:
            return (595.0, 842.0)
        finally:
            if doc is not None:
                try:
                    doc.close()
                except Exception:
                    pass

    @classmethod
    def is_landscape(cls, path: str, page_num: int = 0) -> bool:
        w, h = cls.get_page_size(path, page_num)
        return w > h

    # ── Rendering ────────────────────────────────────────────────────

    def render_page(
        self,
        path: str,
        page_num: int,
        width: Optional[int] = None,
        height: Optional[int] = None,
        dpi: Optional[int] = None,
    ) -> Optional[Image.Image]:
        cache_key = (path, page_num, width, height, dpi)
        cached = self._thumbnail_cache.get(cache_key)
        if cached is not None:
            return cached

        doc = None
        try:
            doc = fitz.open(path)
            if page_num < 0 or page_num >= doc.page_count:
                return None

            page = doc[page_num]
            rect = page.rect

            if width and height:
                s = min(width / rect.width, height / rect.height)
                s = max(0.1, min(4.0, s))
                mat = fitz.Matrix(s, s)
            else:
                s = (dpi or PREVIEW_DPI) / 72.0
                s = max(0.1, min(4.0, s))
                mat = fitz.Matrix(s, s)

            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            # Rough byte estimate for cache eviction
            img_bytes = pix.width * pix.height * 3
            while (
                self._thumbnail_cache
                and (
                    len(self._thumbnail_cache) >= self._max_cache
                    or self._cache_bytes + img_bytes > self._max_cache_bytes
                )
            ):
                old_key = next(iter(self._thumbnail_cache))
                old_img = self._thumbnail_cache.pop(old_key)
                try:
                    self._cache_bytes -= old_img.width * old_img.height * 3
                except Exception:
                    pass
                if self._cache_bytes < 0:
                    self._cache_bytes = 0
            self._thumbnail_cache[cache_key] = img
            self._cache_bytes += img_bytes
            return img
        except Exception:
            return None
        finally:
            if doc is not None:
                try:
                    doc.close()
                except Exception:
                    pass

    # ── Blank page analysis & page filtering ─────────────────────────

    @staticmethod
    def is_blank_page(page, threshold: float = 0.002) -> bool:
        """
        Determine if a PyMuPDF page is blank.
        Returns True if there is no significant text, drawings, or non-white pixels.
        """
        try:
            # 1. Text check
            text = page.get_text().strip()
            if text:
                return False

            # 2. Vector drawings and images check
            drawings = page.get_drawings()
            if drawings:
                return False
            images = page.get_images()
            if images:
                return False

            # 3. Fast pixel scan check on low-res pixmap
            pix = page.get_pixmap(matrix=fitz.Matrix(0.2, 0.2), alpha=False)
            samples = pix.samples
            if not samples:
                return True
            # Sample every 3rd byte (RGB) to check for non-white pixels
            non_white = sum(1 for b in samples[::3] if b < 240)
            ratio = non_white / (pix.width * pix.height)
            return ratio < threshold
        except Exception:
            return False

    @classmethod
    def filter_pages(cls, path: str, pages: list[int], remove_blanks: bool = True) -> tuple[list[int], list[int]]:
        """
        Filter a list of page indices (0-based).
        Returns (active_pages, skipped_blank_pages).
        """
        if not remove_blanks:
            return pages, []

        doc = None
        try:
            doc = fitz.open(path)
            active = []
            skipped = []
            for p_idx in pages:
                if 0 <= p_idx < doc.page_count:
                    page = doc[p_idx]
                    if cls.is_blank_page(page):
                        skipped.append(p_idx)
                    else:
                        active.append(p_idx)
            # If all pages were marked blank, keep at least page 0 to avoid printing nothing
            if not active and pages:
                return [pages[0]], []
            return active, skipped
        except Exception:
            return pages, []
        finally:
            if doc is not None:
                try:
                    doc.close()
                except Exception:
                    pass

    # ── Separator / Slip Sheet Generation ────────────────────────────

    @staticmethod
    def create_separator_sheet(
        filename: str,
        page_count: int,
        copies: int,
        printer_name: str,
        doc_index: int = 1,
    ) -> str:
        """
        Create a temporary 1-page PDF separator slip-sheet.
        Returns the absolute file path of the temporary PDF.
        """
        import tempfile
        import datetime

        doc = fitz.open()
        page = doc.new_page(width=595, height=842)

        now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        # Decorative header bar
        page.draw_rect(fitz.Rect(40, 40, 555, 110), color=(0.14, 0.38, 0.92), fill=(0.14, 0.38, 0.92))
        page.insert_text((60, 80), "TRANG PHAN CACH TAI LIEU", fontsize=18, color=(1, 1, 1))
        page.insert_text((60, 100), "PDF BATCH PRINTER PRO — SEPARATOR SLIP SHEET", fontsize=9, color=(0.85, 0.92, 1.0))

        # Main Info Box (Card)
        card_rect = fitz.Rect(40, 130, 555, 780)
        page.draw_rect(card_rect, color=(0.8, 0.85, 0.92), fill=(0.97, 0.98, 1.0), width=1.5)

        # Content lines
        y = 180
        page.insert_text((70, y), f"TEP TIN SO #{doc_index}", fontsize=15, color=(0.14, 0.38, 0.92))
        y += 40
        page.draw_line(fitz.Point(70, y - 10), fitz.Point(525, y - 10), color=(0.85, 0.88, 0.92), width=1)

        # File details
        clean_fn = filename.encode("ascii", errors="replace").decode("ascii")
        clean_pname = printer_name.encode("ascii", errors="replace").decode("ascii")

        details = [
            ("Ten tep tin:", clean_fn),
            ("So trang tai lieu:", f"{page_count} trang"),
            ("So ban in:", f"{copies} ban"),
            ("May in chi dinh:", clean_pname),
            ("Thoi gian in:", now_str),
            ("Trang thai:", "SAN SANG IN TAP TIEP THEO"),
        ]

        for label, val in details:
            page.insert_text((70, y), label, fontsize=12, color=(0.3, 0.35, 0.45))
            disp_val = val if len(val) <= 45 else val[:42] + "..."
            page.insert_text((220, y), disp_val, fontsize=12, color=(0.06, 0.09, 0.16))
            y += 35

        # Footer notice inside card
        y += 40
        page.draw_rect(fitz.Rect(70, y, 525, y + 60), color=(0.95, 0.75, 0.3), fill=(1.0, 0.98, 0.9))
        page.insert_text((85, y + 26), "Luu y: Day la to bia phan cach tu dong, dung de phan chia cac tap ho so.", fontsize=10, color=(0.6, 0.4, 0.0))
        page.insert_text((85, y + 46), "Vui long giu to nay giua cac bo tai lieu de tranh lan lon ho so.", fontsize=9, color=(0.6, 0.4, 0.0))

        temp_dir = tempfile.gettempdir()
        sep_pdf_path = os.path.join(temp_dir, f"sep_sheet_{datetime.datetime.now():%Y%m%d%H%M%S%f}.pdf")
        doc.save(sep_pdf_path)
        doc.close()
        return sep_pdf_path

    # ── Cache ────────────────────────────────────────────────────────

    def clear_cache(self):
        self._thumbnail_cache.clear()
        self._cache_bytes = 0
