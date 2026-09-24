"""
Utility functions for PDF Batch Printer.
"""

import re
from datetime import datetime


def parse_page_range(text: str, max_pages: int) -> list[int]:
    """
    Parse a page range string into a sorted list of unique 1-indexed page
    numbers.

    Supported formats:
        "1"          → [1]
        "1-5"        → [1, 2, 3, 4, 5]
        "1,3,5"      → [1, 3, 5]
        "1-5,8-12"   → [1, 2, 3, 4, 5, 8, 9, 10, 11, 12]
        "2,4,7-10"   → [2, 4, 7, 8, 9, 10]

    Args:
        text: Page range string.
        max_pages: Maximum valid page number.

    Returns:
        Sorted list of 1-indexed page numbers.

    Raises:
        ValueError: If the range is invalid.
    """
    if not text or not text.strip():
        raise ValueError("Phạm vi trang không được để trống.")

    text = text.strip()

    # Only digits, spaces, commas, and hyphens allowed
    if not re.match(r'^[\d\s,\-]+$', text):
        raise ValueError("Phạm vi trang chứa ký tự không hợp lệ.")

    pages: set[int] = set()
    parts = text.split(',')

    for part in parts:
        part = part.strip()
        if not part:
            continue

        if '-' in part:
            range_parts = part.split('-')
            if len(range_parts) != 2 or not range_parts[0].strip() or not range_parts[1].strip():
                raise ValueError(f"Phạm vi không hợp lệ: '{part}'")
            try:
                start = int(range_parts[0].strip())
                end = int(range_parts[1].strip())
            except ValueError:
                raise ValueError(f"Số trang không hợp lệ trong: '{part}'")
            if start <= 0 or end <= 0:
                raise ValueError("Số trang phải lớn hơn 0.")
            if start > end:
                raise ValueError(f"Phạm vi không hợp lệ: {start} > {end}")
            if end > max_pages:
                raise ValueError(
                    f"Trang {end} vượt quá số trang tài liệu ({max_pages})."
                )
            pages.update(range(start, end + 1))
        else:
            try:
                page = int(part)
            except ValueError:
                raise ValueError(f"Số trang không hợp lệ: '{part}'")
            if page <= 0:
                raise ValueError("Số trang phải lớn hơn 0.")
            if page > max_pages:
                raise ValueError(
                    f"Trang {page} vượt quá số trang tài liệu ({max_pages})."
                )
            pages.add(page)

    if not pages:
        raise ValueError("Không có trang nào được chọn.")

    return sorted(pages)


def format_file_size(size_bytes: int) -> str:
    """Return a human-readable file size string."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.1f} GB"


def get_timestamp() -> str:
    """Return the current time as ``HH:MM:SS``."""
    return datetime.now().strftime("%H:%M:%S")


# ── Per-file page selection ────────────────────────────────────────────
# mode: "all" | "custom" | "odd" | "even". value: custom range text.

def resolve_page_selection(mode: str, value: str, max_pages: int) -> list[int]:
    """Resolve a per-file page selection to 0-based page indices.

    Raises ValueError with Vietnamese message on invalid input.
    """
    if max_pages <= 0:
        raise ValueError("Tài liệu chưa xác định được số trang.")
    mode = (mode or "all").strip().lower()
    if mode in ("all", "", "tat ca"):
        return list(range(max_pages))
    if mode == "odd":
        return [p - 1 for p in range(1, max_pages + 1) if p % 2 == 1]
    if mode == "even":
        return [p - 1 for p in range(1, max_pages + 1) if p % 2 == 0]
    if mode == "custom":
        p1 = parse_page_range(value or "", max_pages)
        return [p - 1 for p in p1]
    raise ValueError(f"Chế độ trang in không hợp lệ: '{mode}'.")


def describe_pages(mode: str, value: str, max_len: int = 18) -> str:
    """Short display string for the queue 'Trang in' column."""
    mode = (mode or "all").strip().lower()
    if mode in ("all", "", "tat ca"):
        return "Tất cả"
    if mode == "odd":
        return "Lẻ"
    if mode == "even":
        return "Chẵn"
    text = (value or "").strip()
    if not text:
        return "Tất cả"
    return text if len(text) <= max_len else text[: max_len - 1] + "…"
