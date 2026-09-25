# -*- coding: utf-8 -*-
"""Nhan dien hoa don dien tu (e-invoice) bang phan tich text PDF + ten file.

Nguyen tac (theo spec):
- Uu tien text PDF that; ban scan khong text -> khong ket luan (needs_ocr),
  tuyet doi khong doan mo.
- Ket hop ten file + text + so trang + truong dac trung de cham diem.
- Chi de xuat / tu ap dung duplex khi chac chan (score >= threshold).
- OCR la tuy chon, chi dung khi da cai san (pytesseract) va duoc bat.
"""
from __future__ import annotations

import os
import re
import unicodedata

# Diem ngoung: >= thi coi la hoa don
INVOICE_THRESHOLD = 0.5

# Tu khoa manh (dac trung hoa don VN) — moi hit +0.25
STRONG_KEYWORDS = [
    "hoa don gia tri gia tang",
    "hoa don dien tu",
    "vat invoice",
    "ma so thue",
    "mau so",
    "ky hieu",
    "ma cua co quan thue",
    "tong tien thanh toan",
]

# Tu khoa yeu — moi hit +0.1
WEAK_KEYWORDS = [
    "hoa don",
    "invoice",
    "so hoa don",
    "ngay lap",
    "nguoi ban",
    "nguoi mua",
    "tien hang",
    "thue suat",
    "tien thue",
    "nguoi mua hang",
    "don vi ban",
    "don vi mua",
    "thanh toan",
    "tien thue gtgt",
    "cong tien",
]

# Goi y tu ten file (normalized) — +0.25, KHONG bao gio dung mot minh
FILENAME_PATTERNS = [
    r"hoa[\s_\-]*don",
    r"invoice",
    r"\binv[\s_\-]*\d",
    r"\bhd[\s_\-]*\d",
]

MAX_TEXT_CHARS = 30000


def normalize_vi(text: str) -> str:
    """Lowercase + bo dau tieng Viet de so khop on dinh."""
    text = (text or "").lower()
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    text = text.replace("đ", "d")
    return re.sub(r"\s+", " ", text).strip()


def extract_pdf_text(pdf_path: str, max_chars: int = MAX_TEXT_CHARS) -> tuple[str, int]:
    """Tra ve (text, page_count). Text rong neu scan/ma hoa/loi."""
    try:
        import pymupdf as fitz
        doc = fitz.open(pdf_path)
        try:
            n = doc.page_count
            parts: list[str] = []
            total = 0
            for i in range(n):
                try:
                    t = doc[i].get_text() or ""
                except Exception:
                    t = ""
                if t.strip():
                    parts.append(t)
                    total += len(t)
                    if total >= max_chars:
                        break
            return "".join(parts)[:max_chars], n
        finally:
            try:
                doc.close()
            except Exception:
                pass
    except Exception:
        return "", 0


def try_ocr_first_page(pdf_path: str) -> str:
    """OCR trang 1 neu co pytesseract+Pillow. Tra ve '' khi khong kha dung."""
    try:
        import pymupdf as fitz
        from PIL import Image
        import pytesseract
    except Exception:
        return ""
    try:
        doc = fitz.open(pdf_path)
        try:
            if doc.page_count < 1:
                return ""
            pix = doc[0].get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            return pytesseract.image_to_string(img, lang="vie+eng") or ""
        finally:
            try:
                doc.close()
            except Exception:
                pass
    except Exception:
        return ""


def score_text(norm_text: str) -> tuple[float, list[str]]:
    """Cham diem text da normalize. Tra ve (score, signals)."""
    score = 0.0
    signals: list[str] = []
    for kw in STRONG_KEYWORDS:
        if kw in norm_text:
            score += 0.25
            signals.append(f"text:{kw}")
    for kw in WEAK_KEYWORDS:
        if kw in norm_text and f"text:{kw}" not in signals:
            # tranh dem trung khi weak la substring cua strong da hit
            if not any(kw in s for s in STRONG_KEYWORDS if f"text:{s}" in signals):
                score += 0.10
                signals.append(f"text:{kw}")
    return min(score, 1.0), signals


def score_filename(filename: str) -> tuple[float, list[str]]:
    norm = normalize_vi(os.path.basename(filename or ""))
    for pat in FILENAME_PATTERNS:
        if re.search(pat, norm):
            return 0.25, [f"name:{pat}"]
    return 0.0, []


def detect_invoice(pdf_path: str = "", filename: str = "",
                   use_ocr: bool = False) -> dict:
    """Nhan dien hoa don. Khong bao gio raise.

    Tra ve dict: is_invoice, confidence, signals, pages, method
    (text|filename|scan|none), needs_ocr, text_chars.
    """
    out = {"is_invoice": False, "confidence": 0.0, "signals": [],
           "pages": 0, "method": "none", "needs_ocr": False, "text_chars": 0}
    text, pages = ("", 0)
    if pdf_path and os.path.exists(pdf_path):
        text, pages = extract_pdf_text(pdf_path)
    out["pages"] = pages
    out["text_chars"] = len(text)

    if text.strip():
        out["method"] = "text"
        score, signals = score_text(normalize_vi(text))
        fscore, fsignals = score_filename(filename)
        # Ten file chi cong them khi da co tin hieu text (tranh false positive)
        if score >= 0.2:
            score = min(1.0, score + fscore)
            signals += fsignals
        out.update(confidence=round(score, 2), signals=signals)
        out["is_invoice"] = score >= INVOICE_THRESHOLD
        return out

    # Khong co text: thu OCR neu duoc bat, neu khong -> chua xac dinh
    if use_ocr and pdf_path:
        ocr_text = try_ocr_first_page(pdf_path)
        if ocr_text.strip():
            out["method"] = "text"
            score, signals = score_text(normalize_vi(ocr_text))
            out.update(confidence=round(score, 2), signals=signals,
                       text_chars=len(ocr_text))
            out["is_invoice"] = score >= INVOICE_THRESHOLD
            return out
    out["method"] = "scan" if (pdf_path and pages > 0) else "none"
    out["needs_ocr"] = out["method"] == "scan"
    # Ten file khong du de ket luan mot minh
    return out


def suggest_duplex(is_invoice: bool, actual_pages: int,
                   user_overridden: bool = False) -> str | None:
    """Tra ve 'long' neu nen tu in 2 mat canh dai, None neu giu nguyen.

    Quy tac: hoa don + thuc te >= 2 trang + nguoi dung khong ghi de.
    """
    if user_overridden:
        return None
    if is_invoice and (actual_pages or 0) >= 2:
        return "long"
    return None


def estimate_sheets_saved(jobs: list[tuple[int, bool]]) -> tuple[int, int, int]:
    """Tinh giay tiet kiem: jobs = [(so_trang, co_duplex), ...].

    Tra ve (tong_trang, tong_to_can, tiet_kiem).
    """
    total = sum(p for p, _ in jobs)
    sheets = sum(((p + 1) // 2 if dx else p) for p, dx in jobs)
    return total, sheets, total - sheets
