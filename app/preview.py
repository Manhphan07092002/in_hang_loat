"""
Modern PDF & Document Preview system.
Provides:
1. `PDFPreview`: Embeddable preview card for the main dashboard.
2. `PreviewWindow`: Dedicated high-resolution, full-featured interactive viewing modal.
"""

import os
import threading
import tkinter as tk
from tkinter import ttk
from typing import Optional, Union

import customtkinter as ctk
import fitz
from PIL import Image, ImageTk, ImageOps

from app.pdf_manager import PDFManager, FileInfo
from app.settings import THEME_COLORS
from app.utils import format_file_size


def get_paper_name(w_pt: float, h_pt: float) -> str:
    """Identify standard paper sizes from dimensions in points."""
    w, h = min(w_pt, h_pt), max(w_pt, h_pt)
    # Tolerances within 5 points
    if abs(w - 595.28) < 6 and abs(h - 841.89) < 6:
        return "A4 (210 × 297 mm)"
    elif abs(w - 841.89) < 6 and abs(h - 1190.55) < 6:
        return "A3 (297 × 420 mm)"
    elif abs(w - 419.53) < 6 and abs(h - 595.28) < 6:
        return "A5 (148 × 210 mm)"
    elif abs(w - 612.0) < 6 and abs(h - 792.0) < 6:
        return "Letter (8.5 × 11 in)"
    elif abs(w - 612.0) < 6 and abs(h - 1008.0) < 6:
        return "Legal (8.5 × 14 in)"
    else:
        # Convert pt to mm (1 pt = 25.4 / 72 mm)
        mm_w = round(w_pt * 25.4 / 72.0)
        mm_h = round(h_pt * 25.4 / 72.0)
        return f"Tùy chỉnh ({mm_w} × {mm_h} mm)"


# ═══════════════════════════════════════════════════════════════════════════
#  DEDICATED FULL-SCREEN / LARGE PREVIEW WINDOW MODAL
# ═══════════════════════════════════════════════════════════════════════════

class PreviewWindow(ctk.CTkToplevel):
    """
    Dedicated high-resolution document preview modal.
    Provides crystal clear rendering, zoom (25% - 400%), fit page, fit width,
    90° rotation, mouse drag-to-pan, direct page jumping, and keyboard shortcuts.
    """

    def __init__(self, master, pdf_manager: PDFManager, target: Union[str, FileInfo], initial_page: int = 0, **kwargs):
        super().__init__(master, **kwargs)
        self.pdf_manager = pdf_manager
        self.target = target
        self.current_path: Optional[str] = None
        self.current_finfo: Optional[FileInfo] = None
        self.current_page: int = initial_page
        self.total_pages: int = 0
        self.zoom_level: float = 1.0
        self.rotation: int = 0
        self.fit_mode: str = "page"  # 'page', 'width', 'custom'

        self._photo_image: Optional[ImageTk.PhotoImage] = None
        self._render_token: int = 0
        self._resize_job: Optional[str] = None
        self._drag_start_x: int = 0
        self._drag_start_y: int = 0

        self._setup_window()
        self._build_ui()
        self._bind_events()
        self.load_target(target, initial_page)

    def _setup_window(self):
        self.title("🔍 Xem Chi Tiết Bản In — Đang tải...")
        try:
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            w = min(1200, max(880, int(sw * 0.82)))
            h = min(920, max(660, int(sh * 0.86)))
            x = max(0, (sw - w) // 2)
            y = max(0, (sh - h) // 2)
            self.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            self.geometry("1100x820")

        self.minsize(750, 520)
        self.focus_force()
        self.lift()
        self.attributes("-topmost", True)
        self.after(200, lambda: self.attributes("-topmost", False))

    def _build_ui(self):
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # ── 1. Header Frame ──────────────────────────────────────────
        header = ctk.CTkFrame(
            self, fg_color=THEME_COLORS["card"],
            corner_radius=0, height=52,
        )
        header.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        header.grid_columnconfigure(0, weight=1)

        header_content = ctk.CTkFrame(header, fg_color="transparent")
        header_content.pack(fill="x", padx=16, pady=8)

        left_info = ctk.CTkFrame(header_content, fg_color="transparent")
        left_info.pack(side="left", fill="y")

        self.lbl_title = ctk.CTkLabel(
            left_info, text="📄 Đang mở tài liệu...",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color=THEME_COLORS["text"],
            anchor="w",
        )
        self.lbl_title.pack(side="top", anchor="w")

        self.lbl_doc_info = ctk.CTkLabel(
            left_info, text="",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=THEME_COLORS["text_muted"],
            anchor="w",
        )
        self.lbl_doc_info.pack(side="top", anchor="w")

        # Right Action Buttons
        actions = ctk.CTkFrame(header_content, fg_color="transparent")
        actions.pack(side="right")

        self.btn_rotate = ctk.CTkButton(
            actions, text="🔄 Xoay 90°", width=95, height=30,
            command=self._rotate_90,
            fg_color=THEME_COLORS["btn_secondary"],
            hover_color=THEME_COLORS["btn_secondary_hover"],
            text_color=THEME_COLORS["btn_secondary_text"],
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            corner_radius=6,
        )
        self.btn_rotate.pack(side="left", padx=4)

        self.btn_close = ctk.CTkButton(
            actions, text="✖ Đóng (Esc)", width=95, height=30,
            command=self.destroy,
            fg_color="#DC2626",
            hover_color="#B91C1C",
            text_color="#FFFFFF",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            corner_radius=6,
        )
        self.btn_close.pack(side="left", padx=(4, 0))

        # ── 2. Navigation & Zoom Controls Toolbar ───────────────────
        ctrl_bar = ctk.CTkFrame(
            self, fg_color=THEME_COLORS["card_alt"],
            corner_radius=0, height=44,
        )
        ctrl_bar.grid(row=1, column=0, sticky="ew", padx=0, pady=0)

        bar_inner = ctk.CTkFrame(ctrl_bar, fg_color="transparent")
        bar_inner.pack(fill="x", padx=16, pady=6)

        # Left: Page Navigation Group
        nav_grp = ctk.CTkFrame(bar_inner, fg_color="transparent")
        nav_grp.pack(side="left")

        self.btn_first = ctk.CTkButton(
            nav_grp, text="⏮", width=34, height=28,
            command=self._first_page,
            fg_color=THEME_COLORS["btn_secondary"],
            hover_color=THEME_COLORS["btn_secondary_hover"],
            text_color=THEME_COLORS["btn_secondary_text"],
            font=ctk.CTkFont(size=12, weight="bold"),
            corner_radius=6,
        )
        self.btn_first.pack(side="left", padx=2)

        self.btn_prev = ctk.CTkButton(
            nav_grp, text="◀ Trước", width=75, height=28,
            command=self._prev_page,
            fg_color=THEME_COLORS["primary"],
            hover_color=THEME_COLORS["primary_hover"],
            text_color="#FFFFFF",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            corner_radius=6,
        )
        self.btn_prev.pack(side="left", padx=2)

        # Page jump input
        ctk.CTkLabel(
            nav_grp, text="Trang",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=THEME_COLORS["text"],
        ).pack(side="left", padx=(6, 4))

        self.entry_page = ctk.CTkEntry(
            nav_grp, width=46, height=28,
            justify="center",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            corner_radius=6,
        )
        self.entry_page.pack(side="left", padx=2)
        self.entry_page.bind("<Return>", self._on_page_jump)

        self.lbl_total_pages = ctk.CTkLabel(
            nav_grp, text="/ 0",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=THEME_COLORS["text_muted"],
        )
        self.lbl_total_pages.pack(side="left", padx=(2, 6))

        self.btn_next = ctk.CTkButton(
            nav_grp, text="Sau ▶", width=75, height=28,
            command=self._next_page,
            fg_color=THEME_COLORS["primary"],
            hover_color=THEME_COLORS["primary_hover"],
            text_color="#FFFFFF",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            corner_radius=6,
        )
        self.btn_next.pack(side="left", padx=2)

        self.btn_last = ctk.CTkButton(
            nav_grp, text="⏭", width=34, height=28,
            command=self._last_page,
            fg_color=THEME_COLORS["btn_secondary"],
            hover_color=THEME_COLORS["btn_secondary_hover"],
            text_color=THEME_COLORS["btn_secondary_text"],
            font=ctk.CTkFont(size=12, weight="bold"),
            corner_radius=6,
        )
        self.btn_last.pack(side="left", padx=2)

        # Right: Zoom & View Mode Controls
        zoom_grp = ctk.CTkFrame(bar_inner, fg_color="transparent")
        zoom_grp.pack(side="right")

        self.btn_zoom_out = ctk.CTkButton(
            zoom_grp, text="−", width=32, height=28,
            command=self._zoom_out,
            fg_color=THEME_COLORS["btn_secondary"],
            hover_color=THEME_COLORS["btn_secondary_hover"],
            text_color=THEME_COLORS["btn_secondary_text"],
            font=ctk.CTkFont(size=14, weight="bold"),
            corner_radius=6,
        )
        self.btn_zoom_out.pack(side="left", padx=2)

        self.lbl_zoom = ctk.CTkLabel(
            zoom_grp, text="100%", width=46,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=THEME_COLORS["text"],
        )
        self.lbl_zoom.pack(side="left", padx=2)

        self.btn_zoom_in = ctk.CTkButton(
            zoom_grp, text="+", width=32, height=28,
            command=self._zoom_in,
            fg_color=THEME_COLORS["btn_secondary"],
            hover_color=THEME_COLORS["btn_secondary_hover"],
            text_color=THEME_COLORS["btn_secondary_text"],
            font=ctk.CTkFont(size=14, weight="bold"),
            corner_radius=6,
        )
        self.btn_zoom_in.pack(side="left", padx=2)

        self.btn_fit_page = ctk.CTkButton(
            zoom_grp, text="↕ Vừa trang (F)", width=96, height=28,
            command=self._fit_page,
            fg_color=THEME_COLORS["btn_secondary"],
            hover_color=THEME_COLORS["btn_secondary_hover"],
            text_color=THEME_COLORS["btn_secondary_text"],
            font=ctk.CTkFont(family="Segoe UI", size=11),
            corner_radius=6,
        )
        self.btn_fit_page.pack(side="left", padx=4)

        self.btn_fit_width = ctk.CTkButton(
            zoom_grp, text="↔ Vừa ngang (W)", width=105, height=28,
            command=self._fit_width,
            fg_color=THEME_COLORS["btn_secondary"],
            hover_color=THEME_COLORS["btn_secondary_hover"],
            text_color=THEME_COLORS["btn_secondary_text"],
            font=ctk.CTkFont(family="Segoe UI", size=11),
            corner_radius=6,
        )
        self.btn_fit_width.pack(side="left", padx=2)

        # Quick zoom presets menu
        self.zoom_menu = ctk.CTkOptionMenu(
            zoom_grp,
            values=["50%", "75%", "100%", "125%", "150%", "200%", "300%"],
            width=78, height=28,
            command=self._on_zoom_preset,
            fg_color=THEME_COLORS["card"],
            text_color=THEME_COLORS["text"],
            dropdown_fg_color=THEME_COLORS["card"],
            font=ctk.CTkFont(family="Segoe UI", size=11),
            corner_radius=6,
        )
        self.zoom_menu.set("100%")
        self.zoom_menu.pack(side="left", padx=(4, 0))

        # ── 3. High-Resolution Canvas Area with Dual Scrollbars ───────
        canvas_container = tk.Frame(self, bg="#0F172A" if ctk.get_appearance_mode() == "Dark" else "#E2E8F0", bd=0, highlightthickness=0)
        canvas_container.grid(row=2, column=0, sticky="nsew", padx=0, pady=0)
        canvas_container.grid_rowconfigure(0, weight=1)
        canvas_container.grid_columnconfigure(0, weight=1)

        self._canvas_bg = "#0F172A" if ctk.get_appearance_mode() == "Dark" else "#F1F5F9"
        self.canvas = tk.Canvas(
            canvas_container, bg=self._canvas_bg,
            highlightthickness=0, bd=0,
        )
        self.canvas.grid(row=0, column=0, sticky="nsew")

        self.vsb = ttk.Scrollbar(canvas_container, orient="vertical", command=self.canvas.yview)
        self.vsb.grid(row=0, column=1, sticky="ns")

        self.hsb = ttk.Scrollbar(canvas_container, orient="horizontal", command=self.canvas.xview)
        self.hsb.grid(row=1, column=0, sticky="ew")

        self.canvas.configure(xscrollcommand=self.hsb.set, yscrollcommand=self.vsb.set)

    def _bind_events(self):
        # Resize debounce
        self.canvas.bind("<Configure>", self._on_canvas_resize)

        # Mouse Drag & Pan
        self.canvas.bind("<ButtonPress-1>", self._on_drag_start)
        self.canvas.bind("<B1-Motion>", self._on_drag_motion)
        self.canvas.bind("<ButtonRelease-1>", self._on_drag_end)

        # Mousewheel
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Shift-MouseWheel>", self._on_shift_mousewheel)
        self.canvas.bind("<Control-MouseWheel>", self._on_ctrl_mousewheel)

        # Keyboard shortcuts
        self.bind("<Escape>", lambda e: self.destroy())
        self.bind("<Left>", lambda e: self._prev_page())
        self.bind("<Prior>", lambda e: self._prev_page())
        self.bind("<Up>", lambda e: self._prev_page())
        self.bind("<Right>", lambda e: self._next_page())
        self.bind("<Next>", lambda e: self._next_page())
        self.bind("<Down>", lambda e: self._next_page())
        self.bind("<space>", lambda e: self._next_page())
        self.bind("<Home>", lambda e: self._first_page())
        self.bind("<End>", lambda e: self._last_page())
        self.bind("<plus>", lambda e: self._zoom_in())
        self.bind("<equal>", lambda e: self._zoom_in())
        self.bind("<KP_Add>", lambda e: self._zoom_in())
        self.bind("<minus>", lambda e: self._zoom_out())
        self.bind("<underscore>", lambda e: self._zoom_out())
        self.bind("<KP_Subtract>", lambda e: self._zoom_out())
        self.bind("<w>", lambda e: self._fit_width())
        self.bind("<W>", lambda e: self._fit_width())
        self.bind("<f>", lambda e: self._fit_page())
        self.bind("<F>", lambda e: self._fit_page())
        self.bind("<r>", lambda e: self._rotate_90())
        self.bind("<R>", lambda e: self._rotate_90())

    # ── Public API ───────────────────────────────────────────────────────

    def load_target(self, target: Union[str, FileInfo], initial_page: int = 0):
        self.target = target
        self.current_page = initial_page

        if isinstance(target, FileInfo):
            self.current_finfo = target
            fname = target.filename
            self.lbl_title.configure(text=f"📄 {fname}")
            if target.is_converted_ready and target.pdf_path and os.path.exists(target.pdf_path):
                self._load_pdf_path(target.pdf_path)
            else:
                self._show_loading(f"Đang chuẩn bị bản xem trước độ nét cao cho: {fname}...")
                def _bg():
                    try:
                        pdf_path = self.pdf_manager.ensure_pdf(target)
                        self.after(0, lambda: self._load_pdf_path(pdf_path))
                    except Exception as e:
                        self.after(0, lambda: self._show_error(f"Không thể mở xem trước:\n{e}"))
                threading.Thread(target=_bg, daemon=True).start()
        elif isinstance(target, str):
            self.lbl_title.configure(text=f"📄 {os.path.basename(target)}")
            self._load_pdf_path(target)

    def _load_pdf_path(self, path: str):
        if not path or not os.path.exists(path):
            self._show_error("Tệp tin không tồn tại hoặc không thể đọc.")
            return

        self.current_path = path
        self.total_pages = self.pdf_manager.get_page_count(path)
        self.current_page = max(0, min(self.total_pages - 1, self.current_page))
        self.title(f"🔍 Xem Chi Tiết Bản In — {os.path.basename(path)}")

        # Update metadata info badge
        try:
            w_pt, h_pt = self.pdf_manager.get_page_size(path, self.current_page)
            paper_str = get_paper_name(w_pt, h_pt)
            orient_str = "Ngang (Landscape)" if w_pt > h_pt else "Dọc (Portrait)"
            size_str = format_file_size(os.path.getsize(path))
            self.lbl_doc_info.configure(
                text=f"Khổ giấy: {paper_str}  •  Hướng in: {orient_str}  •  Dung lượng: {size_str}"
            )
        except Exception:
            self.lbl_doc_info.configure(text="")

        self._sync_nav()
        self._render()

    # ── Rendering Pipeline ───────────────────────────────────────────────

    def _render(self):
        if not self.current_path:
            return

        cw = max(self.canvas.winfo_width(), 400)
        ch = max(self.canvas.winfo_height(), 400)

        if not hasattr(self, "_render_token"):
            self._render_token = 0
        self._render_token += 1
        cur_token = self._render_token

        path = self.current_path
        page_num = self.current_page
        fit_mode = self.fit_mode
        zoom = self.zoom_level
        rotation = self.rotation
        is_dark = ctk.get_appearance_mode() == "Dark"

        def _bg_render():
            doc = None
            try:
                doc = fitz.open(path)
                if page_num < 0 or page_num >= doc.page_count:
                    return

                page = doc[page_num]
                rect = page.rect
                # Adjust rect dimensions if rotated 90 or 270
                if rotation in (90, 270):
                    pw, ph = rect.height, rect.width
                else:
                    pw, ph = rect.width, rect.height

                # Determine scale factor
                if fit_mode == "page":
                    scale = min((cw - 48) / pw, (ch - 48) / ph) * zoom
                elif fit_mode == "width":
                    scale = ((cw - 48) / pw) * zoom
                else:
                    scale = zoom * (120.0 / 72.0)  # Standard sharp baseline

                scale = max(0.15, min(6.0, scale))

                mat = fitz.Matrix(scale, scale).prerotate(rotation)
                pix = page.get_pixmap(matrix=mat, alpha=False)
                img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

                border_col = "#334155" if is_dark else "#CBD5E1"
                img_bordered = ImageOps.expand(img, border=2, fill=border_col)

                def _apply():
                    if getattr(self, "_render_token", 0) != cur_token:
                        return
                    try:
                        if not self.winfo_exists():
                            return
                    except Exception:
                        return
                    if self.current_path != path or self.current_page != page_num:
                        return

                    self.canvas.delete("all")
                    self._photo_image = ImageTk.PhotoImage(img_bordered)

                    cur_w = max(self.canvas.winfo_width(), cw)
                    cur_h = max(self.canvas.winfo_height(), ch)

                    scroll_w = max(cur_w, img_bordered.width + 40)
                    scroll_h = max(cur_h, img_bordered.height + 40)

                    cx = scroll_w // 2
                    cy = scroll_h // 2

                    self.canvas.create_image(cx, cy, image=self._photo_image, anchor="center")
                    self.canvas.configure(scrollregion=(0, 0, scroll_w, scroll_h))

                try:
                    self.after(0, _apply)
                except Exception:
                    pass
            except Exception as exc:
                err = str(exc)
                def _show_err():
                    try:
                        if not self.winfo_exists():
                            return
                    except Exception:
                        return
                    if getattr(self, "_render_token", 0) != cur_token:
                        return
                    self._show_error(f"Không thể render trang:\n{err}")
                try:
                    self.after(0, _show_err)
                except Exception:
                    pass
            finally:
                if doc is not None:
                    try:
                        doc.close()
                    except Exception:
                        pass

        threading.Thread(target=_bg_render, daemon=True).start()

    def _sync_nav(self):
        if self.total_pages > 0:
            self.lbl_total_pages.configure(text=f"/ {self.total_pages}")
            self.entry_page.delete(0, "end")
            self.entry_page.insert(0, str(self.current_page + 1))

            self.btn_first.configure(state="normal" if self.current_page > 0 else "disabled")
            self.btn_prev.configure(state="normal" if self.current_page > 0 else "disabled")
            self.btn_next.configure(state="normal" if self.current_page < self.total_pages - 1 else "disabled")
            self.btn_last.configure(state="normal" if self.current_page < self.total_pages - 1 else "disabled")
        else:
            self.lbl_total_pages.configure(text="/ 0")
            self.entry_page.delete(0, "end")
            self.btn_first.configure(state="disabled")
            self.btn_prev.configure(state="disabled")
            self.btn_next.configure(state="disabled")
            self.btn_last.configure(state="disabled")

        pct = int(self.zoom_level * 100)
        self.lbl_zoom.configure(text=f"{pct}%")

    def _show_loading(self, text: str):
        self.canvas.delete("all")
        cw = max(self.canvas.winfo_width(), 400)
        ch = max(self.canvas.winfo_height(), 400)
        self.canvas.create_text(
            cw // 2, ch // 2 - 20,
            text="⏳", font=("Segoe UI Emoji", 40),
            fill="#2563EB",
        )
        self.canvas.create_text(
            cw // 2, ch // 2 + 30,
            text=text,
            font=("Segoe UI", 12, "bold"),
            fill="#2563EB",
        )

    def _show_error(self, text: str):
        self.canvas.delete("all")
        cw = max(self.canvas.winfo_width(), 400)
        ch = max(self.canvas.winfo_height(), 400)
        self.canvas.create_text(
            cw // 2, ch // 2 - 20,
            text="⚠️", font=("Segoe UI Emoji", 40),
            fill="#DC2626",
        )
        self.canvas.create_text(
            cw // 2, ch // 2 + 30,
            text=text,
            font=("Segoe UI", 11, "bold"),
            fill="#DC2626",
            justify="center",
        )

    # ── Navigation Actions ───────────────────────────────────────────────

    def _first_page(self):
        if self.current_page != 0:
            self.current_page = 0
            self._sync_nav()
            self._render()

    def _last_page(self):
        if self.total_pages > 0 and self.current_page != self.total_pages - 1:
            self.current_page = self.total_pages - 1
            self._sync_nav()
            self._render()

    def _prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self._sync_nav()
            self._render()

    def _next_page(self):
        if self.current_page < self.total_pages - 1:
            self.current_page += 1
            self._sync_nav()
            self._render()

    def _on_page_jump(self, _event=None):
        try:
            val = int(self.entry_page.get().strip())
            if 1 <= val <= self.total_pages:
                self.current_page = val - 1
                self._sync_nav()
                self._render()
            else:
                self._sync_nav()
        except ValueError:
            self._sync_nav()

    # ── Zoom Actions ─────────────────────────────────────────────────────

    def _zoom_in(self):
        self.fit_mode = "custom"
        self.zoom_level = min(4.0, self.zoom_level + 0.15)
        self._sync_nav()
        self._render()

    def _zoom_out(self):
        self.fit_mode = "custom"
        self.zoom_level = max(0.25, self.zoom_level - 0.15)
        self._sync_nav()
        self._render()

    def _fit_page(self):
        self.fit_mode = "page"
        self.zoom_level = 1.0
        self._sync_nav()
        self._render()

    def _fit_width(self):
        self.fit_mode = "width"
        self.zoom_level = 1.0
        self._sync_nav()
        self._render()

    def _rotate_90(self):
        self.rotation = (self.rotation + 90) % 360
        self._render()

    def _on_zoom_preset(self, val_str: str):
        try:
            pct = int(val_str.replace("%", "").strip())
            self.fit_mode = "custom"
            self.zoom_level = pct / 100.0
            self._sync_nav()
            self._render()
        except Exception:
            pass

    # ── Mouse & Pan Handling ─────────────────────────────────────────────

    def _on_canvas_resize(self, _event=None):
        if self._resize_job is not None:
            try:
                self.after_cancel(self._resize_job)
            except Exception:
                pass
        self._resize_job = self.after(50, self._do_canvas_resize)

    def _do_canvas_resize(self):
        self._resize_job = None
        if self.current_path:
            self._render()

    def _on_drag_start(self, event):
        self.canvas.scan_mark(event.x, event.y)
        self.canvas.configure(cursor="fleur")
        self._drag_start_x = event.x
        self._drag_start_y = event.y

    def _on_drag_motion(self, event):
        self.canvas.scan_dragto(event.x, event.y, gain=1)

    def _on_drag_end(self, _event):
        self.canvas.configure(cursor="arrow")

    def _on_mousewheel(self, event):
        # Windows: event.delta ±120; Linux: Button-4/5; macOS: delta nhỏ
        try:
            delta = getattr(event, "delta", 0)
            if delta:
                self.canvas.yview_scroll(int(-1 * (delta / 120)), "units")
            elif getattr(event, "num", None) == 4:
                self.canvas.yview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                self.canvas.yview_scroll(1, "units")
        except Exception:
            pass

    def _on_shift_mousewheel(self, event):
        try:
            delta = getattr(event, "delta", 0)
            if delta:
                self.canvas.xview_scroll(int(-1 * (delta / 120)), "units")
            elif getattr(event, "num", None) == 4:
                self.canvas.xview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                self.canvas.xview_scroll(1, "units")
        except Exception:
            pass

    def _on_ctrl_mousewheel(self, event):
        if event.delta > 0:
            self._zoom_in()
        else:
            self._zoom_out()


# ═══════════════════════════════════════════════════════════════════════════
#  EMBEDDABLE DASHBOARD PREVIEW WIDGET
# ═══════════════════════════════════════════════════════════════════════════

class PDFPreview(ctk.CTkFrame):
    """
    Modern embeddable widget that previews document pages with navigation,
    zoom controls, and clean document card rendering.
    Clicking anywhere on the preview canvas immediately opens the full-screen modal!
    """

    def __init__(self, master, pdf_manager: PDFManager, **kwargs):
        kwargs.setdefault("fg_color", THEME_COLORS["card"])
        kwargs.setdefault("corner_radius", 12)
        kwargs.setdefault("border_width", 1)
        kwargs.setdefault("border_color", THEME_COLORS["card_border"])
        super().__init__(master, **kwargs)

        self.pdf_manager = pdf_manager
        self.current_path: Optional[str] = None
        self.current_finfo: Optional[FileInfo] = None
        self.current_page: int = 0
        self.total_pages: int = 0
        self.zoom_level: float = 1.0
        self._photo_image: Optional[ImageTk.PhotoImage] = None
        self._resize_job: Optional[str] = None
        self._fullscreen_win: Optional[PreviewWindow] = None

        self._build_ui()

    # ── UI construction ──────────────────────────────────────────────────

    def _build_ui(self):
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # ── Header row ───────────────────────────────────────────────
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 6))

        title_box = ctk.CTkFrame(header, fg_color="transparent")
        title_box.pack(side="left")

        ctk.CTkLabel(
            title_box, text="👁️ Xem Trước",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=THEME_COLORS["text"],
        ).pack(side="left")

        self.preview_info_lbl = ctk.CTkLabel(
            title_box, text="",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=THEME_COLORS["text_muted"],
        )
        self.preview_info_lbl.pack(side="left", padx=(8, 0))

        # Prominent Fullscreen Preview Button
        self.btn_fullscreen = ctk.CTkButton(
            header, text="🔍 Mở Cửa Sổ Xem (Toàn Màn Hình)", height=28,
            command=self.open_fullscreen_preview,
            fg_color=THEME_COLORS["primary"],
            hover_color=THEME_COLORS["primary_hover"],
            text_color="#FFFFFF",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            corner_radius=6,
        )
        self.btn_fullscreen.pack(side="right")

        # ── Canvas Container (Interactive Click-to-Open) ─────────────
        canvas_card = ctk.CTkFrame(
            self, fg_color=THEME_COLORS["card_alt"],
            corner_radius=8, border_width=1,
            border_color=THEME_COLORS["card_border"],
        )
        canvas_card.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 6))
        canvas_card.grid_rowconfigure(0, weight=1)
        canvas_card.grid_columnconfigure(0, weight=1)

        self._canvas_bg = "#F1F5F9" if ctk.get_appearance_mode() == "Light" else "#0F172A"

        self.canvas = tk.Canvas(
            canvas_card, bg=self._canvas_bg,
            highlightthickness=0, bd=0,
            cursor="hand2",
        )
        self.canvas.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)

        # Bind click anywhere on canvas to open the large window
        self.canvas.bind("<Button-1>", lambda e: self.open_fullscreen_preview())
        self.canvas.bind("<Double-1>", lambda e: self.open_fullscreen_preview())
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Configure>", self._on_canvas_resize)

        self._show_placeholder()

        # ── Floating Navigation Toolbar ──────────────────────────────
        nav = ctk.CTkFrame(
            self, fg_color=THEME_COLORS["card_alt"],
            corner_radius=8, height=36,
        )
        nav.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 10))
        nav.grid_columnconfigure(1, weight=1)

        self.btn_prev = ctk.CTkButton(
            nav, text="◀ Trang trước", width=105, height=28,
            command=self._prev_page, state="disabled",
            fg_color=THEME_COLORS["btn_secondary"],
            hover_color=THEME_COLORS["btn_secondary_hover"],
            text_color=THEME_COLORS["btn_secondary_text"],
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            corner_radius=6,
        )
        self.btn_prev.grid(row=0, column=0, padx=6, pady=4)

        self.page_label = ctk.CTkLabel(
            nav, text="Chưa chọn file",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=THEME_COLORS["text"],
        )
        self.page_label.grid(row=0, column=1, pady=4)

        self.btn_next = ctk.CTkButton(
            nav, text="Trang sau ▶", width=105, height=28,
            command=self._next_page, state="disabled",
            fg_color=THEME_COLORS["btn_secondary"],
            hover_color=THEME_COLORS["btn_secondary_hover"],
            text_color=THEME_COLORS["btn_secondary_text"],
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            corner_radius=6,
        )
        self.btn_next.grid(row=0, column=2, padx=6, pady=4)

    # ── Public API ───────────────────────────────────────────────────────

    def load_file(self, target: Union[str, FileInfo]):
        """
        Load a FileInfo object or file path for preview with async conversion if needed.
        """
        if isinstance(target, FileInfo):
            self.current_finfo = target
            if target.is_converted_ready and target.pdf_path and os.path.exists(target.pdf_path):
                self.load_pdf(target.pdf_path)
            else:
                self.current_path = None
                self._show_loading(f"Đang chuẩn bị bản xem trước: {target.filename}...")
                self._sync_nav()

                def _bg_convert():
                    try:
                        pdf_path = self.pdf_manager.ensure_pdf(target)
                        self.after(0, lambda: self.load_pdf(pdf_path))
                    except Exception as e:
                        self.after(0, lambda: self._show_error(f"Không thể mở xem trước:\n{e}"))
                threading.Thread(target=_bg_convert, daemon=True).start()
        elif isinstance(target, str):
            self.current_finfo = None
            self.load_pdf(target)

    def load_pdf(self, path: str):
        """Load a document for preview asynchronously."""
        if not path or not os.path.exists(path):
            self.clear()
            return
        self.current_path = path
        self.current_page = 0
        self.total_pages = self.pdf_manager.get_page_count(path)
        self.zoom_level = 1.0
        self._sync_nav()
        self._render()

    def open_fullscreen_preview(self, target: Optional[Union[str, FileInfo]] = None):
        """Open a large high-resolution modal preview window."""
        item = target or self.current_finfo or self.current_path
        if not item:
            return

        # If window is already open, focus it and load new item
        if self._fullscreen_win is not None:
            try:
                if self._fullscreen_win.winfo_exists():
                    self._fullscreen_win.load_target(item, self.current_page)
                    self._fullscreen_win.lift()
                    self._fullscreen_win.focus_force()
                    return
            except Exception:
                pass

        top_master = self.winfo_toplevel()
        self._fullscreen_win = PreviewWindow(
            top_master,
            pdf_manager=self.pdf_manager,
            target=item,
            initial_page=self.current_page,
        )

    def clear(self):
        """Reset the preview widget to its empty state."""
        self.current_path = None
        self.current_finfo = None
        self.current_page = 0
        self.total_pages = 0
        self._photo_image = None
        self.preview_info_lbl.configure(text="")
        self._show_placeholder()
        self._sync_nav()

    def update_theme(self):
        """Adapt canvas background to current Light/Dark theme."""
        is_light = ctk.get_appearance_mode() == "Light"
        self._canvas_bg = "#F1F5F9" if is_light else "#0F172A"
        self.canvas.configure(bg=self._canvas_bg)
        if self.current_path:
            self._render()
        else:
            self._show_placeholder()

    # ── Internal helpers ─────────────────────────────────────────────────

    def _show_placeholder(self):
        self.canvas.delete("all")
        cw = max(self.canvas.winfo_width(), 300)
        ch = max(self.canvas.winfo_height(), 350)
        text_color = "#94A3B8" if ctk.get_appearance_mode() == "Light" else "#475569"

        self.canvas.create_text(
            cw // 2, ch // 2 - 30,
            text="📄", font=("Segoe UI Emoji", 36),
            fill=text_color,
        )
        self.canvas.create_text(
            cw // 2, ch // 2 + 15,
            text="Chọn file trong hàng đợi để xem trước",
            font=("Segoe UI", 12, "bold"),
            fill=text_color,
        )
        self.canvas.create_text(
            cw // 2, ch // 2 + 38,
            text="🔍 Bấm vào đây để mở cửa sổ xem lớn",
            font=("Segoe UI", 10),
            fill="#2563EB",
        )

    def _show_loading(self, text: str = "Đang tải bản xem trước..."):
        self.canvas.delete("all")
        cw = max(self.canvas.winfo_width(), 300)
        ch = max(self.canvas.winfo_height(), 350)
        text_color = "#2563EB"

        self.canvas.create_text(
            cw // 2, ch // 2 - 15,
            text="⏳", font=("Segoe UI Emoji", 32),
            fill=text_color,
        )
        self.canvas.create_text(
            cw // 2, ch // 2 + 25,
            text=text,
            font=("Segoe UI", 11, "bold"),
            fill=text_color,
        )

    def _show_error(self, err_text: str):
        self.canvas.delete("all")
        cw = max(self.canvas.winfo_width(), 300)
        ch = max(self.canvas.winfo_height(), 350)

        self.canvas.create_text(
            cw // 2, ch // 2 - 15,
            text="⚠️", font=("Segoe UI Emoji", 32),
            fill="#DC2626",
        )
        self.canvas.create_text(
            cw // 2, ch // 2 + 25,
            text=err_text,
            font=("Segoe UI", 10),
            fill="#DC2626",
            justify="center",
        )

    def _render(self):
        """Render the current page asynchronously with debouncing."""
        if self.current_path is None:
            return

        cw = max(self.canvas.winfo_width(), 200)
        ch = max(self.canvas.winfo_height(), 200)

        tw = int((cw - 24) * self.zoom_level)
        th = int((ch - 24) * self.zoom_level)

        if tw <= 10 or th <= 10:
            return

        if not hasattr(self, "_render_token"):
            self._render_token = 0
        self._render_token += 1
        cur_token = self._render_token

        path = self.current_path
        page_num = self.current_page
        is_dark = ctk.get_appearance_mode() == "Dark"

        def _bg_render():
            try:
                img = self.pdf_manager.render_page(path, page_num, width=tw, height=th)
                if img is None:
                    def _show_none():
                        try:
                            if not self.winfo_exists():
                                return
                        except Exception:
                            return
                        if getattr(self, "_render_token", 0) != cur_token:
                            return
                        if self.current_path != path or self.current_page != page_num:
                            return
                        self._show_error("Không thể render trang này.")
                    try:
                        self.after(0, _show_none)
                    except Exception:
                        pass
                    return

                border_col = "#334155" if is_dark else "#CBD5E1"
                img_bordered = ImageOps.expand(img, border=1, fill=border_col)

                def _apply():
                    if getattr(self, "_render_token", 0) != cur_token:
                        return
                    try:
                        if not self.winfo_exists():
                            return
                    except Exception:
                        return
                    if self.current_path != path or self.current_page != page_num:
                        return
                    self.canvas.delete("all")
                    self._photo_image = ImageTk.PhotoImage(img_bordered)

                    cur_w = max(self.canvas.winfo_width(), 200)
                    cur_h = max(self.canvas.winfo_height(), 200)
                    cx = max(cur_w, img_bordered.width + 20) // 2
                    cy = max(cur_h, img_bordered.height + 20) // 2

                    self.canvas.create_image(cx, cy, image=self._photo_image, anchor="center")
                    # Click hint overlay badge at bottom
                    self.canvas.create_text(
                        cx, cy + (img_bordered.height // 2) + 12,
                        text="🔍 Nhấp vào để mở cửa sổ xem lớn",
                        font=("Segoe UI", 9, "bold"),
                        fill="#2563EB",
                    )
                    self.canvas.configure(
                        scrollregion=(0, 0, max(cur_w, img_bordered.width + 20), max(cur_h, img_bordered.height + 30)),
                    )

                try:
                    self.after(0, _apply)
                except Exception:
                    pass
            except Exception as exc:
                err = str(exc)
                def _show_err():
                    try:
                        if not self.winfo_exists():
                            return
                    except Exception:
                        return
                    if getattr(self, "_render_token", 0) != cur_token:
                        return
                    self._show_error(f"Không thể render:\n{err}")
                try:
                    self.after(0, _show_err)
                except Exception:
                    pass

        threading.Thread(target=_bg_render, daemon=True).start()

    def _sync_nav(self):
        if self.total_pages > 0:
            self.page_label.configure(
                text=f"Trang {self.current_page + 1} / {self.total_pages}",
            )
            self.btn_prev.configure(
                state="normal" if self.current_page > 0 else "disabled",
            )
            self.btn_next.configure(
                state="normal" if self.current_page < self.total_pages - 1 else "disabled",
            )
            self.preview_info_lbl.configure(
                text=f"({self.total_pages} trang)",
            )
        else:
            self.page_label.configure(text="Chưa chọn file")
            self.btn_prev.configure(state="disabled")
            self.btn_next.configure(state="disabled")
            self.preview_info_lbl.configure(text="")

    def _on_canvas_resize(self, _event=None):
        if self._resize_job is not None:
            try:
                self.after_cancel(self._resize_job)
            except Exception:
                pass
        self._resize_job = self.after(50, self._do_canvas_resize)

    def _do_canvas_resize(self):
        self._resize_job = None
        if self.current_path:
            self._render()
        else:
            self._show_placeholder()

    # ── Navigation ───────────────────────────────────────────────────────

    def _prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self._render()
            self._sync_nav()

    def _next_page(self):
        if self.current_page < self.total_pages - 1:
            self.current_page += 1
            self._render()
            self._sync_nav()

    def _on_mousewheel(self, event):
        try:
            delta = getattr(event, "delta", 0)
            if delta:
                self.canvas.yview_scroll(int(-1 * (delta / 120)), "units")
            elif getattr(event, "num", None) == 4:
                self.canvas.yview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                self.canvas.yview_scroll(1, "units")
        except Exception:
            pass
