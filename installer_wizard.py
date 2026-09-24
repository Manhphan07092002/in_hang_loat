"""
Windows Setup / Installation Wizard for PDF Batch Printer Pro.
Provides a modern 5-step installer wizard with Desktop & Start Menu shortcuts,
Windows Registry uninstallation registration, and custom install directory picker.
"""
import os
import sys
import shutil
import zipfile
import winreg
import threading
import subprocess
from typing import Optional

# Ensure UTF-8 on Windows
if sys.stdout is not None:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import customtkinter as ctk
from PIL import Image, ImageTk

APP_NAME = "PDF Batch Printer Pro"
APP_VERSION = "1.0.0.1"
PUBLISHER = "PDF Batch Printer Pro Team"

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")


def get_bundle_dir() -> str:
    """Get absolute path to resource, works for dev and PyInstaller."""
    try:
        return sys._MEIPASS
    except Exception:
        return os.path.dirname(os.path.abspath(__file__))


def create_shortcut(target_path: str, shortcut_path: str, icon_path: str, description: str = ""):
    """Creates a Windows .lnk shortcut using WScript.Shell."""
    try:
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(shortcut_path)
        shortcut.TargetPath = target_path
        shortcut.WorkingDirectory = os.path.dirname(target_path)
        shortcut.IconLocation = icon_path
        shortcut.Description = description
        shortcut.save()
    except Exception as exc:
        print(f"Error creating shortcut {shortcut_path}: {exc}")


def register_uninstall(install_dir: str, exe_path: str, uninstaller_path: str, icon_path: str):
    """Registers the application in Windows Add/Remove Programs."""
    try:
        key_path = rf"Software\Microsoft\Windows\CurrentVersion\Uninstall\PDFBatchPrinterPro"
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, APP_NAME)
            winreg.SetValueEx(key, "DisplayVersion", 0, winreg.REG_SZ, APP_VERSION)
            winreg.SetValueEx(key, "Publisher", 0, winreg.REG_SZ, PUBLISHER)
            winreg.SetValueEx(key, "DisplayIcon", 0, winreg.REG_SZ, icon_path)
            winreg.SetValueEx(key, "InstallLocation", 0, winreg.REG_SZ, install_dir)
            winreg.SetValueEx(key, "UninstallString", 0, winreg.REG_SZ, f'"{uninstaller_path}"')
            winreg.SetValueEx(key, "QuietUninstallString", 0, winreg.REG_SZ, f'"{uninstaller_path}" /S')
            winreg.SetValueEx(key, "NoModify", 0, winreg.REG_DWORD, 1)
            winreg.SetValueEx(key, "NoRepair", 0, winreg.REG_DWORD, 1)
    except Exception as exc:
        print(f"Error registering uninstall: {exc}")


class SetupWizardApp(ctk.CTk):
    """Modern Multi-Step Windows Setup Wizard."""

    def __init__(self):
        super().__init__()

        self.title(f"Cài đặt — {APP_NAME}")
        self.geometry("680x480")
        self.minsize(680, 480)
        self.resizable(False, False)

        # Set App Icon
        ico_file = os.path.join(get_bundle_dir(), "assets", "app.ico")
        if os.path.exists(ico_file):
            try:
                self.iconbitmap(default=ico_file)
            except Exception:
                pass

        # Default Destination Folder in LocalAppData
        local_app_data = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
        self.install_dir_var = ctk.StringVar(
            value=os.path.join(local_app_data, "Programs", "PDFBatchPrinterPro")
        )

        # Options
        self.desktop_shortcut_var = ctk.BooleanVar(value=True)
        self.start_menu_shortcut_var = ctk.BooleanVar(value=True)
        self.open_guide_var = ctk.BooleanVar(value=True)
        self.launch_after_var = ctk.BooleanVar(value=True)

        self.step = 1
        self._build_layout()
        self._show_step(1)

    def _build_layout(self):
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Main Container
        self.main_frame = ctk.CTkFrame(self, fg_color="#F8FAFC", corner_radius=0)
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=0)
        self.main_frame.grid_columnconfigure(0, weight=1)

        # Content Frame
        self.content_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.content_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        self.content_frame.grid_rowconfigure(0, weight=1)
        self.content_frame.grid_columnconfigure(0, weight=1)

        # Bottom Action Bar
        self.bottom_bar = ctk.CTkFrame(self.main_frame, fg_color="#E2E8F0", height=60, corner_radius=0)
        self.bottom_bar.grid(row=1, column=0, sticky="ew")
        self.bottom_bar.grid_propagate(False)

        # Buttons
        self.btn_cancel = ctk.CTkButton(
            self.bottom_bar, text="Hủy bỏ", width=100, height=34,
            command=self._cancel, fg_color="#94A3B8", hover_color="#64748B",
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        self.btn_cancel.pack(side="right", padx=(10, 20), pady=13)

        self.btn_next = ctk.CTkButton(
            self.bottom_bar, text="Tiếp tục >", width=110, height=34,
            command=self._next_step, fg_color="#2563EB", hover_color="#1D4ED8",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
        )
        self.btn_next.pack(side="right", padx=10, pady=13)

        self.btn_back = ctk.CTkButton(
            self.bottom_bar, text="< Quay lại", width=100, height=34,
            command=self._prev_step, fg_color="#CBD5E1", hover_color="#94A3B8",
            text_color="#0F172A",
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        self.btn_back.pack(side="right", padx=10, pady=13)

    def _clear_content(self):
        for widget in self.content_frame.winfo_children():
            widget.destroy()

    def _show_step(self, step: int):
        self.step = step
        self._clear_content()

        if step == 1:
            self._render_step1_welcome()
        elif step == 2:
            self._render_step2_directory()
        elif step == 3:
            self._render_step3_options()
        elif step == 4:
            self._render_step4_installing()
        elif step == 5:
            self._render_step5_complete()

    # ── Step 1: Welcome ──────────────────────────────────────────────

    def _render_step1_welcome(self):
        self.btn_back.configure(state="disabled")
        self.btn_next.configure(text="Tiếp tục >", state="normal")
        self.btn_cancel.configure(state="normal")

        wrap = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        wrap.pack(fill="both", expand=True)

        # Left Banner / Logo
        logo_path = os.path.join(get_bundle_dir(), "assets", "logo.png")
        if os.path.exists(logo_path):
            try:
                logo_pil = Image.open(logo_path)
                self.logo_ctk = ctk.CTkImage(light_image=logo_pil, dark_image=logo_pil, size=(110, 110))
                ctk.CTkLabel(wrap, image=self.logo_ctk, text="").pack(pady=(10, 10))
            except Exception:
                pass

        ctk.CTkLabel(
            wrap, text="Chào mừng đến với Trình Cài Đặt\nPDF BATCH PRINTER PRO",
            font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
            text_color="#0F172A", justify="center",
        ).pack(pady=(0, 12))

        ctk.CTkLabel(
            wrap,
            text="Phần mềm in ấn hàng loạt đa định dạng (PDF, Word, Excel, PowerPoint, Ảnh)\n"
                 "chuyên nghiệp dành cho hệ điều hành Windows.\n\n"
                 "Trình cài đặt sẽ hướng dẫn bạn cài đặt các tệp tin cần thiết lên máy tính.\n"
                 "Nhấn 'Tiếp tục' để bắt đầu thiết lập.",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color="#475569", justify="center",
        ).pack(pady=(0, 10))

    # ── Step 2: Choose Directory ─────────────────────────────────────

    def _render_step2_directory(self):
        self.btn_back.configure(state="normal")
        self.btn_next.configure(text="Tiếp tục >", state="normal")

        wrap = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        wrap.pack(fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(
            wrap, text="Chọn Thư Mục Cài Đặt",
            font=ctk.CTkFont(family="Segoe UI", size=17, weight="bold"),
            text_color="#0F172A",
        ).pack(anchor="w", pady=(0, 6))

        ctk.CTkLabel(
            wrap, text=f"Trình cài đặt sẽ cài đặt {APP_NAME} vào thư mục sau.\nĐể chọn thư mục khác, nhấn nút 'Duyệt...'",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#475569",
        ).pack(anchor="w", pady=(0, 16))

        dir_box = ctk.CTkFrame(wrap, fg_color="#FFFFFF", corner_radius=8, border_width=1, border_color="#CBD5E1")
        dir_box.pack(fill="x", pady=(0, 16), ipady=4)
        dir_box.grid_columnconfigure(0, weight=1)

        self.dir_entry = ctk.CTkEntry(
            dir_box, textvariable=self.install_dir_var,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            border_width=0, fg_color="transparent",
        )
        self.dir_entry.grid(row=0, column=0, sticky="ew", padx=(10, 6), pady=6)

        ctk.CTkButton(
            dir_box, text="Duyệt...", width=80, height=32,
            command=self._browse_dir,
            fg_color="#0284C7", hover_color="#0369A1",
            font=ctk.CTkFont(family="Segoe UI", size=12),
        ).grid(row=0, column=1, padx=(0, 8), pady=6)

        info_frame = ctk.CTkFrame(wrap, fg_color="#F1F5F9", corner_radius=8)
        info_frame.pack(fill="x", pady=10, padx=2)

        ctk.CTkLabel(
            info_frame, text="• Dung lượng yêu cầu: ~60.0 MB\n• Dung lượng trống khả dụng: Đầy đủ",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#334155", justify="left",
        ).pack(anchor="w", padx=14, pady=10)

    def _browse_dir(self):
        from tkinter import filedialog
        folder = filedialog.askdirectory(
            title="Chọn thư mục cài đặt",
            initialdir=self.install_dir_var.get(),
        )
        if folder:
            self.install_dir_var.set(os.path.join(folder, "PDFBatchPrinterPro"))

    # ── Step 3: Additional Options ───────────────────────────────────

    def _render_step3_options(self):
        self.btn_back.configure(state="normal")
        self.btn_next.configure(text="Cài đặt 🚀", state="normal")

        wrap = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        wrap.pack(fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(
            wrap, text="Tùy Chọn Lối Tắt & Bổ Sung",
            font=ctk.CTkFont(family="Segoe UI", size=17, weight="bold"),
            text_color="#0F172A",
        ).pack(anchor="w", pady=(0, 6))

        ctk.CTkLabel(
            wrap, text="Chọn các biểu tượng lối tắt bạn muốn tạo cho ứng dụng:",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#475569",
        ).pack(anchor="w", pady=(0, 16))

        opts_card = ctk.CTkFrame(wrap, fg_color="#FFFFFF", corner_radius=8, border_width=1, border_color="#CBD5E1")
        opts_card.pack(fill="x", pady=(0, 10), padx=2)

        ctk.CTkCheckBox(
            opts_card, text="Tạo biểu tượng lối tắt ngoài Màn hình nền (Desktop Shortcut)",
            variable=self.desktop_shortcut_var,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#0F172A",
            checkbox_width=20, checkbox_height=20,
        ).pack(anchor="w", padx=16, pady=(14, 8))

        ctk.CTkCheckBox(
            opts_card, text="Tạo lối tắt trong Menu Bắt Đầu (Start Menu Programs)",
            variable=self.start_menu_shortcut_var,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#0F172A",
            checkbox_width=20, checkbox_height=20,
        ).pack(anchor="w", padx=16, pady=(8, 14))

    # ── Step 4: Installing ───────────────────────────────────────────

    def _render_step4_installing(self):
        self.btn_back.configure(state="disabled")
        self.btn_next.configure(state="disabled")
        self.btn_cancel.configure(state="disabled")

        wrap = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        wrap.pack(fill="both", expand=True, padx=10, pady=20)

        ctk.CTkLabel(
            wrap, text="Đang Tiến Hành Cài Đặt...",
            font=ctk.CTkFont(family="Segoe UI", size=17, weight="bold"),
            text_color="#0F172A",
        ).pack(anchor="w", pady=(0, 6))

        self.status_install_lbl = ctk.CTkLabel(
            wrap, text="Đang chuẩn bị trích xuất tệp tin...",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#475569",
        )
        self.status_install_lbl.pack(anchor="w", pady=(0, 16))

        self.progress_bar = ctk.CTkProgressBar(wrap, height=14, corner_radius=7, progress_color="#2563EB")
        self.progress_bar.pack(fill="x", pady=(0, 10))
        self.progress_bar.set(0)

        # Start thread
        threading.Thread(target=self._run_installation, daemon=True).start()

    def _run_installation(self):
        dest_dir = self.install_dir_var.get().strip()
        os.makedirs(dest_dir, exist_ok=True)

        bundle = get_bundle_dir()

        # Update progress
        self._update_prog(0.2, "Đang sao chép các tệp tin ứng dụng...")

        # 1. Copy Main Executable
        src_exe = os.path.join(bundle, "PDFBatchPrinterPro.exe")
        target_exe = os.path.join(dest_dir, "PDFBatchPrinterPro.exe")
        if os.path.exists(src_exe):
            shutil.copy2(src_exe, target_exe)
        else:
            # Fallback for dev mode
            dev_exe = os.path.join(os.path.dirname(bundle), "dist", "PDFBatchPrinterPro.exe")
            if os.path.exists(dev_exe):
                shutil.copy2(dev_exe, target_exe)

        # 2. Copy Assets & User Guide
        self._update_prog(0.5, "Đang cài đặt tài liệu hướng dẫn, biểu tượng và logo...")
        dest_assets = os.path.join(dest_dir, "assets")
        os.makedirs(dest_assets, exist_ok=True)

        src_assets = os.path.join(bundle, "assets")
        if os.path.exists(src_assets):
            for item in os.listdir(src_assets):
                sp = os.path.join(src_assets, item)
                dp = os.path.join(dest_assets, item)
                if os.path.isfile(sp):
                    shutil.copy2(sp, dp)

        # Copy HTML user guide to install root
        for guide_src in [
            os.path.join(bundle, "HUONG_DAN_SU_DUNG.html"),
            os.path.join(bundle, "assets", "HUONG_DAN_SU_DUNG.html"),
            os.path.join(os.path.dirname(bundle), "HUONG_DAN_SU_DUNG.html"),
        ]:
            if os.path.exists(guide_src):
                shutil.copy2(guide_src, os.path.join(dest_dir, "HUONG_DAN_SU_DUNG.html"))
                break

        # 3. Create uninstaller in target dir
        self._update_prog(0.7, "Đang tạo trình gỡ cài đặt (Uninstaller)...")
        uninstaller_exe = os.path.join(dest_dir, "uninstall.exe")
        src_uninst = os.path.join(bundle, "uninstall.exe")
        if os.path.exists(src_uninst):
            shutil.copy2(src_uninst, uninstaller_exe)
        else:
            # Create lightweight uninstall batch script if exe not prebuilt
            uninst_bat = os.path.join(dest_dir, "uninstall.bat")
            local_app = os.environ.get("LOCALAPPDATA", "")
            roaming_app = os.environ.get("APPDATA", "")
            with open(uninst_bat, "w", encoding="utf-8") as fp:
                fp.write(f'''@echo off
title Go cai dat {APP_NAME}
echo Dang go cai dat va xoa toan bo du lieu {APP_NAME}...
taskkill /F /IM PDFBatchPrinterPro.exe >nul 2>&1
timeout /t 1 /nobreak >nul
reg delete "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\PDFBatchPrinterPro" /f >nul 2>&1
reg delete "HKCU\\Software\\PDFBatchPrinterPro" /f >nul 2>&1
del /f /q "%USERPROFILE%\\Desktop\\{APP_NAME}.lnk" >nul 2>&1
rd /s /q "%APPDATA%\\Microsoft\\Windows\\Start Menu\\Programs\\{APP_NAME}" >nul 2>&1
rd /s /q "{local_app}\\PDFBatchPrinterPro" >nul 2>&1
rd /s /q "{roaming_app}\\PDFBatchPrinterPro" >nul 2>&1
rd /s /q "{dest_dir}" >nul 2>&1
echo Da xoa sach toan bo du lieu va hoan tat go cai dat!
pause
''')

        # 4. Create Shortcuts
        self._update_prog(0.85, "Đang tạo các biểu tượng lối tắt...")
        ico_dest = os.path.join(dest_assets, "app.ico")

        try:
            import win32com.client
            shell = win32com.client.Dispatch("WScript.Shell")

            if self.desktop_shortcut_var.get():
                desktop_dir = shell.SpecialFolders("Desktop")
                shortcut_file = os.path.join(desktop_dir, f"{APP_NAME}.lnk")
                create_shortcut(target_exe, shortcut_file, ico_dest, "In ấn tài liệu hàng loạt chuyên nghiệp")

            if self.start_menu_shortcut_var.get():
                programs_dir = shell.SpecialFolders("Programs")
                app_menu_dir = os.path.join(programs_dir, APP_NAME)
                os.makedirs(app_menu_dir, exist_ok=True)
                shortcut_file = os.path.join(app_menu_dir, f"{APP_NAME}.lnk")
                create_shortcut(target_exe, shortcut_file, ico_dest, "In ấn tài liệu hàng loạt chuyên nghiệp")
                
                # Start menu shortcut to user guide
                guide_file = os.path.join(dest_dir, "HUONG_DAN_SU_DUNG.html")
                if os.path.exists(guide_file):
                    guide_lnk = os.path.join(app_menu_dir, "Hướng Dẫn Sử Dụng.lnk")
                    create_shortcut(guide_file, guide_lnk, ico_dest, "Xem tài liệu hướng dẫn sử dụng")
        except Exception as e:
            print(f"Error creating shortcuts: {e}")

        # 5. Register in Windows Registry
        self._update_prog(0.95, "Đang đăng ký ứng dụng với hệ thống Windows...")
        register_uninstall(dest_dir, target_exe, uninstaller_exe if os.path.exists(uninstaller_exe) else target_exe, ico_dest)

        import time
        time.sleep(0.5)
        self._update_prog(1.0, "Cài đặt hoàn tất!")
        self.after(300, lambda: self._show_step(5))

    def _update_prog(self, val: float, text: str):
        self.after(0, lambda: (
            self.progress_bar.set(val),
            self.status_install_lbl.configure(text=text)
        ))

    # ── Step 5: Complete ─────────────────────────────────────────────

    def _render_step5_complete(self):
        self.btn_back.pack_forget()
        self.btn_cancel.pack_forget()
        self.btn_next.configure(text="Hoàn Tất ✓", state="normal", fg_color="#16A34A", hover_color="#15803D")

        wrap = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        wrap.pack(fill="both", expand=True)

        ctk.CTkLabel(
            wrap, text="🎉",
            font=ctk.CTkFont(size=44),
        ).pack(pady=(12, 4))

        ctk.CTkLabel(
            wrap, text=f"Cài Đặt {APP_NAME} Thành Công!",
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
            text_color="#0F172A",
        ).pack(pady=(0, 8))

        ctk.CTkLabel(
            wrap,
            text=f"Ứng dụng đã được cài đặt hoàn tất vào máy tính của bạn.\n"
                 f"Bạn có thể mở ứng dụng bất kỳ lúc nào từ màn hình Desktop hoặc Start Menu.",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#475569", justify="center",
        ).pack(pady=(0, 16))

        cb_card = ctk.CTkFrame(wrap, fg_color="#FFFFFF", corner_radius=8, border_width=1, border_color="#CBD5E1")
        cb_card.pack(fill="x", padx=40, pady=(0, 10))

        ctk.CTkCheckBox(
            cb_card, text="📖 Mở tài liệu Hướng Dẫn Sử Dụng (Quick Start Guide)",
            variable=self.open_guide_var,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="#1E40AF",
            checkbox_width=20, checkbox_height=20,
        ).pack(padx=16, pady=(12, 6), anchor="w")

        ctk.CTkCheckBox(
            cb_card, text=f"🚀 Khởi chạy {APP_NAME} ngay bây giờ",
            variable=self.launch_after_var,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="#1E293B",
            checkbox_width=20, checkbox_height=20,
        ).pack(padx=16, pady=(6, 12), anchor="w")

    # ── Navigation Buttons ───────────────────────────────────────────

    def _next_step(self):
        if self.step == 1:
            self._show_step(2)
        elif self.step == 2:
            dest = self.install_dir_var.get().strip()
            if not dest:
                from tkinter import messagebox
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn thư mục cài đặt hợp lệ.")
                return
            self._show_step(3)
        elif self.step == 3:
            self._show_step(4)
        elif self.step == 5:
            dest_dir = self.install_dir_var.get().strip()
            # Open Guide if checked
            if self.open_guide_var.get():
                guide_file = os.path.join(dest_dir, "HUONG_DAN_SU_DUNG.html")
                if os.path.exists(guide_file):
                    try:
                        os.startfile(guide_file)
                    except Exception:
                        import webbrowser
                        webbrowser.open(f"file:///{guide_file}")

            # Launch app if checked
            if self.launch_after_var.get():
                target_exe = os.path.join(dest_dir, "PDFBatchPrinterPro.exe")
                if os.path.exists(target_exe):
                    try:
                        subprocess.Popen([target_exe], close_fds=True)
                    except Exception as exc:
                        print(f"Error launching target app: {exc}")
            self.quit()
            self.destroy()

    def _prev_step(self):
        if self.step > 1 and self.step < 4:
            self._show_step(self.step - 1)

    def _cancel(self):
        from tkinter import messagebox
        if messagebox.askyesno("Hủy cài đặt", "Bạn có chắc muốn thoát khỏi trình cài đặt?"):
            self.quit()
            self.destroy()


def main():
    app = SetupWizardApp()
    app.mainloop()


if __name__ == "__main__":
    main()
