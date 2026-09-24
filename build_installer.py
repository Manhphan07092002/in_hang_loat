"""
Build script to compile the complete Windows Installer (PDFBatchPrinterPro_Setup.exe) using PyInstaller.
"""
import os
import sys
import shutil

# Ensure UTF-8 output encoding on Windows consoles
if sys.stdout is not None:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
if sys.stderr is not None:
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import PyInstaller.__main__


def build_installer():
    print("=" * 60)
    print("📦 BẮT ĐẦU ĐÓNG GÓI BỘ CÀI ĐẶT WINDOWS (SETUP WIZARD)")
    print("=" * 60)

    base_dir = os.path.dirname(os.path.abspath(__file__))
    dist_dir = os.path.join(base_dir, "dist")
    build_dir = os.path.join(base_dir, "build")
    assets_dir = os.path.join(base_dir, "assets")

    main_exe = os.path.join(dist_dir, "PDFBatchPrinterPro.exe")
    print("⏳ Tiến hành biên dịch ứng dụng chính (PDFBatchPrinterPro.exe) trước...")
    import build_exe
    res = build_exe.build()
    if res != 0:
        print("❌ Lỗi khi tạo PDFBatchPrinterPro.exe")
        return 1

    # 1. Build Uninstaller
    print("\n🔨 Đang đóng gói trình gỡ cài đặt (uninstall.exe)...")
    uninst_script = os.path.join(base_dir, "uninstaller.py")
    uninst_args = [
        uninst_script,
        "--name=uninstall",
        "--onefile",
        "--noconsole",
        "--clean",
        "--collect-all=customtkinter",
        f"--distpath={dist_dir}",
        f"--workpath={build_dir}",
    ]
    icon_path = os.path.join(assets_dir, "app.ico")
    if os.path.exists(icon_path):
        uninst_args.append(f"--icon={icon_path}")

    try:
        PyInstaller.__main__.run(uninst_args)
    except Exception as exc:
        print(f"Lỗi build uninstaller: {exc}")

    # 2. Build Setup Wizard
    print("\n🔨 Đang đóng gói Trình Cài Đặt (PDFBatchPrinterPro_Setup.exe)...")
    setup_script = os.path.join(base_dir, "installer_wizard.py")
    uninst_exe = os.path.join(dist_dir, "uninstall.exe")

    setup_args = [
        setup_script,
        "--name=PDFBatchPrinterPro_Setup",
        "--onefile",
        "--noconsole",
        "--clean",
        "--collect-all=customtkinter",
        "--hidden-import=win32com.client",
        "--hidden-import=pythoncom",
        "--hidden-import=winreg",
        "--hidden-import=PIL",
        "--hidden-import=PIL.Image",
        "--hidden-import=PIL.ImageTk",
        f"--add-data={main_exe};.",
        f"--add-data={assets_dir};assets",
        f"--distpath={dist_dir}",
        f"--workpath={build_dir}",
    ]

    if os.path.exists(uninst_exe):
        setup_args.append(f"--add-data={uninst_exe};.")

    if os.path.exists(icon_path):
        setup_args.append(f"--icon={icon_path}")

    try:
        PyInstaller.__main__.run(setup_args)
        final_setup = os.path.join(dist_dir, "PDFBatchPrinterPro_Setup.exe")
        if os.path.exists(final_setup):
            print("\n" + "=" * 60)
            print("🎉 TẠO BỘ CÀI ĐẶT SETUP HOÀN TẤT THÀNH CÔNG!")
            print(f"📁 Tệp cài đặt: {final_setup}")
            print(f"📊 Dung lượng: {os.path.getsize(final_setup) / (1024 * 1024):.2f} MB")
            print("=" * 60)
            return 0
        else:
            print("\n❌ Không tìm thấy file PDFBatchPrinterPro_Setup.exe.")
            return 1
    except Exception as exc:
        print(f"\n❌ Lỗi khi đóng gói bộ cài đặt: {exc}")
        return 1


if __name__ == "__main__":
    sys.exit(build_installer())
