"""
Build script to package PDF Batch Printer Pro into a standalone Windows .exe file using PyInstaller.
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


def build():
    print("=" * 60)
    print("[BUILD] DANG DONG GOI UNG DUNG PDF BATCH PRINTER PRO")
    print("=" * 60)

    base_dir = os.path.dirname(os.path.abspath(__file__))
    dist_dir = os.path.join(base_dir, "dist")
    build_dir = os.path.join(base_dir, "build")

    main_script = os.path.join(base_dir, "main.py")

    args = [
        main_script,
        "--name=PDFBatchPrinterPro",
        "--onefile",
        "--noconsole",
        "--clean",
        "--collect-all=customtkinter",
        "--collect-all=pymupdf",
        "--collect-all=windnd",
        "--hidden-import=win32print",
        "--hidden-import=win32ui",
        "--hidden-import=win32gui",
        "--hidden-import=win32con",
        "--hidden-import=pythoncom",
        "--hidden-import=win32com.client",
        "--hidden-import=PIL",
        "--hidden-import=PIL.Image",
        "--hidden-import=PIL.ImageTk",
        "--exclude-module=tensorflow",
        "--exclude-module=torch",
        "--exclude-module=matplotlib",
        "--exclude-module=scipy",
        "--exclude-module=pandas",
        "--exclude-module=IPython",
        "--exclude-module=jupyter",
        "--exclude-module=cv2",
        "--exclude-module=sklearn",
        f"--distpath={dist_dir}",
        f"--workpath={build_dir}",
    ]

    # Include assets directory if present
    assets_dir = os.path.join(base_dir, "assets")
    if os.path.exists(assets_dir):
        args.append(f"--add-data={assets_dir};assets")

    # Include icon if present
    icon_path = os.path.join(assets_dir, "app.ico")
    if os.path.exists(icon_path):
        args.append(f"--icon={icon_path}")

    # Embed Windows file version (1.0.0.1)
    version_file = os.path.join(base_dir, "version_info.txt")
    if os.path.exists(version_file):
        args.append(f"--version-file={version_file}")

    print("Tham so PyInstaller:")
    for a in args:
        print(f"  {a}")
    print("\n[INFO] Qua trinh bien dich dang bat dau, vui long doi...")

    try:
        PyInstaller.__main__.run(args)
        exe_path = os.path.join(dist_dir, "PDFBatchPrinterPro.exe")
        if os.path.exists(exe_path):
            print("\n" + "=" * 60)
            print("[SUCCESS] BIEN DICH THANH CONG!")
            print(f"File exe: {exe_path}")
            print(f"Dung luong: {os.path.getsize(exe_path) / (1024 * 1024):.2f} MB")
            print("=" * 60)
            return 0
        else:
            print("\n[ERROR] Khong tim thay file .exe sau khi bien dich.")
            return 1
    except Exception as exc:
        print(f"\n[ERROR] Loi trong qua trinh dong goi: {exc}")
        return 1


if __name__ == "__main__":
    sys.exit(build())
