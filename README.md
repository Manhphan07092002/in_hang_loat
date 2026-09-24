# 🖨️ PDF Batch Printer Pro — v1.0.0.2

Phần mềm Windows chuyên nghiệp để **in ấn hàng loạt tệp tài liệu đa định dạng (PDF, Word, Excel, PowerPoint, Ảnh)** với giao diện hiện đại, tối ưu tốc độ và an toàn quy trình.

---

## ✨ Tính Năng Nổi Bật

1. **Hỗ trợ đa định dạng tài liệu**:
   - PDF (`.pdf`)
   - Microsoft Word (`.doc`, `.docx`, `.rtf`)
   - Microsoft Excel (`.xls`, `.xlsx`)
   - Microsoft PowerPoint (`.ppt`, `.pptx`)
   - Định dạng ảnh (`.jpg`, `.jpeg`, `.png`, `.bmp`, `.tiff`, `.tif`, `.gif`, `.webp`)
   - File Office/ảnh được tự động chuyển sang PDF (tái dùng Office COM, có timeout chống treo).

2. **Kéo & Thả Trực Tiếp (Drag & Drop)**:
   - Kéo thả file hoặc cả thư mục từ Desktop / File Explorer vào ứng dụng để nạp nhanh.
   - Tự động loại file trùng lặp.

3. **Chỉnh Sửa Số Bản In Linh Hoạt & An Toàn**:
   - Sửa trực tiếp trên từng file: click cột Số bản, double-click, `F2`, `Enter`, `+`, `-`, phím số `1-9`, nút `✏️ Sửa Số Bản` hoặc menu chuột phải.
   - **Khóa an toàn**: không cho sửa file đang in hoặc đã in xong.

4. **In Chọn Lọc & In Lại Tệp Lỗi**:
   - **🎯 Chỉ In Mục Chọn**: chọn một nhóm file để in riêng.
   - **🔄 In Lại Tệp Lỗi / Hủy**: 1-click lọc và in lại toàn bộ file lỗi/hủy.

5. **Sắp Xếp & Quản Lý Hàng Đợi**:
   - Di chuyển file lên/xuống (`▲`, `▼`, `Ctrl+Up/Down`).
   - Click tiêu đề cột để sắp xếp (Tên, Số trang, Định dạng, Kích thước, Số bản...).
   - Lưu / mở phiên hàng đợi (`.json`) bằng `💾 Lưu DS`, `📂 Mở DS`.

6. **Tự Động Lưu Cài Đặt**:
   - Ghi nhớ máy in, in 2 mặt, khổ giấy (**A3**/A4/A5/Letter/Legal), giao diện Sáng/Tối...
   - Lưu tại `%APPDATA%\PDFBatchPrinterPro\config.json` (không lỗi quyền Program Files, tự migration config cũ).

7. **Trình Xem Trước**:
   - Xem độ phân giải cao, chuyển trang, zoom 25–400%, vừa trang/vừa ngang, xoay 90°, kéo-pan, phím tắt đầy đủ.

8. **In Nâng Cao**:
   - In song song nhiều máy, tự chuyển máy dự phòng (failover) khi lỗi.
   - Bỏ trang trắng tự động, tờ bìa phân cách, lề đóng gáy, đảo thứ tự trang, vừa trang giấy.

9. **Giao Diện CLI**: chạy tương tác hoặc tích hợp script tự động hóa.

---

## 💻 Yêu Cầu Hệ Thống

- Windows 10/11 (64-bit), Python 3.12+ (nếu chạy từ source).
- Máy in đã cài driver trên Windows.
- Microsoft Office (nếu cần in file Word/Excel/PowerPoint — dùng COM automation).

---

## 🚀 Cài Đặt & Khởi Chạy

### Cách 1: Dùng bản dựng sẵn (khuyên dùng)
1. Chạy `dist/PDFBatchPrinterPro_Setup.exe`, làm theo wizard (tạo shortcut Desktop/Start Menu).
2. Mở app từ shortcut **PDF Batch Printer Pro**.

### Cách 2: Chạy từ source
```bash
pip install -r requirements.txt

# Giao diện đồ họa:
python main.py
# hoặc nháy đúp run.bat

# Giao diện dòng lệnh (tương tác):
python cli.py
```

---

## 📖 Hướng Dẫn Sử Dụng GUI (từng bước)

1. **Nạp file**: bấm `+ File` / `+ Thư mục`, hoặc kéo-thả file/thư mục vào cửa sổ. Tích `Quét thư mục con` nếu cần.
2. **Chọn file** trong hàng đợi để xem trước bên phải.
3. **Chỉnh số bản in** cho từng file (double-click cột Số bản hoặc phím `1-9`, `+`/`-`).
4. **Cấu hình in**: chọn máy in, số bản mặc định, khổ giấy, chiều in (Tự động/Dọc/Ngang), 1 mặt / 2 mặt, vừa trang, các tùy chọn thông minh (bỏ trang trắng, tờ bìa, đảo thứ tự, lề gáy), máy in dự phòng.
5. **Bấm `▶ BẮT ĐẦU IN`** → kiểm tra hộp thoại xác nhận → in. Theo dõi 2 thanh tiến trình (từng file + tổng).
6. **Tạm dừng / Hủy** bất cứ lúc nào bằng `⏸` / `⏹`. File lỗi → bấm `🔄 In Lại Tệp Lỗi`.
7. **Lưu phiên** hàng đợi (`💾 Lưu DS`) để dùng lại hôm sau.

### 🛠️ Phím tắt trong hàng đợi

| Phím tắt | Thao tác |
|---|---|
| `Click đúp` / `F2` / `Enter` | Sửa trực tiếp số bản in |
| `+` / `=` | Tăng +1 bản in |
| `-` / `_` | Giảm -1 bản in |
| `1` → `9` | Đặt nhanh số bản (1–9) |
| `Ctrl+Up` / `Ctrl+Down` | Di chuyển file lên/xuống |
| `Ctrl+A` | Chọn toàn bộ |
| `Delete` | Xóa file đang chọn |
| `Chuột phải` | Menu thao tác nhanh |

---

## ⌨️ Hướng Dẫn CLI

```bash
# Tương tác từng bước:
python cli.py

# Liệt kê máy in:
python cli.py -l

# In trực tiếp nhiều file:
python cli.py -f baocao.pdf hopdong.docx -p "HP LaserJet Pro 4003" -c 2 --duplex long

# In cả thư mục (kèm thư mục con):
python cli.py --dir "D:/TaiLieu" -r -p "Canon LBP2900" --paper A3

# Tùy chọn đầy đủ:
python cli.py -f a.pdf -p "HP LaserJet" -c 2 --pages 1-5,8 --duplex short --paper A4 --orient auto
```

| Tham số | Ý nghĩa |
|---|---|
| `-f/--files` | Danh sách file cần in |
| `-d/--dir`, `-r` | Thư mục cần in, `-r` quét cả thư mục con |
| `-p/--printer` | Tên máy in (mặc định: máy in Default) |
| `-c/--copies` | Số bản in (mặc định 1) |
| `--pages` | Phạm vi trang, vd `1-5,8` (mặc định in tất cả) |
| `--duplex` | `1` (1 mặt), `long` (lật sách), `short` (lật lịch) |
| `--paper` | `A3`, `A4`, `A5`, `Letter`, `Legal` |
| `--orient` | `auto`, `portrait`, `landscape` |
| `--no-fit` | Tắt tự co vừa trang giấy |
| `-l/--list-printers` | Liệt kê máy in rồi thoát |

---

## 📦 Đóng Gói EXE / Setup

```bash
# EXE độc lập (ra dist/PDFBatchPrinterPro.exe):
python build_exe.py
# hoặc nháy đúp build.bat

# Bộ cài đặt Setup wizard (EXE + uninstall + Setup):
python build_installer.py
```
> ⚠️ Tắt app đang chạy trước khi build, nếu không PyInstaller báo `WinError 5 Access denied` do file `dist/*.exe` bị khóa.

---

## ✅ Kiểm Thử

```bash
python -m pytest test_full.py -q        # 60 test toàn diện (khuyên dùng)
python -m pytest test_suite.py -q       # bộ test gốc
```
5 test GUI trong `test_suite.py` cần màn hình + Office/máy in thật nên chỉ chạy trên máy dev Windows đầy đủ.

---

## 📁 Cấu Trúc Mã Nguồn

```
pdf-batch-printer/
├── main.py                 # Entry point GUI (python main.py, --cli để dùng CLI)
├── cli.py                  # Giao diện dòng lệnh (tương tác + args)
├── launch.py               # Launcher desktop
├── build_exe.py            # Đóng gói PyInstaller (+ version_info.txt)
├── build_installer.py      # Đóng gói bộ Setup wizard
├── build.bat / run.bat     # 1-click build / chạy trên Windows
├── installer.iss           # Script Inno Setup (dự phòng)
├── installer_wizard.py     # Setup wizard (APP_VERSION 1.0.0.1)
├── uninstaller.py          # Trình gỡ cài đặt
├── version_info.txt        # Metadata version cho file EXE
├── requirements.txt        # Thư viện ghim version
├── test_full.py            # Bộ test toàn diện (60 test)
├── test_suite.py           # Bộ test gốc
└── app/
    ├── gui.py              # Giao diện CustomTkinter
    ├── config_store.py     # Config %APPDATA% + migration
    ├── queue_store.py      # Logic hàng đợi thuần (test được headless)
    ├── pdf_manager.py      # Nạp/render/lọc trang/separator
    ├── file_converter.py   # Office (COM dùng chung) / ảnh → PDF
    ├── printer_manager.py  # Win32 GDI: DEVMODE + in raster
    ├── print_worker.py     # Worker đa máy in + failover
    ├── preview.py          # Xem trước nhúng + modal
    ├── settings.py         # Hằng số (APP_VERSION 1.0.0.1, khổ giấy, theme...)
    └── utils.py            # parse_page_range, format_file_size...
```

---

## ❓ Xử Lý Sự Cố Thường Gặp

| Hiện tượng | Cách xử lý |
|---|---|
| Không in được file Office | Cài Microsoft Office bản quyền, tắt hộp thoại Office đang mở |
| Build báo `WinError 5` | Tắt app/setup đang chạy rồi build lại |
| Treo khi chuẩn bị file | File Office nặng — chờ hộp progress, hoặc convert tay sang PDF |
| Config không lưu | Kiểm tra quyền ghi `%APPDATA%\PDFBatchPrinterPro` |
| Máy in Offline/kẹt giấy | Kiểm tra trạng thái ngay dưới tên máy in trong app |

---

## 📝 Lịch Sử Version

- **v1.0.0.2** — Version một nguồn (`settings.py`, tự sinh `version_info.txt` + `installer.iss` khi build), log file xoay vòng (`%APPDATA%/logs/app.log`), khổ giấy tùy chỉnh R×C mm (DMPAPER_USER), `requirements-dev.txt`, gỡ `config.json` khỏi git, thống nhất số liệu thống kê, theme repaint full.
- **v1.0.0.1** — Nhúng version EXE, config `%APPDATA%` + migration, khổ A3, tái dùng Office COM (timeout 180s), tách `config_store`/`queue_store`, hỗ trợ `.webp`, in nền không treo UI, bộ test 60 case.
- **v1.0.0** — Bản đầu tiên: GUI + CLI, in đa định dạng, preview, failover, setup wizard.
