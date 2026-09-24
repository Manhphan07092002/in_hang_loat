# 🖨️ PDF Batch Printer Pro

Phần mềm Windows chuyên nghiệp để **in ấn hàng loạt tệp tài liệu đa định dạng (PDF, Word, Excel, PowerPoint, Ảnh)** với giao diện hiện đại, tối ưu tốc độ và an toàn quy trình.

---

## ✨ Tính Năng Nổi Bật

1. **Hỗ trợ đa định dạng tài liệu**:
   - PDF (`.pdf`)
   - Microsoft Word (`.docx`, `.doc`, `.rtf`)
   - Microsoft Excel (`.xlsx`, `.xls`)
   - Microsoft PowerPoint (`.pptx`, `.ppt`)
   - Định dạng ảnh (`.jpg`, `.png`, `.bmp`, `.tiff`, `.gif`)

2. **Kéo & Thả Trực Tiếp (Drag & Drop)**:
   - Kéo thả file hoặc cả thư mục trực tiếp từ Desktop / File Explorer vào ứng dụng để nạp nhanh.

3. **Chỉnh Sửa Số Bản In Linh Hoạt & An Toàn (Per-File Copies & Print State Lock)**:
   - Sửa số bản in trực tiếp trên từng file (Click đơn cột Số bản, Double click, bấm phím `F2`, `Enter`, `+`, `-`, phím số `1-9`, nút `✏️ Sửa Số Bản` hoặc Menu chuột phải).
   - **Cơ chế khóa an toàn**: Tự động khóa không cho chỉnh sửa đối với file đang in (`Đang in...`) hoặc đã in xong (`Đã in xong`).

4. **In Chọn Lọc & Tự Động In Lại Tệp Lỗi (Selective & Retry Print)**:
   - **🎯 Chỉ In Mục Chọn**: Cho phép chọn một nhóm tệp cụ thể trong danh sách để in riêng.
   - **🔄 In Lại Tệp Lỗi / Hủy**: 1-click tự động lọc và in lại toàn bộ các tệp bị lỗi hoặc bị hủy.

5. **Sắp Xếp & Quản Lý Hàng Đợi Nâng Cao**:
   - Di chuyển vị trí file lên / xuống (`▲`, `▼`, `Ctrl+Up`, `Ctrl+Down`, `Alt+Up`, `Alt+Down`).
   - Sắp xếp nhanh danh sách theo cột bằng cách click vào tiêu đề cột (Tên file, Số trang, Định dạng, Kích thước, Số bản...).
   - Lưu & Mở phiên hàng đợi in ra tệp `.json` (`💾 Lưu DS`, `📂 Mở DS`).

6. **Tự Động Lưu Cài Đặt (Settings Persistence)**:
   - Tự động ghi nhớ máy in, chế độ in 2 mặt (lật sách / lật lịch), khổ giấy, giao diện Sáng / Tối trong `config.json`.

7. **Trình Xem Trước (Interactive PDF Preview)**:
   - Xem trước độ phân giải cao, chuyển trang, phóng to / thu nhỏ.

8. **Giao Diện CLI Mạnh Mẽ**:
   - Chạy dòng lệnh tương tác hoặc tích hợp script tự động hóa.

---

## 🚀 Cài Đặt & Khởi Chạy

### 1. Cài đặt thư viện
```bash
pip install -r requirements.txt
```

### 2. Khởi chạy Giao diện đồ họa (GUI)
- Nháy đúp vào tệp **`run.bat`** trên Desktop hoặc chạy:
```bash
python main.py
```

### 3. Giao diện Dòng lệnh (CLI)
```bash
python cli.py
```

### 4. Đóng gói thành file `.exe` độc lập
- Nháy đúp vào tệp **`build.bat`** hoặc chạy:
```bash
python build_exe.py
```
- File `.exe` độc lập sẽ được tạo tại: `dist/PDFBatchPrinterPro.exe`.

---

## 🛠️ Phím Tắt Tiện Lợi Trong Hàng Đợi

| Phím tắt | Thao tác |
|---|---|
| `Click đúp` / `F2` / `Enter` | Sửa trực tiếp số bản in cho file đang chọn |
| `+` / `=` | Tăng +1 bản in |
| `-` / `_` | Giảm -1 bản in |
| `1` → `9` | Đặt nhanh số bản in (1 đến 9 bản) |
| `Ctrl+Up` / `Ctrl+Down` | Di chuyển file lên / xuống trong hàng đợi |
| `Ctrl+A` | Chọn toàn bộ file trong hàng đợi |
| `Delete` | Xóa các file đang chọn khỏi hàng đợi |
| `Chuột phải` | Mở menu thao tác nhanh đầy đủ |

---

## 📁 Cấu Trúc Mã Nguồn

```
pdf-batch-printer/
├── main.py                 # Entry point khởi chạy GUI
├── cli.py                  # Entry point giao diện dòng lệnh
├── launch.py               # Launcher desktop
├── build_exe.py            # Script đóng gói PyInstaller
├── build.bat               # 1-Click đóng gói EXE cho Windows
├── run.bat                 # 1-Click khởi chạy ứng dụng
├── test_suite.py           # Bộ kiểm thử tự động
├── config.json             # Tệp lưu cấu hình người dùng
├── requirements.txt        # Danh sách thư viện phụ thuộc
├── README.md               # Tài liệu hướng dẫn sử dụng
└── app/
    ├── gui.py              # Giao diện đồ họa CustomTkinter
    ├── pdf_manager.py      # Quản lý đọc, render, phân tích file
    ├── file_converter.py   # Chuyển đổi Office/Image sang PDF
    ├── printer_manager.py  # Quản lý giao tiếp Win32 Print API
    ├── print_worker.py     # Worker thread in nền đa luồng
    ├── preview.py          # Khung xem trước tài liệu
    ├── settings.py         # Hằng số, màu sắc, chế độ in
    └── utils.py            # Hàm tiện ích xử lý chuỗi và số
```
