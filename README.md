# 🧾 VN-InvoiceParser

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](./LICENSE)
[![Python Version](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![Security Policy](https://img.shields.io/badge/Security-Policy-brightgreen.svg)](./SECURITY.md)
[![Last Commit](https://img.shields.io/github/last-commit/lequoccuong141/VN-InvoiceParser)](https://github.com/lequoccuong141/VN-InvoiceParser)

Ứng dụng web mã nguồn mở xử lý tài liệu thông minh (**IDP - Intelligent Document Processing**) chuyên biệt cho **hóa đơn điện tử Việt Nam** (theo Nghị định 123/2020/NĐ-CP & Thông tư 78/2021/TT-BTC).

Cho phép tải lên hàng loạt hóa đơn (**PDF, JPG, PNG, XML**), tự động trích xuất dữ liệu chuẩn hóa JSON bằng **Google Gemini 2.5 Flash (Structured Outputs)** hoặc **XML Parser gốc**, tự động đối chiếu sai số tiền / MST / trùng lặp, cho phép chỉnh sửa trực tiếp trên giao diện và xuất file **Excel kế toán 2 sheet** có công thức tính tổng.

---

## ✨ Tính năng chính

1. **Tiền xử lý & Tự động phân loại thông minh (Ingestion):**
   - **XML (NĐ123 / TT78):** Parse trực tiếp bằng XML parser chuẩn Tổng cục Thuế — chính xác 100%, không tốn chi phí API.
   - **PDF có sẵn lớp text:** Tự động phát hiện bằng `PyMuPDF (fitz)` và trích xuất trực tiếp text gốc để đưa qua Gemini Structured Outputs (hoặc Regex fallback).
   - **PDF scan & Ảnh (JPG/PNG):** Chuyển đổi sang ảnh PNG độ phân giải cao (250 DPI), tự động căn chỉnh chiều EXIF và trích xuất bằng **Gemini 2.5 Flash Vision** (hoặc Tesseract OCR fallback).
2. **Trích xuất chuẩn hóa JSON Schema:**
   - Đầy đủ 12 trường cốt lõi: `invoice_symbol`, `invoice_number`, `invoice_date` (`DD/MM/YYYY`), `seller_name`, `seller_tax_code`, `buyer_name`, `buyer_tax_code`, `sub_total`, `vat_rate`, `vat_amount`, `total_amount`, và danh sách chi tiết `items`.
3. **Đối chiếu & Kiểm soát lỗi tự động (Validation & Reconciliation):**
   - Kiểm tra logic toán học: `sub_total + vat_amount == total_amount` (hỗ trợ cấu hình dung sai làm tròn VNĐ).
   - Kiểm tra định dạng Mã số thuế Việt Nam (10 số hoặc `10 số-3 số`).
   - Phát hiện trùng lặp hóa đơn theo bộ khóa `(invoice_symbol, invoice_number, seller_tax_code)` trong cùng lô tải lên.
   - Tự động gắn nhãn `Cần kiểm tra` và tô màu đỏ cảnh báo trên cả giao diện web và file Excel.
4. **Chỉnh sửa trực tiếp & Xuất Excel chuẩn sổ sách kế toán:**
   - Sửa trực tiếp bảng kê tổng hợp và chi tiết hàng hóa ngay trên Streamlit, tự động đối chiếu lại sau khi sửa.
   - Xuất file `.xlsx` 2 sheet bằng `openpyxl`:
     - **Sheet 1 (`Bảng kê hóa đơn tổng hợp`):** Border chuẩn, format tiền `#,##0`, tô đỏ dòng lỗi, có dòng **TỔNG CỘNG** sử dụng công thức `=SUM(...)`.
     - **Sheet 2 (`Chi tiết hàng hóa`):** Toàn bộ mặt hàng liên kết theo Ký hiệu/Số hóa đơn, kèm dòng tổng cộng.
5. **Tích hợp sẵn FastAPI REST API:** Dễ dàng mở rộng sang Frontend Next.js hoặc tích hợp vào hệ thống ERP/MISA/FAST.

---

## 📂 Cấu trúc thư mục dự án

```text
VN-InvoiceParser/
├── app.py                  # Giao diện Streamlit (Upload, Edit, Validate, Export)
├── requirements.txt        # Danh sách thư viện Python
├── .env.example            # Mẫu biến môi trường (GEMINI_API_KEY, DPI, Tolerance)
├── .gitignore              # Quy tắc loại trừ git (chặn .env, .venv, *.xlsx)
├── LICENSE                 # Bản quyền mã nguồn mở (MIT License)
├── SECURITY.md             # Chính sách bảo mật & báo cáo lỗ hổng
├── README.md               # Hướng dẫn cài đặt và vận hành
├── src/
│   ├── __init__.py
│   ├── config.py           # Quản lý cấu hình từ .env
│   ├── schemas.py          # Pydantic JSON Schema chuẩn hóa hóa đơn Việt Nam
│   ├── extractor.py        # Lõi phân loại file, XML parser, PyMuPDF, Gemini 2.5 Flash & Fallback
│   ├── validator.py        # Kiểm tra logic toán học, MST và trùng lặp lô hóa đơn
│   ├── exporter.py         # Xuất Excel 2 sheet chuẩn kế toán bằng OpenPyXL
│   └── api.py              # FastAPI REST API backend
└── tests/
    ├── __init__.py
    └── test_pipeline.py    # Bộ kiểm thử tự động (pytest)
```

---

## 🚀 Hướng dẫn cài đặt & Chạy dự án

### 1. Yêu cầu môi trường
- Python **3.10 trở lên** (khuyên dùng 3.11 hoặc 3.12).
- **Gemini API Key** (lấy miễn phí tại [Google AI Studio](https://aistudio.google.com/apikey)).
- *(Tùy chọn)* Cài đặt [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) kèm gói ngôn ngữ `vie` nếu bạn muốn dùng chế độ Offline Fallback cho ảnh/PDF scan.

### 2. Tạo môi trường ảo & Cài đặt thư viện

**Windows (PowerShell):**
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

**macOS / Linux:**
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 3. Cấu hình biến môi trường

Sao chép file `.env.example` thành `.env` và điền `GEMINI_API_KEY`:

```powershell
Copy-Item .env.example .env
```

Nội dung `.env`:
```env
GEMINI_API_KEY=AIzaSy...your_key_here
GEMINI_MODEL=gemini-2.5-flash
PDF_RENDER_DPI=250
AMOUNT_TOLERANCE=1.0
```

> **Mẹo:** Bạn cũng có thể nhập trực tiếp `GEMINI_API_KEY` trên thanh bên (Sidebar) của giao diện Streamlit, hoặc nhấn nút **"Nạp mẫu Demo"** để trải nghiệm ngay mà không cần API Key.

### 4. Chạy giao diện web (Streamlit)

```bash
streamlit run app.py
```

Mở trình duyệt tại địa chỉ: `http://localhost:8501`

### 5. (Tùy chọn) Chạy Backend API (FastAPI)

```bash
uvicorn src.api:app --reload --port 8000
```

Truy cập tài liệu Swagger UI tương tác tại: `http://localhost:8000/docs`

### 6. Chạy kiểm thử tự động (Unit Tests)

```bash
pytest -v
```
