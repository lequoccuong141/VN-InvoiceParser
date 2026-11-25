"""Module lõi trích xuất dữ liệu hóa đơn điện tử Việt Nam (XML, PDF Text, PDF Scan, Ảnh).

Hỗ trợ:
- Parse trực tiếp XML hóa đơn điện tử chuẩn Nghị định 123/2020/NĐ-CP & Thông tư 78/2021/TT-BTC.
- Tự động phân loại PDF có lớp text gốc vs PDF scan bằng PyMuPDF (fitz).
- Trích xuất chuẩn hóa JSON Schema bằng Google Gemini 2.5 Flash (Structured Outputs).
- Fallback bằng Tesseract OCR + Regex khi không có Gemini API Key.
"""

from __future__ import annotations

import csv
import io
import json
import logging
import re
import time
import unicodedata
import uuid
import xml.etree.ElementTree as ET
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import fitz  # PyMuPDF
from openpyxl import load_workbook
from PIL import Image, ImageOps

from src.config import Settings, get_settings
from src.schemas import InvoiceData, InvoiceItem, InvoiceRecord, InvoiceStatus

logger = logging.getLogger(__name__)

SUPPORTED_EXTENSIONS = {".pdf", ".png", ".jpg", ".jpeg", ".xml", ".xlsx", ".xlsm", ".csv"}

EXTRACTION_SYSTEM_PROMPT = """Bạn là chuyên gia kế toán và xử lý tài liệu thông minh (IDP) chuyên về hóa đơn điện tử Việt Nam (Nghị định 123/2020/NĐ-CP, Thông tư 78/2021/TT-BTC).

Nhiệm vụ: Trích xuất chính xác toàn bộ thông tin từ hóa đơn và trả về đúng JSON Schema yêu cầu.

Quy tắc bắt buộc:
1. `invoice_symbol`: Ký hiệu hóa đơn (nếu hóa đơn tách riêng "Ký hiệu mẫu số" như `1` và "Ký hiệu" như `C26TAA`, hãy ghép thành `1C26TAA`; nếu đã ghi liền `1C26TAA` thì giữ nguyên).
2. `invoice_number`: Số hóa đơn (giữ nguyên các số 0 ở đầu nếu có, VD: `00001234`).
3. `invoice_date`: Chuẩn hóa bắt buộc về định dạng `DD/MM/YYYY` (VD: `Ngày 5 tháng 3 năm 2026` -> `05/03/2026`).
4. `seller_tax_code` và `buyer_tax_code`: Chỉ giữ chữ số và dấu gạch ngang `-` (định dạng 10 số hoặc `10 số-3 số`, VD: `0101234567` hoặc `0101234567-001`). Bỏ khoảng trắng thừa. Nếu người mua là cá nhân không có MST thì để chuỗi rỗng `""`.
5. Định dạng số tiền (`sub_total`, `vat_amount`, `total_amount`, `quantity`, `unit_price`, `amount`):
   - Hóa đơn Việt Nam thường dùng dấu chấm `.` làm phân cách hàng nghìn (VD: `1.500.000` nghĩa là `1500000`) và dấu phẩy `,` làm phân cách thập phân (VD: `1.500,50` nghĩa là `1500.5`).
   - Hãy chuyển tất cả về số thực (float) chuẩn JSON (không có dấu phân cách hàng nghìn, không có ký hiệu tiền tệ).
   - Nếu không chịu thuế GTGT hoặc thuế suất 0%, `vat_amount` là `0`.
6. `vat_rate`: Chuẩn hóa thành `"0%"`, `"5%"`, `"8%"`, `"10%"`, `"KCT"` (không chịu thuế), `"KKKNT"` (không kê khai nộp thuế), hoặc `"Nhiều thuế suất"` nếu hóa đơn có nhiều mức thuế suất khác nhau.
7. `items`: Trích xuất đầy đủ từng dòng hàng hóa, dịch vụ trong bảng chi tiết (bỏ qua dòng tiêu đề bảng và dòng tổng cộng).
8. Không tự bịa dữ liệu. Nếu trường nào không có trên hóa đơn, trả về `""` cho chuỗi hoặc `0` cho số.
"""


# ---------------------------------------------------------------------------
# Normalization helpers
# ---------------------------------------------------------------------------

def normalize_date(raw_date: str) -> str:
    """Chuẩn hóa chuỗi ngày lập hóa đơn về định dạng DD/MM/YYYY."""
    if not raw_date:
        return ""
    text = raw_date.strip()

    # Dạng "Ngày 05 tháng 03 năm 2026"
    vn_match = re.search(
        r"ng[àa]y\s*(\d{1,2})\s*th[áa]ng\s*(\d{1,2})\s*n[ăa]m\s*(\d{4})",
        text,
        flags=re.IGNORECASE,
    )
    if vn_match:
        d, m, y = int(vn_match.group(1)), int(vn_match.group(2)), int(vn_match.group(3))
        return f"{d:02d}/{m:02d}/{y:04d}"

    # Dạng ISO YYYY-MM-DD hoặc YYYY-MM-DDTHH:MM:SS
    iso_match = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})", text)
    if iso_match:
        y, m, d = int(iso_match.group(1)), int(iso_match.group(2)), int(iso_match.group(3))
        return f"{d:02d}/{m:02d}/{y:04d}"

    # Dạng DD/MM/YYYY hoặc DD-MM-YYYY hoặc DD.MM.YYYY
    dmy_match = re.match(r"^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$", text)
    if dmy_match:
        d, m, y = int(dmy_match.group(1)), int(dmy_match.group(2)), int(dmy_match.group(3))
        return f"{d:02d}/{m:02d}/{y:04d}"

    # Thử parse với datetime
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            dt = datetime.strptime(text[:10], fmt)
            return dt.strftime("%d/%m/%Y")
        except ValueError:
            continue
    return text


def normalize_tax_code(raw_mst: str) -> str:
    """Làm sạch Mã số thuế: bỏ khoảng trắng, chuẩn hóa dấu gạch ngang."""
    if not raw_mst:
        return ""
    cleaned = re.sub(r"\s+", "", raw_mst.strip())
    # Loại bỏ tiền tố như "MST:" nếu dính vào
    cleaned = re.sub(r"^(MST|Tax\s*code)[:.\s]*", "", cleaned, flags=re.IGNORECASE)
    # Chuẩn hóa các ký tự gạch ngang Unicode
    cleaned = cleaned.replace("–", "-").replace("—", "-")
    # Nếu là 13 chữ số liền nhau không có dấu '-', tự tách thành 10-3
    if re.fullmatch(r"\d{13}", cleaned):
        cleaned = f"{cleaned[:10]}-{cleaned[10:]}"
    return cleaned


def normalize_vat_rate(raw_rate: str) -> str:
    """Chuẩn hóa thuế suất GTGT."""
    if not raw_rate:
        return ""
    text = raw_rate.strip().upper()
    if text in {"KCT", "KHONG CHIU THUE", "KHÔNG CHỊU THUẾ"}:
        return "KCT"
    if text in {"KKKNT", "KHONG KE KHAI", "KHÔNG KÊ KHAI"}:
        return "KKKNT"
    # Dạng số như "8", "10.0", "8%", "0.08"
    num_match = re.fullmatch(r"(\d+(?:\.\d+)?)\s*%?", text)
    if num_match:
        val = float(num_match.group(1))
        if 0 < val < 1:
            val *= 100
        if val.is_integer():
            return f"{int(val)}%"
        return f"{val:g}%"
    return raw_rate.strip()


def parse_vn_number(raw_value: str | float | int | None) -> float:
    """Chuyển đổi chuỗi số kiểu Việt Nam hoặc chuẩn quốc tế sang float."""
    if raw_value is None:
        return 0.0
    if isinstance(raw_value, (int, float)):
        return float(raw_value)
    text = str(raw_value).strip()
    if not text:
        return 0.0
    # Xóa ký tự tiền tệ và khoảng trắng
    text = re.sub(r"[^\d,.\-]", "", text)
    if not text or text in {"-", ".", ","}:
        return 0.0

    # Nếu có cả '.' và ','
    if "." in text and "," in text:
        if text.rfind(",") > text.rfind("."):
            # Kiểu VN/EU: 1.234.567,89
            text = text.replace(".", "").replace(",", ".")
        else:
            # Kiểu US: 1,234,567.89
            text = text.replace(",", "")
    elif "." in text:
        parts = text.split(".")
        # Nếu có nhiều dấu chấm (1.500.000) hoặc phần sau có đúng 3 chữ số (500.000) -> phân cách nghìn
        if len(parts) > 2 or (len(parts) == 2 and len(parts[1]) == 3 and len(parts[0].lstrip("-")) <= 3):
            text = text.replace(".", "")
    elif "," in text:
        parts = text.split(",")
        if len(parts) > 2 or (len(parts) == 2 and len(parts[1]) == 3):
            text = text.replace(",", "")
        else:
            text = text.replace(",", ".")

    try:
        return float(text)
    except ValueError:
        return 0.0


def normalize_invoice_data(data: InvoiceData) -> InvoiceData:
    """Chuẩn hóa lại toàn bộ các trường sau khi trích xuất."""
    data.invoice_symbol = re.sub(r"\s+", "", data.invoice_symbol.strip().upper())
    data.invoice_number = data.invoice_number.strip()
    data.invoice_date = normalize_date(data.invoice_date)
    data.seller_name = " ".join(data.seller_name.split())
    data.seller_tax_code = normalize_tax_code(data.seller_tax_code)
    data.buyer_name = " ".join(data.buyer_name.split())
    data.buyer_tax_code = normalize_tax_code(data.buyer_tax_code)
    data.vat_rate = normalize_vat_rate(data.vat_rate)
    data.sub_total = round(float(data.sub_total or 0.0), 2)
    data.vat_amount = round(float(data.vat_amount or 0.0), 2)
    data.total_amount = round(float(data.total_amount or 0.0), 2)

    cleaned_items: List[InvoiceItem] = []
    for item in data.items:
        item.name = " ".join(item.name.split())
        item.unit = item.unit.strip()
        item.quantity = round(float(item.quantity or 0.0), 4)
        item.unit_price = round(float(item.unit_price or 0.0), 2)
        item.amount = round(float(item.amount or 0.0), 2)
        if item.name or item.amount > 0:
            cleaned_items.append(item)
    data.items = cleaned_items
    return data


# ---------------------------------------------------------------------------
# XML Parser (Nghị định 123 / Thông tư 78)
# ---------------------------------------------------------------------------

def _strip_ns(tag: str) -> str:
    """Loại bỏ XML namespace khỏi tên thẻ."""
    return tag.split("}", 1)[-1] if "}" in tag else tag


def _find_child_text(parent: Optional[ET.Element], tag_name: str) -> str:
    """Tìm nội dung text của thẻ con theo tên (bỏ qua namespace)."""
    if parent is None:
        return ""
    for elem in parent.iter():
        if _strip_ns(elem.tag) == tag_name and elem.text:
            return elem.text.strip()
    return ""


def _find_element(root: ET.Element, tag_name: str) -> Optional[ET.Element]:
    """Tìm phần tử đầu tiên khớp tên thẻ (bỏ qua namespace)."""
    for elem in root.iter():
        if _strip_ns(elem.tag) == tag_name:
            return elem
    return None


def parse_vn_invoice_xml(xml_bytes: bytes) -> InvoiceData:
    """Parse hóa đơn điện tử Việt Nam định dạng XML (chuẩn NĐ 123 / TT 78)."""
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        # Thử decode UTF-8-SIG (loại bỏ BOM)
        text = xml_bytes.decode("utf-8-sig", errors="replace").strip()
        root = ET.fromstring(text.encode("utf-8"))

    tt_chung = _find_element(root, "TTChung")
    n_ban = _find_element(root, "NBan")
    n_mua = _find_element(root, "NMua")
    t_toan = _find_element(root, "TToan")

    kh_ms = _find_child_text(tt_chung, "KHMSHDon")
    kh_hd = _find_child_text(tt_chung, "KHHDon")
    invoice_symbol = f"{kh_ms}{kh_hd}" if kh_ms and not kh_hd.startswith(kh_ms) else (kh_hd or kh_ms)
    invoice_number = _find_child_text(tt_chung, "SHDon")
    invoice_date = _find_child_text(tt_chung, "NLap")

    seller_name = _find_child_text(n_ban, "Ten")
    seller_tax_code = _find_child_text(n_ban, "MST")

    buyer_name = _find_child_text(n_mua, "Ten") or _find_child_text(n_mua, "HVTNMHang")
    buyer_tax_code = _find_child_text(n_mua, "MST")

    sub_total = parse_vn_number(_find_child_text(t_toan, "TgTCThue"))
    vat_amount = parse_vn_number(_find_child_text(t_toan, "TgTThue"))
    total_amount = parse_vn_number(_find_child_text(t_toan, "TgTTTBSo"))

    # Thuế suất từ THTTLTSuat/LTSuat
    vat_rates: List[str] = []
    if t_toan is not None:
        for elem in t_toan.iter():
            if _strip_ns(elem.tag) == "TSuat" and elem.text and elem.text.strip():
                rate = normalize_vat_rate(elem.text.strip())
                if rate and rate not in vat_rates:
                    vat_rates.append(rate)

    # Chi tiết hàng hóa DSHHDVu/HHDVu
    items: List[InvoiceItem] = []
    for elem in root.iter():
        if _strip_ns(elem.tag) == "HHDVu":
            name = _find_child_text(elem, "THHDVu")
            unit = _find_child_text(elem, "DVTinh")
            qty = parse_vn_number(_find_child_text(elem, "SLuong"))
            price = parse_vn_number(_find_child_text(elem, "DGia"))
            amount = parse_vn_number(_find_child_text(elem, "ThTien"))
            item_rate = normalize_vat_rate(_find_child_text(elem, "TSuat"))
            if item_rate and item_rate not in vat_rates:
                vat_rates.append(item_rate)
            items.append(
                InvoiceItem(
                    name=name,
                    unit=unit,
                    quantity=qty,
                    unit_price=price,
                    amount=amount,
                )
            )

    if len(vat_rates) == 1:
        vat_rate = vat_rates[0]
    elif len(vat_rates) > 1:
        vat_rate = ", ".join(vat_rates)
    else:
        vat_rate = ""

    # Nếu XML thiếu TgTCThue nhưng có items
    if sub_total == 0.0 and items:
        sub_total = sum(i.amount for i in items)

    return normalize_invoice_data(
        InvoiceData(
            invoice_symbol=invoice_symbol,
            invoice_number=invoice_number,
            invoice_date=invoice_date,
            seller_name=seller_name,
            seller_tax_code=seller_tax_code,
            buyer_name=buyer_name,
            buyer_tax_code=buyer_tax_code,
            sub_total=sub_total,
            vat_rate=vat_rate,
            vat_amount=vat_amount,
            total_amount=total_amount,
            items=items,
        )
    )


# ---------------------------------------------------------------------------
# PDF & Image Pre-processing
# ---------------------------------------------------------------------------

def inspect_pdf(
    pdf_bytes: bytes,
    min_text_chars: int = 80,
    dpi: int = 250,
    max_pages: int = 5,
) -> Tuple[str, str, List[bytes]]:
    """Phân loại PDF là dạng text gốc hay scan ảnh.

    Returns:
        (mode, extracted_text, rendered_png_pages)
        mode: "pdf_text" hoặc "pdf_scan"
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        page_count = min(len(doc), max_pages)
        text_chunks: List[str] = []
        for idx in range(page_count):
            page_text = doc[idx].get_text("text") or ""
            text_chunks.append(page_text.strip())

        full_text = "\n\n--- TRANG TIẾP THEO ---\n\n".join(t for t in text_chunks if t)
        alnum_count = len(re.findall(r"[0-9A-Za-zÀ-ỹ]", full_text))

        if alnum_count >= min_text_chars:
            return "pdf_text", full_text, []

        # PDF scan: chuyển đổi từng trang sang ảnh PNG độ phân giải cao
        zoom = dpi / 72.0
        matrix = fitz.Matrix(zoom, zoom)
        images: List[bytes] = []
        for idx in range(page_count):
            pix = doc[idx].get_pixmap(matrix=matrix, alpha=False)
            images.append(pix.tobytes("png"))
        return "pdf_scan", "", images
    finally:
        doc.close()


def preprocess_image(image_bytes: bytes) -> bytes:
    """Chuẩn hóa ảnh đầu vào (xoay đúng chiều EXIF, chuyển RGB, xuất PNG)."""
    with Image.open(io.BytesIO(image_bytes)) as img:
        img = ImageOps.exif_transpose(img)
        if img.mode != "RGB":
            img = img.convert("RGB")
        # Giới hạn cạnh lớn nhất ở 2400px để tối ưu tốc độ & chất lượng OCR
        max_dim = 2400
        if max(img.size) > max_dim:
            img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()


# ---------------------------------------------------------------------------
# Regex + Tesseract Fallback Extractor (Offline / No API Key)
# ---------------------------------------------------------------------------

def _ocr_images_with_tesseract(image_bytes_list: List[bytes], tesseract_cmd: str = "") -> str:
    """Chạy Tesseract OCR lên danh sách ảnh PNG (fallback offline)."""
    import pytesseract

    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    texts: List[str] = []
    for img_bytes in image_bytes_list:
        with Image.open(io.BytesIO(img_bytes)) as img:
            try:
                txt = pytesseract.image_to_string(img, lang="vie+eng")
            except pytesseract.TesseractError:
                txt = pytesseract.image_to_string(img)
            texts.append(txt)
    return "\n".join(texts)


def parse_text_with_regex_fallback(text: str) -> InvoiceData:
    """Bộ trích xuất dự phòng bằng Regex khi không có Gemini API Key."""
    symbol_match = re.search(
        r"(?:Ký hiệu|Serial|Ký hiệu HĐ)\s*[:\-]?\s*([0-9A-Z]{5,10})",
        text,
        flags=re.IGNORECASE,
    )
    number_match = re.search(
        r"(?:Số|No\.?|Số hóa đơn)\s*[:\-]?\s*(\d{1,8})",
        text,
        flags=re.IGNORECASE,
    )
    date_match = re.search(
        r"(Ngày\s*\d{1,2}\s*tháng\s*\d{1,2}\s*năm\s*\d{4}|\d{1,2}/\d{1,2}/\d{4})",
        text,
        flags=re.IGNORECASE,
    )
    mst_matches = re.findall(
        r"(?:Mã số thuế|MST|Tax\s*code)\s*[:\-]?\s*(\d{10}(?:-\d{3})?)",
        text,
        flags=re.IGNORECASE,
    )
    seller_match = re.search(
        r"(?:Đơn vị bán hàng|Người bán|Seller)\s*[:\-]?\s*(.+)",
        text,
        flags=re.IGNORECASE,
    )
    buyer_match = re.search(
        r"(?:Tên đơn vị|Đơn vị mua hàng|Họ tên người mua hàng|Buyer)\s*[:\-]?\s*(.+)",
        text,
        flags=re.IGNORECASE,
    )
    sub_total_match = re.search(
        r"(?:Cộng tiền hàng|Tiền hàng chưa thuế|Sub\s*total)\s*[:\-]?\s*([\d.,]+)",
        text,
        flags=re.IGNORECASE,
    )
    vat_rate_match = re.search(
        r"(?:Thuế suất GTGT|VAT\s*rate)\s*[:\-]?\s*(\d{1,2}\s*%|KCT)",
        text,
        flags=re.IGNORECASE,
    )
    vat_amount_match = re.search(
        r"(?:Tiền thuế GTGT|VAT\s*amount)\s*[:\-]?\s*([\d.,]+)",
        text,
        flags=re.IGNORECASE,
    )
    total_match = re.search(
        r"(?:Tổng cộng tiền thanh toán|Tổng tiền thanh toán|Total\s*payment)\s*[:\-]?\s*([\d.,]+)",
        text,
        flags=re.IGNORECASE,
    )

    return normalize_invoice_data(
        InvoiceData(
            invoice_symbol=symbol_match.group(1) if symbol_match else "",
            invoice_number=number_match.group(1) if number_match else "",
            invoice_date=date_match.group(1) if date_match else "",
            seller_name=seller_match.group(1).strip() if seller_match else "",
            seller_tax_code=mst_matches[0] if len(mst_matches) >= 1 else "",
            buyer_name=buyer_match.group(1).strip() if buyer_match else "",
            buyer_tax_code=mst_matches[1] if len(mst_matches) >= 2 else "",
            sub_total=parse_vn_number(sub_total_match.group(1)) if sub_total_match else 0.0,
            vat_rate=vat_rate_match.group(1) if vat_rate_match else "",
            vat_amount=parse_vn_number(vat_amount_match.group(1)) if vat_amount_match else 0.0,
            total_amount=parse_vn_number(total_match.group(1)) if total_match else 0.0,
            items=[],
        )
    )


# ---------------------------------------------------------------------------
# Excel (.xlsx / .xlsm) & CSV Parser (Kiểu 1: Bảng kê nhiều HĐ | Kiểu 2: HĐ đơn lẻ)
# ---------------------------------------------------------------------------

def _strip_accents(text: str) -> str:
    """Bỏ dấu tiếng Việt và chuẩn hóa chữ thường để so khớp tên cột linh hoạt."""
    if not text:
        return ""
    s = str(text).replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", s.lower()).strip()


# Bản đồ từ khóa nhận diện cột bảng kê hóa đơn (hỗ trợ Tổng cục Thuế, MISA, FAST & VN-InvoiceParser)
_TABLE_COLUMN_ALIASES: Dict[str, List[str]] = {
    "invoice_symbol": ["ky hieu", "invoice_symbol", "kh hd", "khhdon", "serial"],
    "invoice_number": ["so hoa don", "so hd", "invoice_number", "shdon", "invoice no"],
    "invoice_date": ["ngay lap", "ngay hoa don", "ngay hd", "invoice_date", "nlap", "date"],
    "seller_name": [
        "ten don vi ban",
        "ten nguoi ban",
        "don vi ban",
        "nguoi ban",
        "seller_name",
        "nha cung cap",
    ],
    "seller_tax_code": [
        "mst nguoi ban",
        "mst ben ban",
        "mst ban",
        "ma so thue nguoi ban",
        "seller_tax_code",
    ],
    "buyer_name": [
        "ten don vi mua",
        "ten nguoi mua",
        "don vi mua",
        "nguoi mua",
        "buyer_name",
        "khach hang",
    ],
    "buyer_tax_code": [
        "mst nguoi mua",
        "mst ben mua",
        "mst mua",
        "ma so thue nguoi mua",
        "buyer_tax_code",
    ],
    "sub_total": [
        "tien hang chua thue",
        "cong tien hang",
        "doanh so chua thue",
        "gia tri chua thue",
        "tong tien chua thue",
        "sub_total",
    ],
    "vat_rate": ["thue suat", "vat_rate", "tsuat"],
    "vat_amount": [
        "tien thue gtgt",
        "tien thue",
        "thue gtgt",
        "tong tien thue",
        "vat_amount",
    ],
    "total_amount": [
        "tong tien thanh toan",
        "tong cong tien thanh toan",
        "tong thanh toan",
        "tong cong",
        "total_amount",
    ],
}

_ITEM_COLUMN_ALIASES: Dict[str, List[str]] = {
    "invoice_symbol": ["ky hieu", "invoice_symbol"],
    "invoice_number": ["so hoa don", "so hd", "invoice_number"],
    "name": ["ten hang", "hang hoa", "dich vu", "name", "product"],
    "unit": ["dvt", "don vi tinh", "unit"],
    "quantity": ["so luong", "sl", "quantity", "qty"],
    "unit_price": ["don gia", "unit_price", "price"],
    "amount": ["thanh tien", "amount"],
}


def _match_header_row(
    rows: List[List[object]],
    aliases: Dict[str, List[str]],
    min_matches: int = 3,
    max_scan_rows: int = 15,
) -> Tuple[Optional[int], Dict[str, int]]:
    """Tìm dòng tiêu đề khớp nhiều cột hóa đơn nhất trong `max_scan_rows` dòng đầu."""
    best_row_idx: Optional[int] = None
    best_mapping: Dict[str, int] = {}

    for r_idx, row in enumerate(rows[:max_scan_rows]):
        mapping: Dict[str, int] = {}
        for c_idx, cell_val in enumerate(row):
            norm = _strip_accents(str(cell_val or ""))
            if not norm:
                continue
            for field, keywords in aliases.items():
                if field in mapping:
                    continue
                if any(kw == norm or kw in norm for kw in keywords):
                    mapping[field] = c_idx
                    break
        if len(mapping) > len(best_mapping):
            best_mapping = mapping
            best_row_idx = r_idx

    if len(best_mapping) >= min_matches:
        return best_row_idx, best_mapping
    return None, {}


def _cell_to_str(val: object) -> str:
    """Chuyển giá trị ô Excel (kể cả datetime/date/int) về chuỗi sạch."""
    if val is None:
        return ""
    if isinstance(val, (datetime, date)):
        return val.strftime("%d/%m/%Y")
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).strip()


def _is_summary_or_empty_row(row: List[object]) -> bool:
    """Kiểm tra dòng trống hoặc dòng TỔNG CỘNG cuối bảng kê."""
    non_empty = [str(v).strip() for v in row if v is not None and str(v).strip() != ""]
    if not non_empty:
        return True
    first_norm = _strip_accents(non_empty[0])
    if first_norm in {"tong cong", "tong", "total", "grand total", "cong"}:
        return True
    if first_norm.startswith("tong cong"):
        return True
    return False


def _extract_items_from_second_sheet(wb) -> Dict[Tuple[str, str], List[InvoiceItem]]:
    """Đọc sheet 'Chi tiết hàng hóa' (nếu có) để gắn lại vào các hóa đơn tương ứng."""
    items_map: Dict[Tuple[str, str], List[InvoiceItem]] = {}
    if len(wb.sheetnames) < 2:
        return items_map

    for sheet_name in wb.sheetnames[1:]:
        ws = wb[sheet_name]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        header_idx, col_map = _match_header_row(rows, _ITEM_COLUMN_ALIASES, min_matches=4)
        if header_idx is None or "name" not in col_map:
            continue

        for row in rows[header_idx + 1 :]:
            if _is_summary_or_empty_row(row):
                continue
            sym = _cell_to_str(row[col_map["invoice_symbol"]]).upper() if "invoice_symbol" in col_map else ""
            num = _cell_to_str(row[col_map["invoice_number"]]) if "invoice_number" in col_map else ""
            item_name = _cell_to_str(row[col_map["name"]])
            if not item_name:
                continue
            item = InvoiceItem(
                name=item_name,
                unit=_cell_to_str(row[col_map["unit"]]) if "unit" in col_map else "",
                quantity=parse_vn_number(row[col_map["quantity"]]) if "quantity" in col_map else 0.0,
                unit_price=parse_vn_number(row[col_map["unit_price"]]) if "unit_price" in col_map else 0.0,
                amount=parse_vn_number(row[col_map["amount"]]) if "amount" in col_map else 0.0,
            )
            key = (re.sub(r"\s+", "", sym), num.lstrip("0") or num)
            items_map.setdefault(key, []).append(item)
        break
    return items_map


def _rows_to_plain_text(rows: List[List[object]]) -> str:
    """Chuyển toàn bộ các ô không trống của sheet thành văn bản có cấu trúc cho AI/Regex."""
    lines: List[str] = []
    for row in rows:
        parts = [_cell_to_str(v) for v in row if v is not None and _cell_to_str(v)]
        if not parts:
            continue
        # Nối mượt cặp "Nhãn:" + "Giá trị" bằng dấu cách, các cột khác cách nhau bằng " | "
        merged: List[str] = []
        i = 0
        while i < len(parts):
            if parts[i].endswith(":") and i + 1 < len(parts):
                merged.append(f"{parts[i]} {parts[i + 1]}")
                i += 2
            else:
                merged.append(parts[i])
                i += 1
        lines.append(" | ".join(merged))
    return "\n".join(lines)


def _parse_tabular_rows(
    rows: List[List[object]],
    header_idx: int,
    col_map: Dict[str, int],
    filename: str,
    method: str,
    items_map: Optional[Dict[Tuple[str, str], List[InvoiceItem]]] = None,
) -> List[InvoiceRecord]:
    """Chuyển từng dòng trong bảng kê Excel/CSV thành danh sách InvoiceRecord."""
    records: List[InvoiceRecord] = []
    items_map = items_map or {}

    def _get(row: List[object], field: str) -> object:
        idx = col_map.get(field)
        if idx is None or idx >= len(row):
            return None
        return row[idx]

    for offset, row in enumerate(rows[header_idx + 1 :], start=header_idx + 2):
        if _is_summary_or_empty_row(row):
            continue

        symbol = _cell_to_str(_get(row, "invoice_symbol"))
        number = _cell_to_str(_get(row, "invoice_number"))
        date_str = _cell_to_str(_get(row, "invoice_date"))
        seller_name = _cell_to_str(_get(row, "seller_name"))
        seller_mst = _cell_to_str(_get(row, "seller_tax_code"))
        buyer_name = _cell_to_str(_get(row, "buyer_name"))
        buyer_mst = _cell_to_str(_get(row, "buyer_tax_code"))
        sub_total = parse_vn_number(_get(row, "sub_total"))
        vat_rate = _cell_to_str(_get(row, "vat_rate"))
        vat_amount = parse_vn_number(_get(row, "vat_amount"))
        total_amount = parse_vn_number(_get(row, "total_amount"))

        # Bỏ qua dòng rác không có cả số HĐ lẫn số tiền
        if not (symbol or number or seller_mst or total_amount > 0 or sub_total > 0):
            continue

        key = (re.sub(r"\s+", "", symbol.upper()), number.lstrip("0") or number)
        data = normalize_invoice_data(
            InvoiceData(
                invoice_symbol=symbol,
                invoice_number=number,
                invoice_date=date_str,
                seller_name=seller_name,
                seller_tax_code=seller_mst,
                buyer_name=buyer_name,
                buyer_tax_code=buyer_mst,
                sub_total=sub_total,
                vat_rate=vat_rate,
                vat_amount=vat_amount,
                total_amount=total_amount,
                items=items_map.get(key, []),
            )
        )
        records.append(
            InvoiceRecord(
                record_id=uuid.uuid4().hex[:10],
                source_file=f"{filename} [Dòng {offset}]",
                extraction_method=method,
                data=data,
            )
        )
    return records


# ---------------------------------------------------------------------------
# Main Extractor Class (Gemini 2.5 Flash Structured Outputs)
# ---------------------------------------------------------------------------

class InvoiceExtractor:
    """Bộ trích xuất hóa đơn thông minh hỗ trợ XML, PDF text, PDF scan và ảnh."""

    def __init__(
        self,
        api_key: Optional[str] = None,
        model_name: Optional[str] = None,
        settings: Optional[Settings] = None,
        use_fallback_only: bool = False,
    ) -> None:
        self.settings = settings or get_settings()
        self.api_key = (api_key if api_key is not None else self.settings.gemini_api_key).strip()
        self.model_name = (model_name or self.settings.gemini_model).strip()
        self.use_fallback_only = use_fallback_only
        self._client = None

        if self.api_key and not self.use_fallback_only:
            from google import genai

            self._client = genai.Client(api_key=self.api_key)

    def _call_gemini_structured(
        self,
        contents: list,
        max_retries: int = 3,
    ) -> InvoiceData:
        """Gọi Gemini API với Structured Outputs (JSON Schema = InvoiceData)."""
        if self._client is None:
            raise RuntimeError("Chưa cấu hình GEMINI_API_KEY để gọi Gemini API.")

        from google.genai import types

        last_err: Optional[Exception] = None
        for attempt in range(1, max_retries + 1):
            try:
                response = self._client.models.generate_content(
                    model=self.model_name,
                    contents=contents,
                    config=types.GenerateContentConfig(
                        system_instruction=EXTRACTION_SYSTEM_PROMPT,
                        response_mime_type="application/json",
                        response_schema=InvoiceData,
                        temperature=0.0,
                    ),
                )

                # SDK mới có response.parsed khi truyền Pydantic response_schema
                if getattr(response, "parsed", None) is not None:
                    parsed = response.parsed
                    if isinstance(parsed, InvoiceData):
                        return normalize_invoice_data(parsed)
                    return normalize_invoice_data(InvoiceData.model_validate(parsed))

                raw_text = (response.text or "").strip()
                if not raw_text:
                    raise ValueError("Gemini trả về phản hồi rỗng.")
                payload = json.loads(raw_text)
                return normalize_invoice_data(InvoiceData.model_validate(payload))

            except Exception as exc:
                last_err = exc
                logger.warning("Gemini API attempt %d/%d failed: %s", attempt, max_retries, exc)
                if attempt < max_retries:
                    time.sleep(1.5 * attempt)

        raise RuntimeError(f"Gọi Gemini API thất bại sau {max_retries} lần thử: {last_err}")

    def extract_batch_from_bytes(self, file_bytes: bytes, filename: str) -> List[InvoiceRecord]:
        """Trích xuất 1 hoặc nhiều hóa đơn từ 1 file (hỗ trợ cả bảng kê Excel/CSV nhiều dòng)."""
        ext = Path(filename).suffix.lower()

        if ext in {".xlsx", ".xlsm", ".csv"}:
            try:
                if ext == ".csv":
                    text_decoded = file_bytes.decode("utf-8-sig", errors="replace")
                    reader = csv.reader(io.StringIO(text_decoded))
                    rows: List[List[object]] = [list(r) for r in reader]
                    items_map: Dict[Tuple[str, str], List[InvoiceItem]] = {}
                else:
                    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
                    ws = wb.active
                    rows = [list(r) for r in ws.iter_rows(values_only=True)]
                    items_map = _extract_items_from_second_sheet(wb)

                # Kiểu 1: Tự động nhận diện Bảng kê nhiều hóa đơn
                header_idx, col_map = _match_header_row(rows, _TABLE_COLUMN_ALIASES, min_matches=3)
                if header_idx is not None:
                    method = "csv_table" if ext == ".csv" else "excel_table"
                    table_records = _parse_tabular_rows(
                        rows, header_idx, col_map, filename, method=method, items_map=items_map
                    )
                    if table_records:
                        return table_records

                # Kiểu 2: File Excel/CSV biểu mẫu 1 hóa đơn đơn lẻ
                plain_text = _rows_to_plain_text(rows)
                if self._client is not None:
                    prompt = (
                        "Trích xuất thông tin hóa đơn điện tử Việt Nam từ nội dung bảng tính Excel sau:\n\n"
                        f"{plain_text}"
                    )
                    data = self._call_gemini_structured([prompt])
                    method = "excel_single_gemini"
                else:
                    data = parse_text_with_regex_fallback(plain_text)
                    method = "excel_single_regex_fallback"

                return [
                    InvoiceRecord(
                        record_id=uuid.uuid4().hex[:10],
                        source_file=filename,
                        extraction_method=method,
                        data=data,
                    )
                ]
            except Exception as exc:
                logger.exception("Lỗi đọc file bảng tính %s", filename)
                return [
                    InvoiceRecord(
                        record_id=uuid.uuid4().hex[:10],
                        source_file=filename,
                        extraction_method="failed",
                        status=InvoiceStatus.ERROR,
                        raw_error=str(exc),
                        validation_errors=[f"Lỗi đọc file Excel/CSV: {exc}"],
                    )
                ]

        return [self.extract_bytes(file_bytes, filename)]

    def extract_bytes(self, file_bytes: bytes, filename: str) -> InvoiceRecord:
        """Trích xuất một hóa đơn từ dữ liệu bytes và tên file."""
        record_id = uuid.uuid4().hex[:10]
        ext = Path(filename).suffix.lower()

        if ext not in SUPPORTED_EXTENSIONS:
            return InvoiceRecord(
                record_id=record_id,
                source_file=filename,
                extraction_method="unsupported",
                status=InvoiceStatus.ERROR,
                raw_error=(
                    f"Định dạng file không được hỗ trợ: '{ext}'. "
                    "Chỉ hỗ trợ PDF, PNG, JPG, JPEG, XML, XLSX, CSV."
                ),
                validation_errors=[f"Định dạng không hỗ trợ ({ext})"],
            )

        if ext in {".xlsx", ".xlsm", ".csv"}:
            return self.extract_batch_from_bytes(file_bytes, filename)[0]

        try:
            # 1. XML hóa đơn điện tử -> parse trực tiếp, chính xác 100%, không cần API
            if ext == ".xml":
                data = parse_vn_invoice_xml(file_bytes)
                return InvoiceRecord(
                    record_id=record_id,
                    source_file=filename,
                    extraction_method="xml_native",
                    data=data,
                )

            # 2. PDF -> phân loại PDF text gốc vs PDF scan
            if ext == ".pdf":
                mode, extracted_text, page_images = inspect_pdf(
                    file_bytes,
                    min_text_chars=self.settings.min_pdf_text_chars,
                    dpi=self.settings.pdf_render_dpi,
                    max_pages=self.settings.max_pdf_pages,
                )

                if mode == "pdf_text":
                    if self._client is not None:
                        prompt = (
                            "Trích xuất thông tin hóa đơn điện tử Việt Nam từ nội dung text PDF sau:\n\n"
                            f"{extracted_text}"
                        )
                        data = self._call_gemini_structured([prompt])
                        method = "pdf_text_gemini"
                    else:
                        data = parse_text_with_regex_fallback(extracted_text)
                        method = "pdf_text_regex_fallback"
                else:
                    if self._client is not None:
                        from google.genai import types

                        parts: list = [
                            "Đây là các trang ảnh độ phân giải cao của một hóa đơn điện tử Việt Nam. "
                            "Hãy đọc kỹ và trích xuất đầy đủ thông tin theo JSON Schema."
                        ]
                        for img_bytes in page_images:
                            parts.append(types.Part.from_bytes(data=img_bytes, mime_type="image/png"))
                        data = self._call_gemini_structured(parts)
                        method = "pdf_scan_gemini"
                    else:
                        ocr_text = _ocr_images_with_tesseract(page_images, self.settings.tesseract_cmd)
                        data = parse_text_with_regex_fallback(ocr_text)
                        method = "pdf_scan_tesseract_fallback"

                return InvoiceRecord(
                    record_id=record_id,
                    source_file=filename,
                    extraction_method=method,
                    data=data,
                )

            # 3. Hình ảnh (PNG, JPG, JPEG)
            png_bytes = preprocess_image(file_bytes)
            if self._client is not None:
                from google.genai import types

                parts = [
                    "Đây là ảnh chụp / ảnh scan hóa đơn điện tử Việt Nam. "
                    "Hãy trích xuất chính xác toàn bộ thông tin theo JSON Schema.",
                    types.Part.from_bytes(data=png_bytes, mime_type="image/png"),
                ]
                data = self._call_gemini_structured(parts)
                method = "image_gemini"
            else:
                ocr_text = _ocr_images_with_tesseract([png_bytes], self.settings.tesseract_cmd)
                data = parse_text_with_regex_fallback(ocr_text)
                method = "image_tesseract_fallback"

            return InvoiceRecord(
                record_id=record_id,
                source_file=filename,
                extraction_method=method,
                data=data,
            )

        except Exception as exc:
            logger.exception("Lỗi trích xuất file %s", filename)
            return InvoiceRecord(
                record_id=record_id,
                source_file=filename,
                extraction_method="failed",
                status=InvoiceStatus.ERROR,
                raw_error=str(exc),
                validation_errors=[f"Lỗi đọc/trích xuất file: {exc}"],
            )

    def extract_file(self, file_path: str | Path) -> InvoiceRecord:
        """Trích xuất một hóa đơn từ đường dẫn tệp trên ổ đĩa."""
        path = Path(file_path)
        return self.extract_bytes(path.read_bytes(), path.name)
