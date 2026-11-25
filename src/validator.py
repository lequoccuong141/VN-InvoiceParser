"""Module kiểm tra tính hợp lệ, đối chiếu toán học và phát hiện trùng lặp hóa đơn."""

from __future__ import annotations

import re
from collections import defaultdict
from datetime import datetime
from typing import Dict, List, Tuple

from src.schemas import InvoiceData, InvoiceRecord, InvoiceStatus

# Chuẩn Mã số thuế Việt Nam: 10 chữ số (doanh nghiệp độc lập) hoặc 10 chữ số + '-' + 3 chữ số (chi nhánh)
VN_TAX_CODE_PATTERN = re.compile(r"^\d{10}(-\d{3})?$")


def is_valid_tax_code(tax_code: str) -> bool:
    """Kiểm tra định dạng Mã số thuế Việt Nam (10 số hoặc 10-3 số)."""
    if not tax_code:
        return False
    return bool(VN_TAX_CODE_PATTERN.match(tax_code.strip()))


def is_valid_date_ddmmyyyy(date_str: str) -> bool:
    """Kiểm tra ngày lập hóa đơn đúng định dạng DD/MM/YYYY và là ngày hợp lệ."""
    if not date_str or not re.fullmatch(r"\d{2}/\d{2}/\d{4}", date_str.strip()):
        return False
    try:
        datetime.strptime(date_str.strip(), "%d/%m/%Y")
        return True
    except ValueError:
        return False


def _duplicate_key(data: InvoiceData) -> Tuple[str, str, str] | None:
    """Tạo khóa chuẩn hóa (invoice_symbol, invoice_number, seller_tax_code) để phát hiện trùng lặp."""
    symbol = re.sub(r"\s+", "", (data.invoice_symbol or "").upper())
    raw_num = (data.invoice_number or "").strip()
    # Chuẩn hóa số hóa đơn: "00001234" và "1234" được coi là trùng nhau
    number = raw_num.lstrip("0") or ("0" if raw_num else "")
    seller_mst = re.sub(r"\s+", "", (data.seller_tax_code or "").strip())

    if not (symbol and number and seller_mst):
        return None
    return (symbol, number, seller_mst)


def validate_single_invoice(
    record: InvoiceRecord,
    amount_tolerance: float = 1.0,
) -> InvoiceRecord:
    """Kiểm tra các quy tắc nghiệp vụ cho một hóa đơn đơn lẻ."""
    # Giữ nguyên trạng thái lỗi nếu quá trình đọc/trích xuất thất bại hoàn toàn
    if record.raw_error:
        record.status = InvoiceStatus.ERROR
        if not record.validation_errors:
            record.validation_errors = [f"Lỗi xử lý: {record.raw_error}"]
        return record

    errors: List[str] = []
    warnings: List[str] = []
    data = record.data

    # 1. Kiểm tra các trường định danh bắt buộc
    if not data.invoice_symbol.strip():
        errors.append("Thiếu ký hiệu hóa đơn")
    if not data.invoice_number.strip():
        errors.append("Thiếu số hóa đơn")
    if not data.invoice_date.strip():
        errors.append("Thiếu ngày lập hóa đơn")
    elif not is_valid_date_ddmmyyyy(data.invoice_date):
        errors.append(f"Ngày lập hóa đơn không đúng định dạng DD/MM/YYYY ('{data.invoice_date}')")

    # 2. Kiểm tra Mã số thuế người bán (bắt buộc) & người mua (nếu có)
    if not data.seller_tax_code.strip():
        errors.append("Thiếu Mã số thuế người bán")
    elif not is_valid_tax_code(data.seller_tax_code):
        errors.append(
            f"MST người bán sai định dạng ('{data.seller_tax_code}', yêu cầu 10 số hoặc 10-3 số)"
        )

    if data.buyer_tax_code.strip() and not is_valid_tax_code(data.buyer_tax_code):
        errors.append(
            f"MST người mua sai định dạng ('{data.buyer_tax_code}', yêu cầu 10 số hoặc 10-3 số)"
        )

    # 3. Kiểm tra logic toán học: sub_total + vat_amount == total_amount
    expected_total = round(data.sub_total + data.vat_amount, 2)
    diff = round(expected_total - data.total_amount, 2)
    if abs(diff) > amount_tolerance:
        errors.append(
            f"Lệch tiền thanh toán: Tiền hàng ({data.sub_total:,.0f}) + VAT ({data.vat_amount:,.0f}) "
            f"= {expected_total:,.0f} ≠ Tổng tiền ({data.total_amount:,.0f}), chênh lệch {diff:+,.0f} đ"
        )

    if data.total_amount <= 0:
        warnings.append("Tổng tiền thanh toán <= 0 (cần xác nhận hóa đơn điều chỉnh/giảm giá)")

    # 4. Đối chiếu chéo tổng thành tiền các dòng hàng hóa với sub_total (cảnh báo)
    if data.items:
        items_sum = round(sum(item.amount for item in data.items), 2)
        items_diff = round(items_sum - data.sub_total, 2)
        if abs(items_diff) > amount_tolerance:
            warnings.append(
                f"Tổng thành tiền chi tiết hàng hóa ({items_sum:,.0f}) lệch với Tiền hàng chưa thuế "
                f"({data.sub_total:,.0f}) là {items_diff:+,.0f} đ"
            )

    record.validation_errors = errors
    record.validation_warnings = warnings
    record.status = InvoiceStatus.NEEDS_REVIEW if errors else InvoiceStatus.VALID
    return record


def validate_batch(
    records: List[InvoiceRecord],
    amount_tolerance: float = 1.0,
) -> List[InvoiceRecord]:
    """Kiểm tra toàn bộ lô hóa đơn, bao gồm đối chiếu toán học, MST và phát hiện trùng lặp."""
    # Bước 1: Kiểm tra từng hóa đơn
    for rec in records:
        validate_single_invoice(rec, amount_tolerance=amount_tolerance)

    # Bước 2: Phát hiện trùng lặp theo (invoice_symbol, invoice_number, seller_tax_code)
    key_map: Dict[Tuple[str, str, str], List[InvoiceRecord]] = defaultdict(list)
    for rec in records:
        if rec.raw_error:
            continue
        key = _duplicate_key(rec.data)
        if key is not None:
            key_map[key].append(rec)

    for key, dup_records in key_map.items():
        if len(dup_records) > 1:
            symbol, number, mst = key
            file_list = ", ".join(r.source_file for r in dup_records)
            dup_msg = (
                f"Trùng lặp hóa đơn (Ký hiệu: {symbol}, Số: {number}, MST bán: {mst}) "
                f"trong lô tải lên [{file_list}]"
            )
            for rec in dup_records:
                if dup_msg not in rec.validation_errors:
                    rec.validation_errors.append(dup_msg)
                rec.status = InvoiceStatus.NEEDS_REVIEW

    return records
