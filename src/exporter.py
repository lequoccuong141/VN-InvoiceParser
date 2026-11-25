"""Module xuất báo cáo Excel kế toán 2 sheet chuẩn định dạng bằng OpenPyXL."""

from __future__ import annotations

import io
from datetime import datetime
from pathlib import Path
from typing import List

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from src.schemas import InvoiceRecord, InvoiceStatus

# Bảng màu chuẩn báo cáo kế toán
NAVY_HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
SUBHEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
TOTAL_ROW_FILL = PatternFill(start_color="E9EDF4", end_color="E9EDF4", fill_type="solid")
VALID_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
ERROR_ROW_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
ZEBRA_FILL = PatternFill(start_color="F9FAFC", end_color="F9FAFC", fill_type="solid")

FONT_FAMILY = "Arial"
TITLE_FONT = Font(name=FONT_FAMILY, size=14, bold=True, color="1F4E78")
SUBTITLE_FONT = Font(name=FONT_FAMILY, size=10, italic=True, color="595959")
HEADER_FONT = Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF")
BODY_FONT = Font(name=FONT_FAMILY, size=10, color="1A1A1A")
BOLD_FONT = Font(name=FONT_FAMILY, size=10, bold=True, color="1A1A1A")
VALID_FONT = Font(name=FONT_FAMILY, size=10, bold=True, color="276749")
ERROR_FONT = Font(name=FONT_FAMILY, size=10, bold=True, color="C00000")

THIN_SIDE = Side(style="thin", color="BFBFBF")
MEDIUM_TOP = Side(style="medium", color="1F4E78")
DOUBLE_BOTTOM = Side(style="double", color="1F4E78")

CELL_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
TOTAL_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=MEDIUM_TOP, bottom=DOUBLE_BOTTOM)

CURRENCY_FORMAT = "#,##0"
QUANTITY_FORMAT = "#,##0.##"


def _autofit_columns(ws, min_width: int = 10, max_width: int = 45, start_row: int = 4) -> None:
    """Tự động căn chỉnh độ rộng cột dựa trên nội dung thực tế."""
    for col_cells in ws.columns:
        col_idx = col_cells[0].column
        col_letter = get_column_letter(col_idx)
        max_len = min_width
        for cell in col_cells:
            if cell.row < start_row:
                continue
            val = cell.value
            if val is None:
                continue
            str_val = str(val)
            # Công thức Excel ước lượng theo độ rộng số tiền hiển thị
            if str_val.startswith("="):
                max_len = max(max_len, 18)
            else:
                longest_line = max((len(line) for line in str_val.splitlines()), default=0)
                max_len = max(max_len, longest_line + 3)
        ws.column_dimensions[col_letter].width = min(max_len, max_width)


def _build_summary_sheet(wb: Workbook, records: List[InvoiceRecord]) -> None:
    """Tạo Sheet 1: Bảng kê hóa đơn tổng hợp."""
    ws = wb.active
    ws.title = "Bảng kê hóa đơn tổng hợp"
    ws.views.sheetView[0].showGridLines = True

    headers = [
        "STT",
        "Tên file nguồn",
        "Ký hiệu HĐ",
        "Số hóa đơn",
        "Ngày lập",
        "Tên đơn vị bán hàng",
        "MST người bán",
        "Tên đơn vị mua hàng",
        "MST người mua",
        "Tiền hàng chưa thuế",
        "Thuế suất",
        "Tiền thuế GTGT",
        "Tổng tiền thanh toán",
        "Trạng thái",
        "Ghi chú đối chiếu / Cảnh báo",
    ]
    total_cols = len(headers)
    last_col_letter = get_column_letter(total_cols)

    # Tiêu đề báo cáo
    ws.merge_cells(f"A1:{last_col_letter}1")
    ws["A1"] = "BẢNG KÊ HÓA ĐƠN ĐIỆN TỬ TỔNG HỢP (VN-INVOICEPARSER)"
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    valid_count = sum(1 for r in records if r.status == InvoiceStatus.VALID)
    review_count = len(records) - valid_count
    ws.merge_cells(f"A2:{last_col_letter}2")
    ws["A2"] = (
        f"Ngày xuất: {datetime.now().strftime('%d/%m/%Y %H:%M')}  |  "
        f"Tổng số hóa đơn: {len(records)}  |  Hợp lệ: {valid_count}  |  Cần kiểm tra / Lỗi: {review_count}"
    )
    ws["A2"].font = SUBTITLE_FONT
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    # Hàng tiêu đề bảng (Row 4)
    header_row = 4
    ws.row_dimensions[header_row].height = 30
    for col_idx, title in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=title)
        cell.font = HEADER_FONT
        cell.fill = NAVY_HEADER_FILL
        cell.border = CELL_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Dữ liệu từng hóa đơn
    data_start_row = header_row + 1
    for idx, rec in enumerate(records, start=1):
        row_num = data_start_row + idx - 1
        d = rec.data
        notes_parts = []
        if rec.validation_errors:
            notes_parts.append("LỖI: " + "; ".join(rec.validation_errors))
        if rec.validation_warnings:
            notes_parts.append("Cảnh báo: " + "; ".join(rec.validation_warnings))
        notes = "\n".join(notes_parts)

        row_values = [
            idx,
            rec.source_file,
            d.invoice_symbol,
            d.invoice_number,
            d.invoice_date,
            d.seller_name,
            d.seller_tax_code,
            d.buyer_name,
            d.buyer_tax_code,
            d.sub_total,
            d.vat_rate,
            d.vat_amount,
            d.total_amount,
            rec.status.value,
            notes,
        ]

        is_problem = rec.status != InvoiceStatus.VALID
        row_fill = ERROR_ROW_FILL if is_problem else (ZEBRA_FILL if idx % 2 == 0 else None)

        for col_idx, val in enumerate(row_values, start=1):
            cell = ws.cell(row=row_num, column=col_idx, value=val)
            cell.font = BODY_FONT
            cell.border = CELL_BORDER
            if row_fill is not None:
                cell.fill = row_fill

            # Định dạng theo loại cột
            if col_idx in {1, 3, 4, 5, 7, 9, 11}:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in {10, 12, 13}:
                cell.number_format = CURRENCY_FORMAT
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_idx == 14:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if rec.status == InvoiceStatus.VALID:
                    cell.fill = VALID_FILL
                    cell.font = VALID_FONT
                else:
                    cell.fill = ERROR_ROW_FILL
                    cell.font = ERROR_FONT
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        ws.row_dimensions[row_num].height = 24 if not notes else 36

    # Dòng TỔNG CỘNG cuối bảng dùng công thức Excel SUM
    data_end_row = data_start_row + len(records) - 1
    total_row = data_end_row + 1 if records else data_start_row
    ws.row_dimensions[total_row].height = 26

    for col_idx in range(1, total_cols + 1):
        cell = ws.cell(row=total_row, column=col_idx)
        cell.fill = TOTAL_ROW_FILL
        cell.border = TOTAL_BORDER
        cell.font = BOLD_FONT

    ws.cell(row=total_row, column=1, value="TỔNG CỘNG").alignment = Alignment(
        horizontal="center", vertical="center"
    )
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=9)

    if records:
        for col_idx in (10, 12, 13):
            col_letter = get_column_letter(col_idx)
            cell = ws.cell(
                row=total_row,
                column=col_idx,
                value=f"=SUM({col_letter}{data_start_row}:{col_letter}{data_end_row})",
            )
            cell.number_format = CURRENCY_FORMAT
            cell.alignment = Alignment(horizontal="right", vertical="center")
    else:
        for col_idx in (10, 12, 13):
            cell = ws.cell(row=total_row, column=col_idx, value=0)
            cell.number_format = CURRENCY_FORMAT
            cell.alignment = Alignment(horizontal="right", vertical="center")

    ws.freeze_panes = f"A{data_start_row}"
    ws.auto_filter.ref = f"A{header_row}:{last_col_letter}{max(data_end_row, header_row)}"
    _autofit_columns(ws, start_row=header_row)
    # Bảo đảm độ rộng tối thiểu cho các cột tên & ghi chú
    ws.column_dimensions["F"].width = max(ws.column_dimensions["F"].width, 28)
    ws.column_dimensions["H"].width = max(ws.column_dimensions["H"].width, 26)
    ws.column_dimensions["O"].width = max(ws.column_dimensions["O"].width, 36)


def _build_items_sheet(wb: Workbook, records: List[InvoiceRecord]) -> None:
    """Tạo Sheet 2: Chi tiết hàng hóa."""
    ws = wb.create_sheet(title="Chi tiết hàng hóa")
    ws.views.sheetView[0].showGridLines = True

    headers = [
        "STT",
        "Tên file nguồn",
        "Ký hiệu HĐ",
        "Số hóa đơn",
        "Ngày lập",
        "MST người bán",
        "Tên hàng hóa, dịch vụ",
        "ĐVT",
        "Số lượng",
        "Đơn giá",
        "Thành tiền",
    ]
    total_cols = len(headers)
    last_col_letter = get_column_letter(total_cols)

    ws.merge_cells(f"A1:{last_col_letter}1")
    ws["A1"] = "BẢNG CHI TIẾT HÀNG HÓA, DỊCH VỤ THEO HÓA ĐƠN"
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    header_row = 3
    ws.row_dimensions[header_row].height = 28
    for col_idx, title in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=title)
        cell.font = HEADER_FONT
        cell.fill = NAVY_HEADER_FILL
        cell.border = CELL_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    data_start_row = header_row + 1
    item_idx = 0
    for rec in records:
        d = rec.data
        for item in d.items:
            item_idx += 1
            row_num = data_start_row + item_idx - 1
            row_values = [
                item_idx,
                rec.source_file,
                d.invoice_symbol,
                d.invoice_number,
                d.invoice_date,
                d.seller_tax_code,
                item.name,
                item.unit,
                item.quantity,
                item.unit_price,
                item.amount,
            ]
            row_fill = ZEBRA_FILL if item_idx % 2 == 0 else None
            for col_idx, val in enumerate(row_values, start=1):
                cell = ws.cell(row=row_num, column=col_idx, value=val)
                cell.font = BODY_FONT
                cell.border = CELL_BORDER
                if row_fill is not None:
                    cell.fill = row_fill

                if col_idx in {1, 3, 4, 5, 6, 8}:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_idx == 9:
                    cell.number_format = QUANTITY_FORMAT
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif col_idx in {10, 11}:
                    cell.number_format = CURRENCY_FORMAT
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws.row_dimensions[row_num].height = 22

    data_end_row = data_start_row + item_idx - 1
    total_row = data_end_row + 1 if item_idx > 0 else data_start_row
    ws.row_dimensions[total_row].height = 26

    for col_idx in range(1, total_cols + 1):
        cell = ws.cell(row=total_row, column=col_idx)
        cell.fill = TOTAL_ROW_FILL
        cell.border = TOTAL_BORDER
        cell.font = BOLD_FONT

    ws.cell(row=total_row, column=1, value="TỔNG CỘNG").alignment = Alignment(
        horizontal="center", vertical="center"
    )
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=8)

    if item_idx > 0:
        qty_cell = ws.cell(
            row=total_row,
            column=9,
            value=f"=SUM(I{data_start_row}:I{data_end_row})",
        )
        qty_cell.number_format = QUANTITY_FORMAT
        qty_cell.alignment = Alignment(horizontal="right", vertical="center")

        amt_cell = ws.cell(
            row=total_row,
            column=11,
            value=f"=SUM(K{data_start_row}:K{data_end_row})",
        )
        amt_cell.number_format = CURRENCY_FORMAT
        amt_cell.alignment = Alignment(horizontal="right", vertical="center")
    else:
        ws.cell(row=total_row, column=11, value=0).number_format = CURRENCY_FORMAT

    ws.freeze_panes = f"A{data_start_row}"
    ws.auto_filter.ref = f"A{header_row}:{last_col_letter}{max(data_end_row, header_row)}"
    _autofit_columns(ws, start_row=header_row)
    ws.column_dimensions["G"].width = max(ws.column_dimensions["G"].width, 36)


def export_to_excel_bytes(records: List[InvoiceRecord]) -> bytes:
    """Tạo file Excel 2 sheet chuẩn kế toán và trả về dữ liệu bytes."""
    wb = Workbook()
    _build_summary_sheet(wb, records)
    _build_items_sheet(wb, records)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def export_to_excel_file(records: List[InvoiceRecord], output_path: str | Path) -> Path:
    """Xuất file Excel 2 sheet ra ổ đĩa."""
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(export_to_excel_bytes(records))
    return path
