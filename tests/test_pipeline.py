"""Bộ kiểm thử tự động cho XML parser, Validator và Excel Exporter."""

from __future__ import annotations

import io
from openpyxl import load_workbook

from src.exporter import export_to_excel_bytes
from src.extractor import InvoiceExtractor, normalize_date, normalize_tax_code, parse_vn_number
from src.schemas import InvoiceData, InvoiceItem, InvoiceRecord, InvoiceStatus
from src.validator import is_valid_tax_code, validate_batch


SAMPLE_VN_XML = """<?xml version="1.0" encoding="UTF-8"?>
<HDon>
  <DLHDon>
    <TTChung>
      <KHMSHDon>1</KHMSHDon>
      <KHHDon>C26TAA</KHHDon>
      <SHDon>00001234</SHDon>
      <NLap>2026-03-15</NLap>
    </TTChung>
    <NDHDon>
      <NBan>
        <Ten>CÔNG TY CỔ PHẦN CÔNG NGHỆ ABC</Ten>
        <MST>0101234567</MST>
      </NBan>
      <NMua>
        <Ten>CÔNG TY TNHH THƯƠNG MẠI XYZ</Ten>
        <MST>0309876543-001</MST>
      </NMua>
      <DSHHDVu>
        <HHDVu>
          <THHDVu>Phần mềm kế toán bản quyền</THHDVu>
          <DVTinh>Gói</DVTinh>
          <SLuong>1</SLuong>
          <DGia>10000000</DGia>
          <ThTien>10000000</ThTien>
          <TSuat>8%</TSuat>
        </HHDVu>
      </DSHHDVu>
      <TToan>
        <TgTCThue>10000000</TgTCThue>
        <TgTThue>800000</TgTThue>
        <TgTTTBSo>10800000</TgTTTBSo>
      </TToan>
    </NDHDon>
  </DLHDon>
</HDon>
""".encode("utf-8")


def test_normalization_helpers():
    assert normalize_date("2026-03-05") == "05/03/2026"
    assert normalize_date("Ngày 5 tháng 3 năm 2026") == "05/03/2026"
    assert normalize_tax_code("0101234567001") == "0101234567-001"
    assert parse_vn_number("1.500.000") == 1500000.0
    assert parse_vn_number("1.500,50") == 1500.5
    assert is_valid_tax_code("0101234567") is True
    assert is_valid_tax_code("0101234567-001") is True
    assert is_valid_tax_code("01012345") is False


def test_xml_extraction_and_validation():
    extractor = InvoiceExtractor(api_key="", use_fallback_only=True)
    record = extractor.extract_bytes(SAMPLE_VN_XML, "hd_sample.xml")

    assert record.extraction_method == "xml_native"
    assert record.data.invoice_symbol == "1C26TAA"
    assert record.data.invoice_number == "00001234"
    assert record.data.invoice_date == "15/03/2026"
    assert record.data.seller_tax_code == "0101234567"
    assert record.data.buyer_tax_code == "0309876543-001"
    assert record.data.sub_total == 10_000_000.0
    assert record.data.vat_rate == "8%"
    assert record.data.vat_amount == 800_000.0
    assert record.data.total_amount == 10_800_000.0
    assert len(record.data.items) == 1

    validated = validate_batch([record])
    assert validated[0].status == InvoiceStatus.VALID
    assert validated[0].validation_errors == []


def test_validation_math_error_and_duplicate():
    r1 = InvoiceRecord(
        record_id="r1",
        source_file="hd1.pdf",
        data=InvoiceData(
            invoice_symbol="1C26TAA",
            invoice_number="00001234",
            invoice_date="15/03/2026",
            seller_name="Công ty A",
            seller_tax_code="0101234567",
            sub_total=1_000_000,
            vat_rate="10%",
            vat_amount=100_000,
            total_amount=1_100_000,
        ),
    )
    r2 = InvoiceRecord(
        record_id="r2",
        source_file="hd1_copy.pdf",
        data=InvoiceData(
            invoice_symbol="1C26TAA",
            invoice_number="1234",  # Trùng lặp với 00001234
            invoice_date="15/03/2026",
            seller_name="Công ty A",
            seller_tax_code="0101234567",
            sub_total=1_000_000,
            vat_rate="10%",
            vat_amount=100_000,
            total_amount=1_150_000,  # Lệch 50.000đ
        ),
    )

    validated = validate_batch([r1, r2])
    assert validated[0].status == InvoiceStatus.NEEDS_REVIEW
    assert validated[1].status == InvoiceStatus.NEEDS_REVIEW
    assert any("Trùng lặp" in err for err in validated[0].validation_errors)
    assert any("Lệch tiền thanh toán" in err for err in validated[1].validation_errors)


def test_excel_export_two_sheets_and_formulas():
    rec = InvoiceRecord(
        record_id="r1",
        source_file="hd_test.xml",
        data=InvoiceData(
            invoice_symbol="1C26TAA",
            invoice_number="00000001",
            invoice_date="20/03/2026",
            seller_name="Công ty A",
            seller_tax_code="0101234567",
            buyer_name="Công ty B",
            buyer_tax_code="0301234567",
            sub_total=2_000_000,
            vat_rate="10%",
            vat_amount=200_000,
            total_amount=2_200_000,
            items=[
                InvoiceItem(
                    name="Dịch vụ tư vấn",
                    unit="Gói",
                    quantity=1,
                    unit_price=2_000_000,
                    amount=2_000_000,
                )
            ],
        ),
    )
    validated = validate_batch([rec])
    excel_bytes = export_to_excel_bytes(validated)

    wb = load_workbook(io.BytesIO(excel_bytes))
    assert wb.sheetnames == ["Bảng kê hóa đơn tổng hợp", "Chi tiết hàng hóa"]

    ws_sum = wb["Bảng kê hóa đơn tổng hợp"]
    # Row 5 là dữ liệu, Row 6 là dòng TỔNG CỘNG
    assert ws_sum["A6"].value == "TỔNG CỘNG"
    assert ws_sum["J6"].value == "=SUM(J5:J5)"
    assert ws_sum["L6"].value == "=SUM(L5:L5)"
    assert ws_sum["M6"].value == "=SUM(M5:M5)"

    ws_items = wb["Chi tiết hàng hóa"]
    # Row 4 là dòng hàng đầu tiên, Row 5 là TỔNG CỘNG
    assert ws_items["A5"].value == "TỔNG CỘNG"
    assert ws_items["K5"].value == "=SUM(K4:K4)"


def test_excel_type1_multi_invoice_table_and_roundtrip():
    """Kiểm tra đọc ngược file Excel 2 sheet (Bảng kê nhiều hóa đơn + Chi tiết hàng hóa)."""
    rec1 = InvoiceRecord(
        record_id="r1",
        source_file="hd1.xml",
        data=InvoiceData(
            invoice_symbol="1C26TAA",
            invoice_number="00000001",
            invoice_date="20/03/2026",
            seller_name="Công ty A",
            seller_tax_code="0101234567",
            buyer_name="Công ty B",
            buyer_tax_code="0301234567",
            sub_total=2_000_000,
            vat_rate="10%",
            vat_amount=200_000,
            total_amount=2_200_000,
            items=[
                InvoiceItem(name="Dịch vụ A", unit="Gói", quantity=1, unit_price=2_000_000, amount=2_000_000)
            ],
        ),
    )
    rec2 = InvoiceRecord(
        record_id="r2",
        source_file="hd2.xml",
        data=InvoiceData(
            invoice_symbol="1C26TBB",
            invoice_number="00000002",
            invoice_date="21/03/2026",
            seller_name="Công ty C",
            seller_tax_code="0107654321",
            buyer_name="Công ty B",
            buyer_tax_code="0301234567",
            sub_total=1_000_000,
            vat_rate="8%",
            vat_amount=80_000,
            total_amount=1_050_000,  # Lệch 30.000đ
        ),
    )
    excel_bytes = export_to_excel_bytes(validate_batch([rec1, rec2]))

    extractor = InvoiceExtractor(api_key="", use_fallback_only=True)
    reimported = validate_batch(extractor.extract_batch_from_bytes(excel_bytes, "bang_ke.xlsx"))

    assert len(reimported) == 2
    assert reimported[0].extraction_method == "excel_table"
    assert reimported[0].status == InvoiceStatus.VALID
    assert len(reimported[0].data.items) == 1
    assert reimported[0].data.items[0].name == "Dịch vụ A"
    assert reimported[1].status == InvoiceStatus.NEEDS_REVIEW
    assert any("Lệch tiền thanh toán" in e for e in reimported[1].validation_errors)


def test_excel_type2_single_invoice_template():
    """Kiểm tra đọc file Excel biểu mẫu 1 hóa đơn đơn lẻ (Kiểu 2)."""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.append(["HÓA ĐƠN GIÁ TRỊ GIA TĂNG"])
    ws.append(["Ký hiệu:", "1C26TEX"])
    ws.append(["Số hóa đơn:", "00009999"])
    ws.append(["Ngày lập:", "26/03/2026"])
    ws.append(["Đơn vị bán hàng:", "CÔNG TY TNHH PHẦN MỀM Á CHÂU"])
    ws.append(["Mã số thuế:", "0105556667"])
    ws.append(["Cộng tiền hàng:", 10000000])
    ws.append(["Thuế suất GTGT:", "10%"])
    ws.append(["Tiền thuế GTGT:", 1000000])
    ws.append(["Tổng cộng tiền thanh toán:", 11000000])

    buf = io.BytesIO()
    wb.save(buf)

    extractor = InvoiceExtractor(api_key="", use_fallback_only=True)
    records = validate_batch(extractor.extract_batch_from_bytes(buf.getvalue(), "hd_don_le.xlsx"))

    assert len(records) == 1
    assert records[0].extraction_method == "excel_single_regex_fallback"
    assert records[0].data.invoice_symbol == "1C26TEX"
    assert records[0].data.invoice_number == "00009999"
    assert records[0].data.seller_tax_code == "0105556667"
    assert records[0].data.total_amount == 11_000_000.0
    assert records[0].status == InvoiceStatus.VALID
