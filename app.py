"""VN-InvoiceParser — Giao diện Streamlit cho xử lý, đối chiếu và xuất Excel hóa đơn Việt Nam."""

from __future__ import annotations

import json
from datetime import datetime
from typing import List

import pandas as pd
import streamlit as st

from src.config import get_settings
from src.exporter import export_to_excel_bytes
from src.extractor import InvoiceExtractor, normalize_invoice_data
from src.schemas import InvoiceData, InvoiceItem, InvoiceRecord, InvoiceStatus
from src.validator import validate_batch

st.set_page_config(
    page_title="VN-InvoiceParser | Trích xuất & Đối chiếu Hóa đơn Việt Nam",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Dữ liệu mẫu (Demo) giúp người dùng trải nghiệm ngay không cần API Key
# ---------------------------------------------------------------------------

def _create_demo_records() -> List[InvoiceRecord]:
    return [
        InvoiceRecord(
            record_id="demo_001",
            source_file="HD_00001234_FPT.xml",
            extraction_method="xml_native",
            data=InvoiceData(
                invoice_symbol="1C26TAA",
                invoice_number="00001234",
                invoice_date="15/03/2026",
                seller_name="CÔNG TY CỔ PHẦN VIỄN THÔNG FPT",
                seller_tax_code="0101778163",
                buyer_name="CÔNG TY TNHH GIẢI PHÁP CÔNG NGHỆ SAO VIỆT",
                buyer_tax_code="0314567890",
                sub_total=10_000_000,
                vat_rate="8%",
                vat_amount=800_000,
                total_amount=10_800_000,
                items=[
                    InvoiceItem(
                        name="Dịch vụ Internet doanh nghiệp Fiber Business",
                        unit="Tháng",
                        quantity=2,
                        unit_price=3_500_000,
                        amount=7_000_000,
                    ),
                    InvoiceItem(
                        name="Dịch vụ IP Tĩnh doanh nghiệp",
                        unit="Tháng",
                        quantity=2,
                        unit_price=1_500_000,
                        amount=3_000_000,
                    ),
                ],
            ),
        ),
        InvoiceRecord(
            record_id="demo_002",
            source_file="HD_00005678_ScanLoiTien.pdf",
            extraction_method="pdf_scan_gemini",
            data=InvoiceData(
                invoice_symbol="1C26TBB",
                invoice_number="00005678",
                invoice_date="18/03/2026",
                seller_name="CÔNG TY TNHH VĂN PHÒNG PHẨM MINH ANH",
                seller_tax_code="0102345678-001",
                buyer_name="CÔNG TY TNHH GIẢI PHÁP CÔNG NGHỆ SAO VIỆT",
                buyer_tax_code="0314567890",
                sub_total=5_000_000,
                vat_rate="10%",
                vat_amount=500_000,
                total_amount=5_450_000,  # Cố tình lệch 50.000đ để minh họa cảnh báo đỏ
                items=[
                    InvoiceItem(
                        name="Giấy in A4 Double A 80gsm",
                        unit="Ram",
                        quantity=50,
                        unit_price=80_000,
                        amount=4_000_000,
                    ),
                    InvoiceItem(
                        name="Mực in laser HP 107A chính hãng",
                        unit="Hộp",
                        quantity=1,
                        unit_price=1_000_000,
                        amount=1_000_000,
                    ),
                ],
            ),
        ),
        InvoiceRecord(
            record_id="demo_003",
            source_file="HD_00001234_FPT_BanSao.pdf",
            extraction_method="pdf_text_gemini",
            data=InvoiceData(
                invoice_symbol="1C26TAA",
                invoice_number="00001234",  # Cố tình trùng lặp với demo_001
                invoice_date="15/03/2026",
                seller_name="CÔNG TY CỔ PHẦN VIỄN THÔNG FPT",
                seller_tax_code="0101778163",
                buyer_name="CÔNG TY TNHH GIẢI PHÁP CÔNG NGHỆ SAO VIỆT",
                buyer_tax_code="0314567890",
                sub_total=10_000_000,
                vat_rate="8%",
                vat_amount=800_000,
                total_amount=10_800_000,
                items=[
                    InvoiceItem(
                        name="Dịch vụ Internet doanh nghiệp Fiber Business",
                        unit="Tháng",
                        quantity=2,
                        unit_price=3_500_000,
                        amount=7_000_000,
                    ),
                ],
            ),
        ),
    ]


# ---------------------------------------------------------------------------
# Helpers chuyển đổi DataFrame <-> List[InvoiceRecord]
# ---------------------------------------------------------------------------

SUMMARY_COLUMNS = [
    "record_id",
    "source_file",
    "status",
    "invoice_symbol",
    "invoice_number",
    "invoice_date",
    "seller_name",
    "seller_tax_code",
    "buyer_name",
    "buyer_tax_code",
    "sub_total",
    "vat_rate",
    "vat_amount",
    "total_amount",
    "extraction_method",
    "validation_errors",
]


def records_to_dataframe(records: List[InvoiceRecord]) -> pd.DataFrame:
    rows = [r.to_summary_dict() for r in records]
    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(columns=SUMMARY_COLUMNS)
    return df[SUMMARY_COLUMNS]


def apply_dataframe_edits(
    records: List[InvoiceRecord],
    edited_df: pd.DataFrame,
    amount_tolerance: float,
) -> List[InvoiceRecord]:
    """Cập nhật dữ liệu người dùng sửa trên giao diện vào danh sách InvoiceRecord và đối chiếu lại."""
    record_map = {r.record_id: r for r in records}
    updated: List[InvoiceRecord] = []

    for _, row in edited_df.iterrows():
        rec_id = str(row["record_id"])
        rec = record_map.get(rec_id)
        if rec is None:
            continue

        rec.data.invoice_symbol = str(row.get("invoice_symbol") or "")
        rec.data.invoice_number = str(row.get("invoice_number") or "")
        rec.data.invoice_date = str(row.get("invoice_date") or "")
        rec.data.seller_name = str(row.get("seller_name") or "")
        rec.data.seller_tax_code = str(row.get("seller_tax_code") or "")
        rec.data.buyer_name = str(row.get("buyer_name") or "")
        rec.data.buyer_tax_code = str(row.get("buyer_tax_code") or "")
        rec.data.vat_rate = str(row.get("vat_rate") or "")
        rec.data.sub_total = float(row.get("sub_total") or 0.0)
        rec.data.vat_amount = float(row.get("vat_amount") or 0.0)
        rec.data.total_amount = float(row.get("total_amount") or 0.0)

        rec.data = normalize_invoice_data(rec.data)
        # Nếu người dùng tự sửa tay dữ liệu của hóa đơn từng lỗi đọc file, cho phép kiểm tra lại
        if rec.raw_error and (rec.data.invoice_number or rec.data.total_amount > 0):
            rec.raw_error = None
        updated.append(rec)

    return validate_batch(updated, amount_tolerance=amount_tolerance)


def _highlight_status_rows(row: pd.Series) -> List[str]:
    """Tô màu đỏ nhạt cho các dòng Cần kiểm tra / Lỗi xử lý."""
    status = row.get("status", "")
    if status != InvoiceStatus.VALID.value:
        return ["background-color: #FCE4D6; color: #9B1C1C; font-weight: 500;"] * len(row)
    return [""] * len(row)


# ---------------------------------------------------------------------------
# Main Streamlit Application
# ---------------------------------------------------------------------------

def _get_default_api_key(env_key: str) -> str:
    """Đọc GEMINI_API_KEY từ Streamlit Secrets (khi chạy trên Cloud) hoặc từ .env."""
    if env_key:
        return env_key
    try:
        return str(st.secrets.get("GEMINI_API_KEY", ""))
    except Exception:
        return ""


def main() -> None:
    settings = get_settings()

    if "records" not in st.session_state:
        st.session_state["records"] = []

    # Sidebar cấu hình
    with st.sidebar:
        st.title("⚙️ Cấu hình hệ thống")
        api_key = st.text_input(
            "Gemini API Key",
            value=_get_default_api_key(settings.gemini_api_key),
            type="password",
            help="Lấy miễn phí tại Google AI Studio (https://aistudio.google.com/apikey). File XML không cần API Key.",
        )
        model_name = st.selectbox(
            "Mô hình AI (Structured Outputs)",
            options=["gemini-2.5-flash", "gemini-2.5-pro"],
            index=0 if settings.gemini_model == "gemini-2.5-flash" else 1,
        )
        use_fallback_only = st.checkbox(
            "Chế độ Offline Fallback (Tesseract + Regex)",
            value=False,
            help="Bỏ qua Gemini API và chỉ dùng XML parser / PyMuPDF + Regex / Tesseract OCR.",
        )
        amount_tolerance = st.number_input(
            "Dung sai làm tròn tiền (VNĐ)",
            min_value=0.0,
            max_value=1000.0,
            value=float(settings.amount_tolerance),
            step=1.0,
            help="Cho phép sai số tối đa khi kiểm tra: Tiền hàng + VAT == Tổng tiền.",
        )

        st.divider()
        st.subheader("🧪 Dữ liệu thử nghiệm")
        col_demo, col_clear = st.columns(2)
        with col_demo:
            if st.button("Nạp mẫu Demo", use_container_width=True):
                demo = validate_batch(_create_demo_records(), amount_tolerance=amount_tolerance)
                st.session_state["records"] = demo
                st.rerun()
        with col_clear:
            if st.button("Xóa dữ liệu", use_container_width=True):
                st.session_state["records"] = []
                st.rerun()

        st.divider()
        st.caption(
            "**VN-InvoiceParser v1.0**\n\n"
            "Hỗ trợ tự động phân loại:\n"
            "- **XML**: Parse trực tiếp NĐ123/TT78\n"
            "- **PDF Text**: Đọc text gốc + Gemini\n"
            "- **PDF Scan / Ảnh**: Render DPI cao + Gemini Vision"
        )

    # Header chính
    st.title("🧾 VN-InvoiceParser")
    st.markdown(
        "Hệ thống trích xuất dữ liệu hóa đơn điện tử Việt Nam hàng loạt (**PDF, JPG, PNG, XML, Excel/CSV**), "
        "tự động đối chiếu lỗi toán học, MST, trùng lặp và xuất sổ sách Excel kế toán."
    )

    # Khu vực tải file
    uploaded_files = st.file_uploader(
        "Kéo thả nhiều hóa đơn vào đây (PDF, PNG, JPG, JPEG, XML, XLSX, CSV)",
        type=["pdf", "png", "jpg", "jpeg", "xml", "xlsx", "xlsm", "csv"],
        accept_multiple_files=True,
    )

    col_run, col_info = st.columns([1, 3])
    with col_run:
        process_btn = st.button(
            "🚀 Bắt đầu trích xuất & Đối chiếu",
            type="primary",
            disabled=not uploaded_files,
            use_container_width=True,
        )
    with col_info:
        if uploaded_files and not api_key and not use_fallback_only:
            has_ai_files = any(
                not f.name.lower().endswith((".xml", ".xlsx", ".xlsm", ".csv"))
                for f in uploaded_files
            )
            if has_ai_files:
                st.warning(
                    "Bạn chưa nhập `GEMINI_API_KEY`. Các file PDF/Ảnh sẽ tự động chạy ở chế độ "
                    "fallback (Regex/Tesseract). Hãy nhập API Key ở thanh bên để đạt độ chính xác cao nhất."
                )

    if process_btn and uploaded_files:
        extractor = InvoiceExtractor(
            api_key=api_key,
            model_name=model_name,
            settings=settings,
            use_fallback_only=use_fallback_only,
        )
        progress_bar = st.progress(0, text="Đang khởi tạo bộ trích xuất...")
        extracted_records: List[InvoiceRecord] = []
        total_files = len(uploaded_files)

        for idx, file_obj in enumerate(uploaded_files, start=1):
            progress_bar.progress(
                int((idx - 1) / total_files * 100),
                text=f"Đang xử lý ({idx}/{total_files}): {file_obj.name} ...",
            )
            file_bytes = file_obj.getvalue()
            recs = extractor.extract_batch_from_bytes(file_bytes, file_obj.name)
            extracted_records.extend(recs)

        progress_bar.progress(100, text="Đang đối chiếu toán học, MST và kiểm tra trùng lặp...")
        validated_records = validate_batch(extracted_records, amount_tolerance=amount_tolerance)
        st.session_state["records"] = validated_records
        progress_bar.empty()
        st.success(f"Hoàn tất xử lý {len(validated_records)} hóa đơn!")

    records: List[InvoiceRecord] = st.session_state.get("records", [])
    if not records:
        st.info(
            "Chưa có dữ liệu hóa đơn nào. Hãy tải lên file hóa đơn ở trên hoặc nhấn **Nạp mẫu Demo** "
            "ở thanh bên trái để xem thử giao diện đối chiếu & xuất Excel."
        )
        return

    # KPI Cards
    total_count = len(records)
    valid_count = sum(1 for r in records if r.status == InvoiceStatus.VALID)
    review_count = total_count - valid_count
    grand_total = sum(r.data.total_amount for r in records)

    st.divider()
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("📄 Tổng số hóa đơn", f"{total_count}")
    k2.metric("✅ Hợp lệ", f"{valid_count}")
    k3.metric("🚨 Cần kiểm tra / Lỗi", f"{review_count}")
    k4.metric("💰 Tổng tiền thanh toán", f"{grand_total:,.0f} đ")

    # Tabs giao diện làm việc
    tab_edit, tab_items, tab_export = st.tabs(
        ["📝 Bảng kê & Chỉnh sửa trực tiếp", "📦 Chi tiết hàng hóa", "📥 Xuất Excel & JSON"]
    )

    # ------------------------------------------------------------------
    # TAB 1: Bảng kê & Chỉnh sửa trực tiếp
    # ------------------------------------------------------------------
    with tab_edit:
        st.subheader("1. Trạng thái đối chiếu (Tự động tô đỏ hóa đơn lỗi)")
        df_summary = records_to_dataframe(records)
        styled_df = (
            df_summary.drop(columns=["record_id"])
            .style.apply(_highlight_status_rows, axis=1)
            .format(
                {
                    "sub_total": "{:,.0f}",
                    "vat_amount": "{:,.0f}",
                    "total_amount": "{:,.0f}",
                }
            )
        )
        st.dataframe(styled_df, use_container_width=True, hide_index=True)

        st.subheader("2. Chỉnh sửa dữ liệu trực tiếp")
        st.caption(
            "Nhấp đúp vào bất kỳ ô nào bên dưới để chỉnh sửa nếu AI/OCR đọc sai, sau đó nhấn "
            "**Áp dụng chỉnh sửa & Đối chiếu lại**."
        )

        edited_df = st.data_editor(
            df_summary,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            disabled=["record_id", "source_file", "status", "extraction_method", "validation_errors"],
            column_config={
                "record_id": st.column_config.TextColumn("ID", width="small"),
                "source_file": st.column_config.TextColumn("Tên file"),
                "status": st.column_config.TextColumn("Trạng thái"),
                "invoice_symbol": st.column_config.TextColumn("Ký hiệu HĐ"),
                "invoice_number": st.column_config.TextColumn("Số HĐ"),
                "invoice_date": st.column_config.TextColumn("Ngày lập (DD/MM/YYYY)"),
                "seller_name": st.column_config.TextColumn("Tên người bán", width="medium"),
                "seller_tax_code": st.column_config.TextColumn("MST người bán"),
                "buyer_name": st.column_config.TextColumn("Tên người mua", width="medium"),
                "buyer_tax_code": st.column_config.TextColumn("MST người mua"),
                "sub_total": st.column_config.NumberColumn("Tiền hàng chưa thuế", format="%.0f"),
                "vat_rate": st.column_config.TextColumn("Thuế suất"),
                "vat_amount": st.column_config.NumberColumn("Tiền thuế GTGT", format="%.0f"),
                "total_amount": st.column_config.NumberColumn("Tổng tiền thanh toán", format="%.0f"),
                "extraction_method": st.column_config.TextColumn("Phương thức"),
                "validation_errors": st.column_config.TextColumn("Chi tiết lỗi", width="large"),
            },
            key="invoice_summary_editor",
        )

        if st.button("💾 Áp dụng chỉnh sửa & Đối chiếu lại", type="primary"):
            st.session_state["records"] = apply_dataframe_edits(
                records, edited_df, amount_tolerance=amount_tolerance
            )
            st.success("Đã cập nhật dữ liệu và đối chiếu lại toàn bộ lô hóa đơn!")
            st.rerun()

    # ------------------------------------------------------------------
    # TAB 2: Chi tiết hàng hóa
    # ------------------------------------------------------------------
    with tab_items:
        st.subheader("Danh sách hàng hóa / dịch vụ trích xuất")
        invoice_labels = [
            f"{idx}. {r.source_file} ({r.data.invoice_symbol} - {r.data.invoice_number})"
            for idx, r in enumerate(records, start=1)
        ]
        selected_idx = st.selectbox(
            "Chọn hóa đơn để xem / chỉnh sửa chi tiết hàng hóa:",
            options=range(len(records)),
            format_func=lambda i: invoice_labels[i],
        )
        selected_rec = records[selected_idx]

        items_rows = [
            {
                "name": item.name,
                "unit": item.unit,
                "quantity": item.quantity,
                "unit_price": item.unit_price,
                "amount": item.amount,
            }
            for item in selected_rec.data.items
        ]
        items_df = pd.DataFrame(
            items_rows, columns=["name", "unit", "quantity", "unit_price", "amount"]
        )

        edited_items_df = st.data_editor(
            items_df,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "name": st.column_config.TextColumn("Tên hàng hóa / dịch vụ", width="large"),
                "unit": st.column_config.TextColumn("ĐVT", width="small"),
                "quantity": st.column_config.NumberColumn("Số lượng", format="%.2f"),
                "unit_price": st.column_config.NumberColumn("Đơn giá", format="%.0f"),
                "amount": st.column_config.NumberColumn("Thành tiền", format="%.0f"),
            },
            key=f"items_editor_{selected_rec.record_id}",
        )

        if st.button("💾 Lưu chi tiết hàng hóa"):
            new_items: List[InvoiceItem] = []
            for _, row in edited_items_df.iterrows():
                name = str(row.get("name") or "").strip()
                amount = float(row.get("amount") or 0.0)
                if name or amount > 0:
                    new_items.append(
                        InvoiceItem(
                            name=name,
                            unit=str(row.get("unit") or "").strip(),
                            quantity=float(row.get("quantity") or 0.0),
                            unit_price=float(row.get("unit_price") or 0.0),
                            amount=amount,
                        )
                    )
            selected_rec.data.items = new_items
            st.session_state["records"] = validate_batch(
                records, amount_tolerance=amount_tolerance
            )
            st.success("Đã cập nhật chi tiết hàng hóa và đối chiếu lại!")
            st.rerun()

    # ------------------------------------------------------------------
    # TAB 3: Xuất Excel & JSON
    # ------------------------------------------------------------------
    with tab_export:
        st.subheader("Xuất báo cáo sổ sách kế toán (.xlsx)")
        st.markdown(
            "File Excel xuất ra gồm **2 sheet chuẩn kế toán**:\n"
            "- **Sheet 1 (`Bảng kê hóa đơn tổng hợp`)**: Định dạng border, tô đỏ dòng cần kiểm tra, có dòng **TỔNG CỘNG** sử dụng công thức `=SUM(...)`.\n"
            "- **Sheet 2 (`Chi tiết hàng hóa`)**: Toàn bộ mặt hàng liên kết theo Ký hiệu/Số hóa đơn, kèm dòng tổng cộng."
        )

        excel_bytes = export_to_excel_bytes(records)
        filename = f"BangKeHoaDon_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                label="📊 Tải xuống file Excel kế toán (.xlsx)",
                data=excel_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
        with col_dl2:
            json_payload = json.dumps(
                [r.model_dump(mode="json") for r in records],
                ensure_ascii=False,
                indent=2,
            )
            st.download_button(
                label="🧾 Tải xuống dữ liệu chuẩn hóa (.json)",
                data=json_payload.encode("utf-8"),
                file_name=f"VN_InvoiceParser_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                mime="application/json",
                use_container_width=True,
            )

        with st.expander("Xem trước dữ liệu JSON chuẩn hóa"):
            st.json([r.model_dump(mode="json") for r in records])


if __name__ == "__main__":
    main()
