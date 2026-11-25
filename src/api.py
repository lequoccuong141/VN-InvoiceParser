"""FastAPI REST API backend cho VN-InvoiceParser (dành cho tích hợp ERP/Next.js)."""

from __future__ import annotations

import io
from datetime import datetime
from typing import List, Optional

from fastapi import FastAPI, File, Header, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse

from src.config import get_settings
from src.exporter import export_to_excel_bytes
from src.extractor import InvoiceExtractor
from src.schemas import InvoiceRecord
from src.validator import validate_batch

app = FastAPI(
    title="VN-InvoiceParser API",
    description="API trích xuất, đối chiếu và xuất Excel hóa đơn điện tử Việt Nam (PDF, JPG, PNG, XML).",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/health")
def health_check() -> dict:
    """Kiểm tra trạng thái dịch vụ."""
    settings = get_settings()
    return {
        "status": "ok",
        "model": settings.gemini_model,
        "gemini_configured": bool(settings.gemini_api_key),
    }


@app.post("/api/v1/extract", response_model=List[InvoiceRecord])
async def extract_invoices(
    files: List[UploadFile] = File(..., description="Danh sách file hóa đơn (PDF, PNG, JPG, JPEG, XML)"),
    amount_tolerance: float = Query(default=1.0, ge=0.0, description="Dung sai làm tròn tiền (VNĐ)"),
    use_fallback_only: bool = Query(default=False, description="Bỏ qua Gemini, dùng OCR/Regex fallback"),
    x_gemini_api_key: Optional[str] = Header(default=None, alias="X-Gemini-Api-Key"),
) -> List[InvoiceRecord]:
    """Tải lên hàng loạt hóa đơn, trích xuất JSON chuẩn hóa và tự động đối chiếu lỗi."""
    if not files:
        raise HTTPException(status_code=400, detail="Vui lòng tải lên ít nhất 1 file hóa đơn.")

    extractor = InvoiceExtractor(
        api_key=x_gemini_api_key,
        use_fallback_only=use_fallback_only,
    )

    records: List[InvoiceRecord] = []
    for upload in files:
        file_bytes = await upload.read()
        filename = upload.filename or "unnamed_file"
        records.extend(extractor.extract_batch_from_bytes(file_bytes, filename))

    return validate_batch(records, amount_tolerance=amount_tolerance)


@app.post("/api/v1/export-excel")
def export_invoices_excel(
    records: List[InvoiceRecord],
    amount_tolerance: float = Query(default=1.0, ge=0.0),
) -> StreamingResponse:
    """Đối chiếu lại danh sách hóa đơn đã chỉnh sửa và xuất file Excel 2 sheet."""
    validated = validate_batch(records, amount_tolerance=amount_tolerance)
    excel_bytes = export_to_excel_bytes(validated)
    filename = f"BangKeHoaDon_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/api/v1/extract-and-export")
async def extract_and_export_excel(
    files: List[UploadFile] = File(...),
    amount_tolerance: float = Query(default=1.0, ge=0.0),
    use_fallback_only: bool = Query(default=False),
    x_gemini_api_key: Optional[str] = Header(default=None, alias="X-Gemini-Api-Key"),
) -> StreamingResponse:
    """Trích xuất hàng loạt hóa đơn và trả về trực tiếp file Excel kế toán."""
    records = await extract_invoices(
        files=files,
        amount_tolerance=amount_tolerance,
        use_fallback_only=use_fallback_only,
        x_gemini_api_key=x_gemini_api_key,
    )
    excel_bytes = export_to_excel_bytes(records)
    filename = f"BangKeHoaDon_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
