"""Định nghĩa Pydantic schemas chuẩn hóa cho hóa đơn điện tử Việt Nam."""

from __future__ import annotations

from enum import Enum
from typing import List, Optional
from pydantic import BaseModel, Field


class InvoiceStatus(str, Enum):
    """Trạng thái đối chiếu hóa đơn."""

    VALID = "Hợp lệ"
    NEEDS_REVIEW = "Cần kiểm tra"
    ERROR = "Lỗi xử lý"


class InvoiceItem(BaseModel):
    """Một dòng hàng hóa / dịch vụ trên hóa đơn."""

    name: str = Field(default="", description="Tên hàng hóa, dịch vụ")
    unit: str = Field(default="", description="Đơn vị tính (VD: Cái, Bộ, Kg, Dịch vụ)")
    quantity: float = Field(default=0.0, description="Số lượng")
    unit_price: float = Field(default=0.0, description="Đơn giá chưa thuế")
    amount: float = Field(default=0.0, description="Thành tiền chưa thuế")


class InvoiceData(BaseModel):
    """Schema JSON chuẩn hóa trả về từ AI/OCR hoặc XML parser."""

    invoice_symbol: str = Field(
        default="",
        description="Ký hiệu hóa đơn (VD: 1C26TAA, 1K25TYY)",
    )
    invoice_number: str = Field(
        default="",
        description="Số hóa đơn (VD: 00001234)",
    )
    invoice_date: str = Field(
        default="",
        description="Ngày lập hóa đơn chuẩn hóa theo định dạng DD/MM/YYYY",
    )
    seller_name: str = Field(
        default="",
        description="Tên đơn vị bán hàng",
    )
    seller_tax_code: str = Field(
        default="",
        description="Mã số thuế người bán (10 số hoặc 10 số-3 số)",
    )
    buyer_name: str = Field(
        default="",
        description="Tên đơn vị mua hàng hoặc tên người mua",
    )
    buyer_tax_code: str = Field(
        default="",
        description="Mã số thuế người mua (10 số hoặc 10 số-3 số, để trống nếu khách lẻ)",
    )
    sub_total: float = Field(
        default=0.0,
        description="Cộng tiền hàng chưa bao gồm thuế GTGT (VNĐ)",
    )
    vat_rate: str = Field(
        default="",
        description="Thuế suất GTGT (VD: 0%, 5%, 8%, 10%, KCT)",
    )
    vat_amount: float = Field(
        default=0.0,
        description="Tiền thuế GTGT (VNĐ)",
    )
    total_amount: float = Field(
        default=0.0,
        description="Tổng cộng tiền thanh toán đã gồm thuế GTGT (VNĐ)",
    )
    items: List[InvoiceItem] = Field(
        default_factory=list,
        description="Danh sách chi tiết hàng hóa, dịch vụ trên hóa đơn",
    )


class InvoiceRecord(BaseModel):
    """Bản ghi hóa đơn hoàn chỉnh kèm metadata tệp nguồn và kết quả đối chiếu."""

    record_id: str = Field(description="ID duy nhất của bản ghi trong phiên làm việc")
    source_file: str = Field(description="Tên file gốc người dùng tải lên")
    extraction_method: str = Field(
        default="unknown",
        description="Phương thức trích xuất (xml, pdf_text_gemini, pdf_scan_gemini, image_gemini, fallback_regex)",
    )
    data: InvoiceData = Field(default_factory=InvoiceData)
    status: InvoiceStatus = Field(default=InvoiceStatus.VALID)
    validation_errors: List[str] = Field(default_factory=list)
    validation_warnings: List[str] = Field(default_factory=list)
    raw_error: Optional[str] = Field(default=None, description="Lỗi kỹ thuật nếu quá trình đọc file thất bại")

    def to_summary_dict(self) -> dict:
        """Chuyển đổi sang dạng phẳng để hiển thị trên bảng UI và xuất Excel."""
        return {
            "record_id": self.record_id,
            "source_file": self.source_file,
            "extraction_method": self.extraction_method,
            "invoice_symbol": self.data.invoice_symbol,
            "invoice_number": self.data.invoice_number,
            "invoice_date": self.data.invoice_date,
            "seller_name": self.data.seller_name,
            "seller_tax_code": self.data.seller_tax_code,
            "buyer_name": self.data.buyer_name,
            "buyer_tax_code": self.data.buyer_tax_code,
            "sub_total": self.data.sub_total,
            "vat_rate": self.data.vat_rate,
            "vat_amount": self.data.vat_amount,
            "total_amount": self.data.total_amount,
            "status": self.status.value,
            "validation_errors": "; ".join(self.validation_errors),
            "validation_warnings": "; ".join(self.validation_warnings),
        }
