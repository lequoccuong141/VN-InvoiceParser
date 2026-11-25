"""Cấu hình tập trung cho ứng dụng VN-InvoiceParser."""

from __future__ import annotations

import os
from dataclasses import dataclass
from dotenv import load_dotenv

load_dotenv()


@dataclass(frozen=True)
class Settings:
    """Thiết lập runtime đọc từ biến môi trường."""

    gemini_api_key: str = os.getenv("GEMINI_API_KEY", "")
    gemini_model: str = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")
    pdf_render_dpi: int = int(os.getenv("PDF_RENDER_DPI", "250"))
    amount_tolerance: float = float(os.getenv("AMOUNT_TOLERANCE", "1.0"))
    tesseract_cmd: str = os.getenv("TESSERACT_CMD", "")
    min_pdf_text_chars: int = int(os.getenv("MIN_PDF_TEXT_CHARS", "80"))
    max_pdf_pages: int = int(os.getenv("MAX_PDF_PAGES", "5"))


def get_settings() -> Settings:
    """Trả về cấu hình hiện tại."""
    load_dotenv(override=False)
    return Settings(
        gemini_api_key=os.getenv("GEMINI_API_KEY", ""),
        gemini_model=os.getenv("GEMINI_MODEL", "gemini-2.5-flash"),
        pdf_render_dpi=int(os.getenv("PDF_RENDER_DPI", "250")),
        amount_tolerance=float(os.getenv("AMOUNT_TOLERANCE", "1.0")),
        tesseract_cmd=os.getenv("TESSERACT_CMD", ""),
        min_pdf_text_chars=int(os.getenv("MIN_PDF_TEXT_CHARS", "80")),
        max_pdf_pages=int(os.getenv("MAX_PDF_PAGES", "5")),
    )
