"""
Kohler Excel Automation Package

This package provides automation for:
- Extracting product images from Kohler pricebook PDFs
- Auto-generating summary sheets from Excel workbooks
"""

from .pdf_parser import KohlerPDFParser, parse_pdf
from .image_extractor import ImageExtractor, extract_images_from_pdf
from .excel_handler import ExcelHandler
from .summary_builder import SummaryBuilder
from .config import Config

__all__ = [
    "KohlerPDFParser",
    "parse_pdf",
    "ImageExtractor",
    "extract_images_from_pdf",
    "ExcelHandler",
    "SummaryBuilder",
    "Config",
]