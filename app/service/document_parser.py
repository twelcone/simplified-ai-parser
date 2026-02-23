"""
Document Parser - Main coordinator for file parsing

This is the simplified version of the original ai-parser's DocumentParser.
It converts various document formats to Markdown.

Key differences from original:
- Output: Markdown (original outputs HTML)
- XLSX: Uses openpyxl directly (original uses LibreOffice for HTML conversion)
- No Celery/async processing
- No Azure storage integration
- No PII detection/masking

What's preserved:
- Same file type support: .docx, .xlsx, .xls, .xlsm, .pdf, .md, .markdown
- Same image handling: base64 encoding, only png/jpg/jpeg supported
- Same HTML cleaning logic for PDF conversion
- Same mutool usage for PDF → HTML
- Same mammoth usage for DOCX → HTML → Markdown
"""

import os
import logging
from typing import Tuple

from app.service.docx_parser import parse_docx_to_markdown
from app.service.xlsx_parser import parse_xlsx_to_markdown
from app.service.pdf_parser import parse_pdf_to_markdown
from app.service.markdown_parser import parse_markdown
from app.service.pptx_parser import parse_pptx_to_markdown

logger = logging.getLogger(__name__)

# Supported extensions matching original document_parser.py
SUPPORTED_EXTENSIONS = {
    ".docx": "docx",
    ".xlsx": "xlsx",
    ".xls": "xls",
    ".xlsm": "xlsm",
    ".pdf": "pdf",
    ".md": "markdown",
    ".markdown": "markdown",
    ".pptx": "pptx",
    ".ppt": "ppt",
}


def get_file_type(filename: str) -> Tuple[str, str]:
    """
    Get file type from filename.

    Args:
        filename: Name of the file

    Returns:
        Tuple of (extension, file_type) or raises ValueError if unsupported
    """
    _, ext = os.path.splitext(filename.lower())

    if ext not in SUPPORTED_EXTENSIONS:
        raise ValueError(f"Unsupported file type: {ext}")

    return ext, SUPPORTED_EXTENSIONS[ext]


def parse_document(file_path: str, file_type: str) -> str:
    """
    Parse a document and convert it to Markdown.

    Routing logic matches original document_parser.py convert_document_to_html:
    - .docx: mammoth → HTML → markdownify → Markdown
    - .xlsx/.xlsm: openpyxl → Markdown tables (original uses LibreOffice → HTML)
    - .xls: Same as xlsx (may require conversion for old formats)
    - .pdf: mutool → HTML → markdownify → Markdown
    - .md/.markdown: Passthrough with image filtering

    Args:
        file_path: Path to the document file
        file_type: Type of the file (docx, xlsx, xls, xlsm, pdf, markdown)

    Returns:
        Markdown content as string
    """
    logger.info(f"Parsing {file_type} file: {file_path}")

    if file_type == "docx":
        return parse_docx_to_markdown(file_path)

    elif file_type in ("xlsx", "xlsm"):
        return parse_xlsx_to_markdown(file_path)

    elif file_type == "xls":
        # XLS (legacy Excel format) handling
        # Original uses LibreOffice for conversion
        # openpyxl may not support old .xls formats
        try:
            return parse_xlsx_to_markdown(file_path)
        except Exception as e:
            logger.warning(f"Direct XLS parsing failed: {e}")
            raise RuntimeError(
                "XLS format (Excel 97-2003) may require LibreOffice for conversion. "
                "Please convert to XLSX format, or ensure the file is a valid Excel format."
            )

    elif file_type == "pdf":
        return parse_pdf_to_markdown(file_path)

    elif file_type == "markdown":
        return parse_markdown(file_path)

    elif file_type == "pptx":
        return parse_pptx_to_markdown(file_path)

    elif file_type == "ppt":
        # PPT (legacy PowerPoint format) handling
        # python-pptx may not fully support old .ppt formats
        try:
            return parse_pptx_to_markdown(file_path)
        except Exception as e:
            logger.warning(f"Direct PPT parsing failed: {e}")
            raise RuntimeError(
                "PPT format (PowerPoint 97-2003) may not be fully supported. "
                "Please convert to PPTX format for best results."
            )

    else:
        raise ValueError(f"Unknown file type: {file_type}")
