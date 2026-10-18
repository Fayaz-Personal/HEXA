# services/document_export/models.py
"""
Data models and request schemas for HEXA Enterprise Document Export System.
"""

from datetime import datetime
from typing import Any, Dict, List, Optional
from pydantic import BaseModel, Field


class DocumentTable(BaseModel):
    """Normalized representation of a tabular data structure."""
    headers: List[str] = Field(default_factory=list)
    rows: List[List[str]] = Field(default_factory=list)
    caption: Optional[str] = None


class DocumentCodeBlock(BaseModel):
    """Normalized code block representation."""
    language: str = "text"
    code: str = ""


class DocumentReference(BaseModel):
    """Source / citation reference representation."""
    source_name: str
    page: Optional[str] = None
    url: Optional[str] = None
    citation: Optional[str] = None


class DocumentSection(BaseModel):
    """A section of a document containing headings, text, lists, and tables."""
    title: str = ""
    level: int = 1  # 1 = H1, 2 = H2, 3 = H3
    paragraphs: List[str] = Field(default_factory=list)
    bullet_lists: List[List[str]] = Field(default_factory=list)
    numbered_lists: List[List[str]] = Field(default_factory=list)
    tables: List[DocumentTable] = Field(default_factory=list)
    code_blocks: List[DocumentCodeBlock] = Field(default_factory=list)


class DocumentModel(BaseModel):
    """
    Normalized internal document representation.
    All document generators (PDF, DOCX, PPTX, XLSX) consume this model.
    """
    title: str = "HEXA AI Generated Report"
    author: str = "HEXA AI"
    created_at: str = Field(default_factory=lambda: datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    source_type: str = "AI Response"
    summary: str = ""
    sections: List[DocumentSection] = Field(default_factory=list)
    tables: List[DocumentTable] = Field(default_factory=list)
    code_blocks: List[DocumentCodeBlock] = Field(default_factory=list)
    references: List[DocumentReference] = Field(default_factory=list)
    metadata: Dict[str, Any] = Field(default_factory=dict)


# ---------------------------------------------------------------------------
# Format Options & Request Models
# ---------------------------------------------------------------------------

class PDFExportOptions(BaseModel):
    include_header: bool = True
    include_datetime: bool = True
    include_page_numbers: bool = True
    include_sources: bool = True


class WordExportOptions(BaseModel):
    include_title_page: bool = True
    include_toc: bool = True
    include_headers_footers: bool = True
    include_sources: bool = True


class PPTXExportOptions(BaseModel):
    max_slides: int = 25
    include_title_slide: bool = True
    include_summary_slide: bool = True
    include_source_slide: bool = True


class ExcelExportOptions(BaseModel):
    auto_worksheets: bool = True
    include_source_info: bool = True
    auto_format_tables: bool = True
    freeze_header: bool = True
    auto_size_columns: bool = True


class ExportRequest(BaseModel):
    title: str = "HEXA AI Generated Report"
    author: str = "HEXA AI"
    content: str = ""
    source_type: str = "current_response"
    structured_data: Optional[Dict[str, Any]] = None
    pdf_options: PDFExportOptions = Field(default_factory=PDFExportOptions)
    word_options: WordExportOptions = Field(default_factory=WordExportOptions)
    pptx_options: PPTXExportOptions = Field(default_factory=PPTXExportOptions)
    excel_options: ExcelExportOptions = Field(default_factory=ExcelExportOptions)
