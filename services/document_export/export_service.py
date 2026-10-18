# services/document_export/export_service.py
"""
Export Service Orchestrator for HEXA Enterprise Document Export System.
Parses content and generates PDF, DOCX, PPTX, XLSX deliverables and ZIP bundles.
"""

import os
import re
import shutil
import uuid
import zipfile
from datetime import datetime
from typing import Any, Dict, List, Tuple

from src.constants import EXPORTS_DIR
from services.document_export.docx_generator import DocxGenerator
from services.document_export.models import DocumentModel, ExportRequest
from services.document_export.parser import ContentParser
from services.document_export.pdf_generator import PDFGenerator
from services.document_export.pptx_generator import PPTXGenerator
from services.document_export.xlsx_generator import XLSXGenerator


class ExportService:
    """Central orchestrator for document export operations."""

    @classmethod
    def sanitize_filename(cls, title: str, ext: str) -> str:
        """Sanitizes user titles into safe filesystem filenames."""
        safe = re.sub(r"[^\w\s-]", "", title or "Document").strip()
        safe = re.sub(r"[\s]+", "_", safe)
        safe = safe[:60] or "HEXA_Export"
        return f"{safe}.{ext.lstrip('.')}"

    @classmethod
    def get_user_export_dir(cls, owner_id: str, export_id: str) -> str:
        """Returns isolated user export directory path."""
        safe_owner = re.sub(r"[^\w-]", "_", owner_id or "default")
        safe_export_id = re.sub(r"[^\w-]", "_", export_id or str(uuid.uuid4())[:8])
        export_dir = os.path.abspath(os.path.join(EXPORTS_DIR, safe_owner, safe_export_id))

        # Security check: Ensure export_dir is inside EXPORTS_DIR
        base_dir = os.path.abspath(EXPORTS_DIR)
        if not export_dir.startswith(base_dir):
            raise ValueError("Invalid export path — path traversal detected.")

        os.makedirs(export_dir, exist_ok=True)
        return export_dir

    @classmethod
    def parse_request(cls, req: ExportRequest) -> DocumentModel:
        """Converts an ExportRequest into a normalized DocumentModel."""
        return ContentParser.parse(
            content=req.content,
            title=req.title,
            author=req.author,
            source_type=req.source_type,
            structured_data=req.structured_data,
        )

    @classmethod
    def export_pdf(cls, req: ExportRequest, owner_id: str = "default") -> Dict[str, Any]:
        export_id = str(uuid.uuid4())[:8]
        target_dir = cls.get_user_export_dir(owner_id, export_id)
        doc_model = cls.parse_request(req)
        fname = cls.sanitize_filename(doc_model.title, "pdf")
        file_path = os.path.join(target_dir, fname)

        PDFGenerator.generate(doc_model, req.pdf_options, file_path)

        return {
            "status": "ok",
            "export_id": export_id,
            "filename": fname,
            "format": "pdf",
            "file_size": os.path.getsize(file_path),
            "download_url": f"/api/export/download/{export_id}/{fname}",
            "message": f"PDF '{doc_model.title}' generated successfully.",
        }

    @classmethod
    def export_docx(cls, req: ExportRequest, owner_id: str = "default") -> Dict[str, Any]:
        export_id = str(uuid.uuid4())[:8]
        target_dir = cls.get_user_export_dir(owner_id, export_id)
        doc_model = cls.parse_request(req)
        fname = cls.sanitize_filename(doc_model.title, "docx")
        file_path = os.path.join(target_dir, fname)

        DocxGenerator.generate(doc_model, req.word_options, file_path)

        return {
            "status": "ok",
            "export_id": export_id,
            "filename": fname,
            "format": "docx",
            "file_size": os.path.getsize(file_path),
            "download_url": f"/api/export/download/{export_id}/{fname}",
            "message": f"Word document '{doc_model.title}' generated successfully.",
        }

    @classmethod
    def export_pptx(cls, req: ExportRequest, owner_id: str = "default") -> Dict[str, Any]:
        export_id = str(uuid.uuid4())[:8]
        target_dir = cls.get_user_export_dir(owner_id, export_id)
        doc_model = cls.parse_request(req)
        fname = cls.sanitize_filename(doc_model.title, "pptx")
        file_path = os.path.join(target_dir, fname)

        PPTXGenerator.generate(doc_model, req.pptx_options, file_path)

        return {
            "status": "ok",
            "export_id": export_id,
            "filename": fname,
            "format": "pptx",
            "file_size": os.path.getsize(file_path),
            "download_url": f"/api/export/download/{export_id}/{fname}",
            "message": f"PowerPoint '{doc_model.title}' generated successfully.",
        }

    @classmethod
    def export_xlsx(cls, req: ExportRequest, owner_id: str = "default") -> Dict[str, Any]:
        export_id = str(uuid.uuid4())[:8]
        target_dir = cls.get_user_export_dir(owner_id, export_id)
        doc_model = cls.parse_request(req)
        fname = cls.sanitize_filename(doc_model.title, "xlsx")
        file_path = os.path.join(target_dir, fname)

        XLSXGenerator.generate(doc_model, req.excel_options, file_path)

        return {
            "status": "ok",
            "export_id": export_id,
            "filename": fname,
            "format": "xlsx",
            "file_size": os.path.getsize(file_path),
            "download_url": f"/api/export/download/{export_id}/{fname}",
            "message": f"Excel spreadsheet '{doc_model.title}' generated successfully.",
        }

    @classmethod
    def export_all(cls, req: ExportRequest, owner_id: str = "default") -> Dict[str, Any]:
        """
        Generates PDF, DOCX, PPTX, XLSX files and packages them into
        a single ZIP archive: HEXA_Export_<timestamp>.zip
        """
        export_id = str(uuid.uuid4())[:8]
        target_dir = cls.get_user_export_dir(owner_id, export_id)
        doc_model = cls.parse_request(req)

        pdf_name = cls.sanitize_filename(doc_model.title, "pdf")
        docx_name = cls.sanitize_filename(doc_model.title, "docx")
        pptx_name = cls.sanitize_filename(doc_model.title, "pptx")
        xlsx_name = cls.sanitize_filename(doc_model.title, "xlsx")

        pdf_path = os.path.join(target_dir, pdf_name)
        docx_path = os.path.join(target_dir, docx_name)
        pptx_path = os.path.join(target_dir, pptx_name)
        xlsx_path = os.path.join(target_dir, xlsx_name)

        # Generate all 4 files
        PDFGenerator.generate(doc_model, req.pdf_options, pdf_path)
        DocxGenerator.generate(doc_model, req.word_options, docx_path)
        PPTXGenerator.generate(doc_model, req.pptx_options, pptx_path)
        XLSXGenerator.generate(doc_model, req.excel_options, xlsx_path)

        # Create ZIP archive
        ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        safe_title = re.sub(r"[^\w-]", "_", doc_model.title or "Report")[:30]
        zip_name = f"HEXA_Export_{safe_title}_{ts}.zip"
        zip_path = os.path.join(target_dir, zip_name)

        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(pdf_path, arcname=pdf_name)
            zf.write(docx_path, arcname=docx_name)
            zf.write(pptx_path, arcname=pptx_name)
            zf.write(xlsx_path, arcname=xlsx_name)

        files_info = [
            {"filename": pdf_name, "format": "pdf", "size": os.path.getsize(pdf_path), "download_url": f"/api/export/download/{export_id}/{pdf_name}"},
            {"filename": docx_name, "format": "docx", "size": os.path.getsize(docx_path), "download_url": f"/api/export/download/{export_id}/{docx_name}"},
            {"filename": pptx_name, "format": "pptx", "size": os.path.getsize(pptx_path), "download_url": f"/api/export/download/{export_id}/{pptx_name}"},
            {"filename": xlsx_name, "format": "xlsx", "size": os.path.getsize(xlsx_path), "download_url": f"/api/export/download/{export_id}/{xlsx_name}"},
        ]

        return {
            "status": "ok",
            "export_id": export_id,
            "filename": zip_name,
            "format": "zip",
            "file_size": os.path.getsize(zip_path),
            "download_url": f"/api/export/download/{export_id}/{zip_name}",
            "files": files_info,
            "message": f"Exported all formats for '{doc_model.title}' into ZIP bundle.",
        }
