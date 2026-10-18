# routes/export_routes.py
"""
FastAPI router for HEXA Enterprise Document Generation & Export System.
Provides endpoints for PDF, DOCX, PPTX, XLSX, ZIP exports and file downloads.
"""

import logging
import os
import re
from typing import Any, Dict

from fastapi import APIRouter, HTTPException, Request
from fastapi.responses import FileResponse, JSONResponse

from src.auth_helpers import storage_owner_for_request
from src.constants import EXPORTS_DIR
from services.document_export.export_service import ExportService
from services.document_export.models import ExportRequest

logger = logging.getLogger("hexa.export_routes")


def setup_export_routes() -> APIRouter:
    router = APIRouter(prefix="/api/export", tags=["export"])

    # ── PDF Export ────────────────────────────────────────────────────────
    @router.post("/pdf")
    async def export_pdf(req: ExportRequest, request: Request):
        """Generate a professional PDF document."""
        try:
            owner_id = storage_owner_for_request(request) or "default"
            result = ExportService.export_pdf(req, owner_id=owner_id)
            return JSONResponse(result)
        except Exception as e:
            logger.exception("PDF export failed")
            return JSONResponse({"error": f"PDF export failed: {str(e)}"}, status_code=500)

    # ── DOCX Export ───────────────────────────────────────────────────────
    @router.post("/docx")
    async def export_docx(req: ExportRequest, request: Request):
        """Generate a Word (.docx) document."""
        try:
            owner_id = storage_owner_for_request(request) or "default"
            result = ExportService.export_docx(req, owner_id=owner_id)
            return JSONResponse(result)
        except Exception as e:
            logger.exception("DOCX export failed")
            return JSONResponse({"error": f"Word document export failed: {str(e)}"}, status_code=500)

    # ── PPTX Export ───────────────────────────────────────────────────────
    @router.post("/pptx")
    async def export_pptx(req: ExportRequest, request: Request):
        """Generate a PowerPoint (.pptx) presentation."""
        try:
            owner_id = storage_owner_for_request(request) or "default"
            result = ExportService.export_pptx(req, owner_id=owner_id)
            return JSONResponse(result)
        except Exception as e:
            logger.exception("PPTX export failed")
            return JSONResponse({"error": f"PowerPoint export failed: {str(e)}"}, status_code=500)

    # ── XLSX Export ───────────────────────────────────────────────────────
    @router.post("/xlsx")
    async def export_xlsx(req: ExportRequest, request: Request):
        """Generate an Excel (.xlsx) spreadsheet."""
        try:
            owner_id = storage_owner_for_request(request) or "default"
            result = ExportService.export_xlsx(req, owner_id=owner_id)
            return JSONResponse(result)
        except Exception as e:
            logger.exception("XLSX export failed")
            return JSONResponse({"error": f"Excel export failed: {str(e)}"}, status_code=500)

    # ── Generate All ──────────────────────────────────────────────────────
    @router.post("/all")
    async def export_all(req: ExportRequest, request: Request):
        """Generate PDF, DOCX, PPTX, XLSX and return a single ZIP archive."""
        try:
            owner_id = storage_owner_for_request(request) or "default"
            result = ExportService.export_all(req, owner_id=owner_id)
            return JSONResponse(result)
        except Exception as e:
            logger.exception("Export all failed")
            return JSONResponse({"error": f"Generate all export failed: {str(e)}"}, status_code=500)

    # ── Preview ───────────────────────────────────────────────────────────
    @router.post("/preview")
    async def export_preview(req: ExportRequest, request: Request):
        """Returns normalized DocumentModel structure for client-side preview."""
        try:
            doc_model = ExportService.parse_request(req)
            return JSONResponse({
                "status": "ok",
                "document": doc_model.model_dump(),
            })
        except Exception as e:
            logger.exception("Export preview failed")
            return JSONResponse({"error": f"Preview generation failed: {str(e)}"}, status_code=500)

    # ── Secure Download Endpoint ──────────────────────────────────────────
    @router.get("/download/{export_id}/{filename}")
    async def download_export(export_id: str, filename: str, request: Request):
        """Secure download endpoint for generated export files."""
        owner_id = storage_owner_for_request(request) or "default"

        # Sanitize parameters to prevent path traversal
        safe_owner = re.sub(r"[^\w-]", "_", owner_id)
        safe_export_id = re.sub(r"[^\w-]", "_", export_id)
        safe_filename = os.path.basename(filename)

        # Security check & fallback resolution across potential owner subdirectories
        base_exports_dir = os.path.abspath(EXPORTS_DIR)
        candidate_paths = [
            os.path.abspath(os.path.join(base_exports_dir, safe_owner, safe_export_id, safe_filename)),
            os.path.abspath(os.path.join(base_exports_dir, "default", safe_export_id, safe_filename)),
            os.path.abspath(os.path.join(base_exports_dir, "local", safe_export_id, safe_filename)),
        ]

        file_path = None
        for cp in candidate_paths:
            if cp.startswith(base_exports_dir) and os.path.isfile(cp):
                file_path = cp
                break

        if not file_path:
            raise HTTPException(status_code=404, detail="File not found")

        ext = os.path.splitext(safe_filename)[1].lower()
        mime_types = {
            ".pdf": "application/pdf",
            ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".zip": "application/zip",
        }
        media_type = mime_types.get(ext, "application/octet-stream")

        return FileResponse(
            file_path,
            media_type=media_type,
            filename=safe_filename,
            headers={"Content-Disposition": f'attachment; filename="{safe_filename}"'},
        )

    return router
