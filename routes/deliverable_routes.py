"""
routes/deliverable_routes.py

Direct API endpoints for generating real Office deliverables:
  POST /api/deliverable/docx   → Word (.docx)
  POST /api/deliverable/xlsx   → Excel (.xlsx)
  POST /api/deliverable/pptx   → PowerPoint (.pptx)
  POST /api/deliverable/pdf    → PDF (from markdown/HTML content)
  GET  /api/deliverable/list   → list generated files
  GET  /api/deliverable/download/{filename} → download a file

These endpoints work independently of the AI agent/model — no tool-calling
required. The model can also trigger them, but users can call them directly
from the UI without relying on small-context-window models.
"""

import os
import io
import re
import uuid
import logging
from datetime import datetime
from typing import Any, Dict, List, Optional

from fastapi import APIRouter, Request
from fastapi.responses import FileResponse, JSONResponse

from src.constants import DATA_DIR

logger = logging.getLogger("hexa.deliverable_routes")

DELIVERABLES_DIR = os.path.join(DATA_DIR, "deliverables")
os.makedirs(DELIVERABLES_DIR, exist_ok=True)


def _safe_filename(name: str, ext: str) -> str:
    """Sanitize a title into a safe filename."""
    safe = re.sub(r"[^\w\s-]", "", name).strip()
    safe = re.sub(r"[\s]+", "_", safe)
    safe = safe[:60] or "document"
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{safe}_{ts}.{ext}"


# ---------------------------------------------------------------------------
# Word (.docx)
# ---------------------------------------------------------------------------

def _build_docx(title: str, sections: List[Dict], tables: Optional[List] = None) -> str:
    """Generate a .docx file and return its path."""
    import docx
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = docx.Document()

    # Page margins
    for section in doc.sections:
        section.top_margin = Pt(72)
        section.bottom_margin = Pt(72)
        section.left_margin = Pt(72)
        section.right_margin = Pt(72)

    # Title
    title_para = doc.add_heading(title, level=0)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title_para.runs:
        run.font.color.rgb = RGBColor(0, 95, 115)
        run.font.name = "Calibri"
        run.font.size = Pt(24)

    doc.add_paragraph()  # spacer

    # Sections
    for sec in sections:
        sec_title = sec.get("title", "").strip()
        sec_body = sec.get("body", "").strip()

        if sec_title:
            h = doc.add_heading(sec_title, level=1)
            for run in h.runs:
                run.font.color.rgb = RGBColor(10, 147, 150)
                run.font.name = "Calibri"
                run.font.size = Pt(14)

        if sec_body:
            # Split on double newlines for paragraphs
            for para_text in sec_body.split("\n\n"):
                para_text = para_text.strip()
                if para_text:
                    p = doc.add_paragraph(para_text)
                    p.paragraph_format.space_after = Pt(8)

    # Tables
    if tables:
        for tbl_data in tables:
            if not tbl_data:
                continue
            rows = len(tbl_data)
            cols = max(len(r) for r in tbl_data) if rows > 0 else 0
            if rows == 0 or cols == 0:
                continue
            table = doc.add_table(rows=rows, cols=cols)
            table.style = "Table Grid"
            for r_idx, row in enumerate(tbl_data):
                for c_idx, val in enumerate(row):
                    if c_idx < cols:
                        cell = table.cell(r_idx, c_idx)
                        cell.text = str(val)
                        if r_idx == 0:
                            for run in cell.paragraphs[0].runs:
                                run.font.bold = True

    fname = _safe_filename(title, "docx")
    path = os.path.join(DELIVERABLES_DIR, fname)
    doc.save(path)
    logger.info(f"Word deliverable saved: {path}")
    return fname


# ---------------------------------------------------------------------------
# Excel (.xlsx)
# ---------------------------------------------------------------------------

def _build_xlsx(title: str, sheets: Dict[str, List[List[Any]]]) -> str:
    """Generate a .xlsx file and return its filename."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default blank sheet

    header_fill = PatternFill(start_color="0A9396", end_color="0A9396", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    body_font = Font(name="Calibri", size=11)

    for sheet_name, rows in sheets.items():
        ws = wb.create_sheet(title=sheet_name[:31])
        for r_idx, row_data in enumerate(rows, start=1):
            for c_idx, val in enumerate(row_data, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                if r_idx == 1:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.font = body_font
        # Auto-width columns
        for col in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=8)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 50)

    fname = _safe_filename(title, "xlsx")
    path = os.path.join(DELIVERABLES_DIR, fname)
    wb.save(path)
    logger.info(f"Excel deliverable saved: {path}")
    return fname


# ---------------------------------------------------------------------------
# PowerPoint (.pptx)
# ---------------------------------------------------------------------------

def _build_pptx(title: str, slides: List[Dict]) -> str:
    """Generate a .pptx file and return its filename."""
    from pptx import Presentation
    from pptx.util import Pt, Inches
    from pptx.dml.color import RGBColor

    prs = Presentation()
    TEAL = RGBColor(0, 95, 115)

    # Title slide
    title_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_layout)
    slide.shapes.title.text = title
    if slide.placeholders[1]:
        slide.placeholders[1].text = "HEXA AI Workbench"

    # Content slides
    bullet_layout = prs.slide_layouts[1]
    for s in slides:
        slide = prs.slides.add_slide(bullet_layout)
        slide.shapes.title.text = s.get("title", "")
        tf = slide.placeholders[1].text_frame
        points = s.get("points", [])
        for i, pt in enumerate(points):
            if i == 0:
                tf.text = pt
            else:
                p = tf.add_paragraph()
                p.text = pt

    fname = _safe_filename(title, "pptx")
    path = os.path.join(DELIVERABLES_DIR, fname)
    prs.save(path)
    logger.info(f"PowerPoint deliverable saved: {path}")
    return fname


# ---------------------------------------------------------------------------
# PDF (from markdown content via reportlab or html2pdf fallback)
# ---------------------------------------------------------------------------

def _build_pdf(title: str, content: str) -> str:
    """Generate a .pdf from plain text/markdown content."""
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import mm
        from reportlab.lib import colors
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer

        fname = _safe_filename(title, "pdf")
        path = os.path.join(DELIVERABLES_DIR, fname)

        doc = SimpleDocTemplate(path, pagesize=A4,
                                leftMargin=25*mm, rightMargin=25*mm,
                                topMargin=25*mm, bottomMargin=25*mm)
        styles = getSampleStyleSheet()
        teal = colors.Color(0/255, 95/255, 115/255)

        title_style = ParagraphStyle("HexaTitle", parent=styles["Title"],
                                     textColor=teal, fontSize=22, spaceAfter=16)
        body_style = ParagraphStyle("HexaBody", parent=styles["Normal"],
                                    fontSize=11, leading=16, spaceAfter=8)
        heading_style = ParagraphStyle("HexaH1", parent=styles["Heading1"],
                                       textColor=teal, fontSize=14, spaceAfter=8)

        story = [Paragraph(title, title_style), Spacer(1, 6*mm)]

        for line in content.split("\n"):
            stripped = line.strip()
            if not stripped:
                story.append(Spacer(1, 3*mm))
            elif stripped.startswith("## "):
                story.append(Paragraph(stripped[3:], heading_style))
            elif stripped.startswith("# "):
                story.append(Paragraph(stripped[2:], heading_style))
            else:
                # Escape HTML special chars for reportlab
                safe = stripped.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                story.append(Paragraph(safe, body_style))

        doc.build(story)
        logger.info(f"PDF deliverable saved: {path}")
        return fname

    except ImportError:
        logger.warning("reportlab not installed — falling back to text-based PDF via fpdf2")

    try:
        from fpdf import FPDF

        fname = _safe_filename(title, "pdf")
        path = os.path.join(DELIVERABLES_DIR, fname)

        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font("Helvetica", "B", 22)
        pdf.set_text_color(0, 95, 115)
        pdf.cell(0, 12, title, ln=True, align="C")
        pdf.ln(6)
        pdf.set_font("Helvetica", size=11)
        pdf.set_text_color(30, 30, 30)

        for line in content.split("\n"):
            stripped = line.strip()
            if not stripped:
                pdf.ln(4)
            elif stripped.startswith("## ") or stripped.startswith("# "):
                pdf.set_font("Helvetica", "B", 13)
                pdf.set_text_color(0, 95, 115)
                pdf.multi_cell(0, 8, stripped.lstrip("#").strip())
                pdf.set_font("Helvetica", size=11)
                pdf.set_text_color(30, 30, 30)
            else:
                pdf.multi_cell(0, 7, stripped)

        pdf.output(path)
        logger.info(f"PDF deliverable (fpdf2) saved: {path}")
        return fname

    except ImportError:
        raise RuntimeError(
            "No PDF library available. Install either 'reportlab' or 'fpdf2':\n"
            "  pip install reportlab\n  or\n  pip install fpdf2"
        )


# ---------------------------------------------------------------------------
# Router setup
# ---------------------------------------------------------------------------

def setup_deliverable_routes() -> APIRouter:
    router = APIRouter(prefix="/api/deliverable", tags=["deliverable"])

    # ── Word ──────────────────────────────────────────────────────────────
    @router.post("/docx")
    async def create_docx(request: Request):
        """
        Generate a Word (.docx) document.

        Body JSON:
          title    : str                          (required)
          sections : [{"title": str, "body": str}]  (required, at least 1)
          tables   : [[[cell, ...], ...]]         (optional)
        """
        try:
            body = await request.json()
        except Exception:
            return JSONResponse({"error": "Invalid JSON body"}, status_code=400)

        title = str(body.get("title") or "Untitled Document").strip()
        sections = body.get("sections")
        if not sections or not isinstance(sections, list):
            return JSONResponse({"error": "'sections' is required and must be a list"}, status_code=400)

        tables = body.get("tables")

        try:
            fname = _build_docx(title, sections, tables)
        except ImportError as e:
            return JSONResponse({"error": f"python-docx not installed: {e}. Run: pip install python-docx"}, status_code=500)
        except Exception as e:
            logger.exception("docx generation failed")
            return JSONResponse({"error": str(e)}, status_code=500)

        return JSONResponse({
            "status": "ok",
            "filename": fname,
            "download_url": f"/api/deliverable/download/{fname}",
            "message": f"Word document '{title}' created successfully.",
        })

    # ── Excel ─────────────────────────────────────────────────────────────
    @router.post("/xlsx")
    async def create_xlsx(request: Request):
        """
        Generate an Excel (.xlsx) spreadsheet.

        Body JSON:
          title  : str
          sheets : {"Sheet Name": [[header1, header2, ...], [row1col1, row1col2, ...]]}
        """
        try:
            body = await request.json()
        except Exception:
            return JSONResponse({"error": "Invalid JSON body"}, status_code=400)

        title = str(body.get("title") or "Spreadsheet").strip()
        sheets = body.get("sheets")
        if not sheets or not isinstance(sheets, dict):
            return JSONResponse({"error": "'sheets' is required and must be an object"}, status_code=400)

        try:
            fname = _build_xlsx(title, sheets)
        except ImportError as e:
            return JSONResponse({"error": f"openpyxl not installed: {e}. Run: pip install openpyxl"}, status_code=500)
        except Exception as e:
            logger.exception("xlsx generation failed")
            return JSONResponse({"error": str(e)}, status_code=500)

        return JSONResponse({
            "status": "ok",
            "filename": fname,
            "download_url": f"/api/deliverable/download/{fname}",
            "message": f"Excel spreadsheet '{title}' created successfully.",
        })

    # ── PowerPoint ────────────────────────────────────────────────────────
    @router.post("/pptx")
    async def create_pptx(request: Request):
        """
        Generate a PowerPoint (.pptx) presentation.

        Body JSON:
          title  : str
          slides : [{"title": str, "points": [str, ...]}]
        """
        try:
            body = await request.json()
        except Exception:
            return JSONResponse({"error": "Invalid JSON body"}, status_code=400)

        title = str(body.get("title") or "Presentation").strip()
        slides = body.get("slides")
        if not slides or not isinstance(slides, list):
            return JSONResponse({"error": "'slides' is required and must be a list"}, status_code=400)

        try:
            fname = _build_pptx(title, slides)
        except ImportError as e:
            return JSONResponse({"error": f"python-pptx not installed: {e}. Run: pip install python-pptx"}, status_code=500)
        except Exception as e:
            logger.exception("pptx generation failed")
            return JSONResponse({"error": str(e)}, status_code=500)

        return JSONResponse({
            "status": "ok",
            "filename": fname,
            "download_url": f"/api/deliverable/download/{fname}",
            "message": f"PowerPoint '{title}' created successfully.",
        })

    # ── PDF ───────────────────────────────────────────────────────────────
    @router.post("/pdf")
    async def create_pdf(request: Request):
        """
        Generate a PDF from markdown/plain-text content.

        Body JSON:
          title   : str
          content : str  (markdown or plain text)
        """
        try:
            body = await request.json()
        except Exception:
            return JSONResponse({"error": "Invalid JSON body"}, status_code=400)

        title = str(body.get("title") or "Document").strip()
        content = str(body.get("content") or "").strip()
        if not content:
            return JSONResponse({"error": "'content' is required"}, status_code=400)

        try:
            fname = _build_pdf(title, content)
        except RuntimeError as e:
            return JSONResponse({"error": str(e)}, status_code=500)
        except Exception as e:
            logger.exception("pdf generation failed")
            return JSONResponse({"error": str(e)}, status_code=500)

        return JSONResponse({
            "status": "ok",
            "filename": fname,
            "download_url": f"/api/deliverable/download/{fname}",
            "message": f"PDF '{title}' created successfully.",
        })

    # ── List files ────────────────────────────────────────────────────────
    @router.get("/list")
    async def list_deliverables():
        """List all generated deliverable files."""
        try:
            files = []
            for f in sorted(os.listdir(DELIVERABLES_DIR), reverse=True):
                if f.startswith("."):
                    continue
                fpath = os.path.join(DELIVERABLES_DIR, f)
                ext = os.path.splitext(f)[1].lower()
                files.append({
                    "filename": f,
                    "size": os.path.getsize(fpath),
                    "type": ext.lstrip("."),
                    "download_url": f"/api/deliverable/download/{f}",
                    "modified": os.path.getmtime(fpath),
                })
            return JSONResponse({"files": files, "count": len(files)})
        except Exception as e:
            logger.exception("list_deliverables failed")
            return JSONResponse({"error": str(e)}, status_code=500)

    # ── Download ──────────────────────────────────────────────────────────
    @router.get("/download/{filename}")
    async def download_deliverable(filename: str):
        """Download a generated deliverable file."""
        # Security: no path traversal
        safe_name = os.path.basename(filename)
        fpath = os.path.join(DELIVERABLES_DIR, safe_name)

        if not os.path.isfile(fpath):
            return JSONResponse({"error": "File not found"}, status_code=404)

        ext = os.path.splitext(safe_name)[1].lower()
        mime_map = {
            ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            ".pdf":  "application/pdf",
        }
        media_type = mime_map.get(ext, "application/octet-stream")

        return FileResponse(
            fpath,
            media_type=media_type,
            filename=safe_name,
            headers={"Content-Disposition": f'attachment; filename="{safe_name}"'},
        )

    return router
