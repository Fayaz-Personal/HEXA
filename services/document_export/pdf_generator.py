# services/document_export/pdf_generator.py
"""
PDF Generator for HEXA Enterprise Document Export System.
Uses ReportLab to create professional, publication-quality PDFs.
"""

import os
from datetime import datetime
from typing import List

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    HRFlowable,
    KeepTogether,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from services.document_export.models import DocumentModel, PDFExportOptions


class NumberedCanvas(canvas.Canvas):
    """
    Two-pass canvas to dynamically compute and draw total page numbers ('Page X of Y'),
    top header branding, and bottom confidentiality footer.
    """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, page_count: int):
        self.saveState()

        teal = colors.HexColor("#0a9396")
        dark_navy = colors.HexColor("#005f73")
        gray = colors.HexColor("#6c757d")

        # Top Header (pages > 1 or cover option)
        self.setFont("Helvetica-Bold", 8)
        self.setFillColor(teal)
        self.drawString(25 * mm, 285 * mm, "HEXA AI")
        self.setFont("Helvetica", 8)
        self.setFillColor(gray)
        self.drawString(42 * mm, 285 * mm, "|   Sovereign Industrial AI Workbench")

        self.setStrokeColor(colors.HexColor("#e2e8f0"))
        self.setLineWidth(0.5)
        self.line(25 * mm, 282 * mm, 185 * mm, 282 * mm)

        # Bottom Footer
        self.line(25 * mm, 20 * mm, 185 * mm, 20 * mm)

        self.setFont("Helvetica", 8)
        self.setFillColor(gray)
        self.drawString(25 * mm, 14 * mm, "CONFIDENTIAL & PROPRIETARY — LOCAL AIR-GAPPED WORKSPACE")

        page_str = f"Page {self._pageNumber} of {page_count}"
        self.drawRightString(185 * mm, 14 * mm, page_str)

        self.restoreState()


class PDFGenerator:
    """Generates professional PDFs from DocumentModel."""

    @classmethod
    def generate(cls, doc: DocumentModel, options: PDFExportOptions, output_path: str) -> str:
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

        pdf_doc = SimpleDocTemplate(
            output_path,
            pagesize=A4,
            leftMargin=25 * mm,
            rightMargin=25 * mm,
            topMargin=25 * mm,
            bottomMargin=25 * mm,
        )

        styles = getSampleStyleSheet()

        teal = colors.HexColor("#0a9396")
        dark_navy = colors.HexColor("#005f73")
        charcoal = colors.HexColor("#2b2d42")
        light_bg = colors.HexColor("#f8fafc")
        border_color = colors.HexColor("#cbd5e1")

        # Custom Paragraph Styles
        title_style = ParagraphStyle(
            "HexaTitle",
            parent=styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=22,
            leading=26,
            textColor=dark_navy,
            alignment=0,
            spaceAfter=6,
        )

        meta_style = ParagraphStyle(
            "HexaMeta",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=9,
            leading=12,
            textColor=colors.HexColor("#64748b"),
            spaceAfter=15,
        )

        h1_style = ParagraphStyle(
            "HexaH1",
            parent=styles["Heading1"],
            fontName="Helvetica-Bold",
            fontSize=15,
            leading=18,
            textColor=teal,
            spaceBefore=14,
            spaceAfter=8,
            keepWithNext=True,
        )

        h2_style = ParagraphStyle(
            "HexaH2",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=12,
            leading=15,
            textColor=dark_navy,
            spaceBefore=10,
            spaceAfter=6,
            keepWithNext=True,
        )

        body_style = ParagraphStyle(
            "HexaBody",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=10,
            leading=14,
            textColor=charcoal,
            spaceAfter=6,
        )

        bullet_style = ParagraphStyle(
            "HexaBullet",
            parent=body_style,
            leftIndent=15,
            firstLineIndent=-10,
            spaceAfter=3,
        )

        summary_box_style = ParagraphStyle(
            "HexaSummary",
            parent=styles["Normal"],
            fontName="Helvetica-Oblique",
            fontSize=10,
            leading=15,
            textColor=colors.HexColor("#1e293b"),
        )

        code_style = ParagraphStyle(
            "HexaCode",
            parent=styles["Code"],
            fontName="Courier",
            fontSize=8.5,
            leading=11,
            textColor=colors.HexColor("#0f172a"),
            spaceAfter=4,
        )

        story = []

        # Title & Metadata
        story.append(Paragraph(cls._escape(doc.title), title_style))

        meta_text = f"<b>Author:</b> {cls._escape(doc.author)} &nbsp;&nbsp;|&nbsp;&nbsp; <b>Source:</b> {cls._escape(doc.source_type)}"
        if options.include_datetime:
            meta_text += f" &nbsp;&nbsp;|&nbsp;&nbsp; <b>Date:</b> {doc.created_at}"
        story.append(Paragraph(meta_text, meta_style))
        story.append(HRFlowable(width="100%", thickness=1, color=teal, spaceBefore=0, spaceAfter=14))

        # Executive Summary Box
        if doc.summary:
            summary_p = Paragraph(f"<b>Executive Summary:</b><br/>{cls._escape(doc.summary)}", summary_box_style)
            summary_table = Table([[summary_p]], colWidths=[160 * mm])
            summary_table.setStyle(
                TableStyle([
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f0fdfa")),
                    ("BOX", (0, 0), (-1, -1), 1, teal),
                    ("PADDING", (0, 0), (-1, -1), 10),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ])
            )
            story.append(summary_table)
            story.append(Spacer(1, 12))

        # Sections
        for section in doc.sections:
            if section.title:
                h_style = h1_style if section.level <= 1 else h2_style
                story.append(Paragraph(cls._escape(section.title), h_style))

            # Paragraphs
            for p in section.paragraphs:
                story.append(Paragraph(cls._escape(p), body_style))

            # Bullet Lists
            for b_list in section.bullet_lists:
                for item in b_list:
                    story.append(Paragraph(f"• {cls._escape(item)}", bullet_style))
                story.append(Spacer(1, 4))

            # Numbered Lists
            for n_list in section.numbered_lists:
                for n_idx, item in enumerate(n_list, 1):
                    story.append(Paragraph(f"<b>{n_idx}.</b> {cls._escape(item)}", bullet_style))
                story.append(Spacer(1, 4))

            # Code Blocks
            for cb in section.code_blocks:
                code_p = Paragraph(cls._escape(cb.code).replace("\n", "<br/>").replace(" ", "&nbsp;"), code_style)
                code_table = Table([[code_p]], colWidths=[160 * mm])
                code_table.setStyle(
                    TableStyle([
                        ("BACKGROUND", (0, 0), (-1, -1), light_bg),
                        ("BOX", (0, 0), (-1, -1), 0.5, border_color),
                        ("PADDING", (0, 0), (-1, -1), 8),
                    ])
                )
                story.append(code_table)
                story.append(Spacer(1, 8))

            # Tables
            for tbl in section.tables:
                table_flowable = cls._build_table_flowable(tbl, 160 * mm, body_style)
                if table_flowable:
                    story.append(table_flowable)
                    story.append(Spacer(1, 10))

        # Standalone Global Tables (not in sections)
        for tbl in doc.tables:
            # Avoid duplicate table rendering
            if not any(tbl in sec.tables for sec in doc.sections):
                table_flowable = cls._build_table_flowable(tbl, 160 * mm, body_style)
                if table_flowable:
                    story.append(table_flowable)
                    story.append(Spacer(1, 10))

        # Source References Section
        if options.include_sources and doc.references:
            story.append(Spacer(1, 10))
            story.append(Paragraph("References & Sources", h1_style))
            story.append(HRFlowable(width="100%", thickness=0.5, color=border_color, spaceBefore=2, spaceAfter=8))
            for r_idx, ref in enumerate(doc.references, 1):
                ref_text = f"<b>[{r_idx}]</b> {cls._escape(ref.source_name)}"
                if ref.page:
                    ref_text += f" (Page {cls._escape(ref.page)})"
                if ref.url:
                    ref_text += f" - <i>{cls._escape(ref.url)}</i>"
                if ref.citation:
                    ref_text += f"<br/>&nbsp;&nbsp;&nbsp;&nbsp;{cls._escape(ref.citation)}"
                story.append(Paragraph(ref_text, body_style))

        canvas_maker = NumberedCanvas if options.include_page_numbers else canvas.Canvas
        pdf_doc.build(story, canvasmaker=canvas_maker)
        return output_path

    @classmethod
    def _build_table_flowable(cls, tbl, max_width_pt: float, base_style: ParagraphStyle):
        if not tbl.headers and not tbl.rows:
            return None

        all_rows = []
        if tbl.headers:
            all_rows.append(tbl.headers)
        all_rows.extend(tbl.rows)

        if not all_rows:
            return None

        num_cols = max(len(r) for r in all_rows)
        col_width = max_width_pt / num_cols if num_cols > 0 else max_width_pt

        header_style = ParagraphStyle(
            "TableHeader",
            parent=base_style,
            fontName="Helvetica-Bold",
            fontSize=9,
            leading=11,
            textColor=colors.white,
            alignment=1,
        )

        cell_style = ParagraphStyle(
            "TableCell",
            parent=base_style,
            fontName="Helvetica",
            fontSize=8.5,
            leading=11,
            textColor=colors.HexColor("#1e293b"),
        )

        formatted_data = []
        for r_idx, row in enumerate(all_rows):
            row_data = []
            is_header = (r_idx == 0 and bool(tbl.headers))
            p_style = header_style if is_header else cell_style

            for c_idx in range(num_cols):
                val = str(row[c_idx]) if c_idx < len(row) else ""
                row_data.append(Paragraph(cls._escape(val), p_style))
            formatted_data.append(row_data)

        table = Table(formatted_data, colWidths=[col_width] * num_cols)

        t_style = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0a9396")),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
            ("PADDING", (0, 0), (-1, -1), 5),
        ]

        # Alternating row background colors
        for r in range(1, len(formatted_data)):
            if r % 2 == 0:
                t_style.append(("BACKGROUND", (0, r), (-1, r), colors.HexColor("#f8fafc")))

        table.setStyle(TableStyle(t_style))
        return table

    @classmethod
    def _escape(cls, text: str) -> str:
        if not text:
            return ""
        return (
            str(text)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
            .replace("\n", "<br/>")
        )
