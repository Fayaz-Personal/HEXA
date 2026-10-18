# services/document_export/docx_generator.py
"""
Word (.docx) Generator for HEXA Enterprise Document Export System.
Uses python-docx to generate structured, styled Word documents.
"""

import os
from typing import List
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor

from services.document_export.models import DocumentModel, WordExportOptions


class DocxGenerator:
    """Generates Word (.docx) documents from DocumentModel."""

    TEAL_RGB = RGBColor(0, 147, 150)       # #0a9396
    NAVY_RGB = RGBColor(0, 95, 115)        # #005f73
    CHARCOAL_RGB = RGBColor(43, 45, 66)    # #2b2d42
    GRAY_RGB = RGBColor(108, 117, 125)     # #6c757d

    @classmethod
    def generate(cls, doc_model: DocumentModel, options: WordExportOptions, output_path: str) -> str:
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

        doc = docx.Document()

        # Page setup - 1 inch margins
        for section in doc.sections:
            section.top_margin = Pt(72)
            section.bottom_margin = Pt(72)
            section.left_margin = Pt(72)
            section.right_margin = Pt(72)

            if options.include_headers_footers:
                # Header
                header = section.header
                hp = header.paragraphs[0]
                hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                hrun = hp.add_run("HEXA Sovereign Industrial AI Workbench")
                hrun.font.name = "Calibri"
                hrun.font.size = Pt(8.5)
                hrun.font.color.rgb = cls.TEAL_RGB

                # Footer
                footer = section.footer
                fp = footer.paragraphs[0]
                fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
                frun = fp.add_run("CONFIDENTIAL & PROPRIETARY — LOCAL AIR-GAPPED WORKSPACE")
                frun.font.name = "Calibri"
                frun.font.size = Pt(8.5)
                frun.font.color.rgb = cls.GRAY_RGB

        # Title Page
        if options.include_title_page:
            cls._add_title_page(doc, doc_model)
        else:
            # Inline Title
            title_p = doc.add_paragraph()
            run = title_p.add_run(doc_model.title)
            run.font.name = "Calibri"
            run.font.size = Pt(24)
            run.font.bold = True
            run.font.color.rgb = cls.NAVY_RGB

            meta_p = doc.add_paragraph()
            meta_run = meta_p.add_run(f"Author: {doc_model.author} | Date: {doc_model.created_at} | Source: {doc_model.source_type}")
            meta_run.font.name = "Calibri"
            meta_run.font.size = Pt(9.5)
            meta_run.font.color.rgb = cls.GRAY_RGB
            doc.add_paragraph()  # spacer

        # Executive Summary
        if doc_model.summary:
            cls._add_executive_summary_box(doc, doc_model.summary)

        # Table of Contents
        if options.include_toc:
            cls._add_toc(doc)

        # Sections
        for sec in doc_model.sections:
            if sec.title:
                h_level = min(max(sec.level, 1), 3)
                hp = doc.add_paragraph()
                hp.paragraph_format.space_before = Pt(14)
                hp.paragraph_format.space_after = Pt(6)
                hp.paragraph_format.keep_with_next = True

                hrun = hp.add_run(sec.title)
                hrun.font.name = "Calibri"
                hrun.font.bold = True

                if h_level == 1:
                    hrun.font.size = Pt(16)
                    hrun.font.color.rgb = cls.TEAL_RGB
                elif h_level == 2:
                    hrun.font.size = Pt(13)
                    hrun.font.color.rgb = cls.NAVY_RGB
                else:
                    hrun.font.size = Pt(11)
                    hrun.font.color.rgb = cls.CHARCOAL_RGB

            # Paragraphs
            for p_text in sec.paragraphs:
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(6)
                p.paragraph_format.line_spacing = 1.15
                run = p.add_run(p_text)
                run.font.name = "Calibri"
                run.font.size = Pt(10.5)
                run.font.color.rgb = cls.CHARCOAL_RGB

            # Bullet Lists
            for b_list in sec.bullet_lists:
                for item in b_list:
                    p = doc.add_paragraph(style="List Bullet")
                    p.paragraph_format.space_after = Pt(3)
                    run = p.add_run(item)
                    run.font.name = "Calibri"
                    run.font.size = Pt(10.5)

            # Numbered Lists
            for n_list in sec.numbered_lists:
                for item in n_list:
                    p = doc.add_paragraph(style="List Number")
                    p.paragraph_format.space_after = Pt(3)
                    run = p.add_run(item)
                    run.font.name = "Calibri"
                    run.font.size = Pt(10.5)

            # Code Blocks
            for cb in sec.code_blocks:
                cls._add_code_block(doc, cb.code, cb.language)

            # Tables
            for tbl in sec.tables:
                cls._add_table(doc, tbl)

        # Standalone Global Tables
        for tbl in doc_model.tables:
            if not any(tbl in sec.tables for sec in doc_model.sections):
                cls._add_table(doc, tbl)

        # References Section
        if options.include_sources and doc_model.references:
            ref_hp = doc.add_paragraph()
            ref_hp.paragraph_format.space_before = Pt(18)
            ref_hp.paragraph_format.space_after = Pt(6)
            ref_run = ref_hp.add_run("References & Sources")
            ref_run.font.name = "Calibri"
            ref_run.font.size = Pt(14)
            ref_run.font.bold = True
            ref_run.font.color.rgb = cls.TEAL_RGB

            for r_idx, ref in enumerate(doc_model.references, 1):
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(4)
                r_text = f"[{r_idx}] {ref.source_name}"
                if ref.page:
                    r_text += f" (Page {ref.page})"
                if ref.url:
                    r_text += f" - {ref.url}"
                run = p.add_run(r_text)
                run.font.name = "Calibri"
                run.font.size = Pt(9.5)
                run.font.color.rgb = cls.CHARCOAL_RGB

        doc.save(output_path)
        return output_path

    @classmethod
    def _add_title_page(cls, doc: docx.Document, doc_model: DocumentModel):
        # Top spacing
        p_top = doc.add_paragraph()
        p_top.paragraph_format.space_before = Pt(72)

        # Brand header
        p_brand = doc.add_paragraph()
        r_brand = p_brand.add_run("HEXA SOVEREIGN INDUSTRIAL AI WORKBENCH")
        r_brand.font.name = "Calibri"
        r_brand.font.size = Pt(10)
        r_brand.font.bold = True
        r_brand.font.color.rgb = cls.TEAL_RGB

        # Document Title
        p_title = doc.add_paragraph()
        p_title.paragraph_format.space_before = Pt(12)
        p_title.paragraph_format.space_after = Pt(18)
        r_title = p_title.add_run(doc_model.title)
        r_title.font.name = "Calibri"
        r_title.font.size = Pt(28)
        r_title.font.bold = True
        r_title.font.color.rgb = cls.NAVY_RGB

        # Metadata
        p_meta = doc.add_paragraph()
        p_meta.paragraph_format.space_after = Pt(24)
        r_meta = p_meta.add_run(f"Generated by {doc_model.author}\nDate: {doc_model.created_at}\nSource Workflow: {doc_model.source_type}")
        r_meta.font.name = "Calibri"
        r_meta.font.size = Pt(11)
        r_meta.font.color.rgb = cls.GRAY_RGB

        doc.add_page_break()

    @classmethod
    def _add_executive_summary_box(cls, doc: docx.Document, summary: str):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)

        shading = parse_xml(r'<w:shd {} w:fill="F0FDFA"/>'.format(nsdecls('w')))
        cell._tc.get_or_add_tcPr().append(shading)

        borders = parse_xml(
            r'<w:tcBorders {}><w:left w:val="single" w:sz="24" w:space="0" w:color="0A9396"/>'
            r'<w:top w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/></w:tcBorders>'.format(nsdecls('w'))
        )
        cell._tc.get_or_add_tcPr().append(borders)

        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)
        r_head = p.add_run("Executive Summary:\n")
        r_head.font.name = "Calibri"
        r_head.font.bold = True
        r_head.font.size = Pt(11)
        r_head.font.color.rgb = cls.TEAL_RGB

        r_body = p.add_run(summary)
        r_body.font.name = "Calibri"
        r_body.font.italic = True
        r_body.font.size = Pt(10.5)

        doc.add_paragraph()  # spacer

    @classmethod
    def _add_toc(cls, doc: docx.Document):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run("Table of Contents")
        run.font.name = "Calibri"
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = cls.TEAL_RGB

        # TOC XML element for Word
        p_toc = doc.add_paragraph()
        run_toc = p_toc.add_run()
        fldChar1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w'))
        instrText = parse_xml(r'<w:instrText %s xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText>' % nsdecls('w'))
        fldChar2 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w'))
        fldChar3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w'))
        run_toc._r.append(fldChar1)
        run_toc._r.append(instrText)
        run_toc._r.append(fldChar2)
        run_toc._r.append(fldChar3)

        doc.add_paragraph()  # spacer

    @classmethod
    def _add_table(cls, doc: docx.Document, tbl_model):
        all_rows = []
        if tbl_model.headers:
            all_rows.append(tbl_model.headers)
        all_rows.extend(tbl_model.rows)

        if not all_rows:
            return

        num_rows = len(all_rows)
        num_cols = max(len(r) for r in all_rows)

        tbl = doc.add_table(rows=num_rows, cols=num_cols)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.style = "Table Grid"

        for r_idx, row_data in enumerate(all_rows):
            is_header = (r_idx == 0 and bool(tbl_model.headers))
            row = tbl.rows[r_idx]

            for c_idx in range(num_cols):
                val = str(row_data[c_idx]) if c_idx < len(row_data) else ""
                cell = row.cells[c_idx]
                cell.text = val

                p = cell.paragraphs[0]
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.space_before = Pt(2)

                if is_header:
                    shd = parse_xml(r'<w:shd {} w:fill="0A9396"/>'.format(nsdecls('w')))
                    cell._tc.get_or_add_tcPr().append(shd)
                    for run in p.runs:
                        run.font.name = "Calibri"
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(255, 255, 255)
                        run.font.size = Pt(9.5)
                else:
                    if r_idx % 2 == 0:
                        shd = parse_xml(r'<w:shd {} w:fill="F8FAFC"/>'.format(nsdecls('w')))
                        cell._tc.get_or_add_tcPr().append(shd)
                    for run in p.runs:
                        run.font.name = "Calibri"
                        run.font.size = Pt(9.5)

        doc.add_paragraph()  # spacer

    @classmethod
    def _add_code_block(cls, doc: docx.Document, code: str, language: str):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)

        shd = parse_xml(r'<w:shd {} w:fill="F8FAFC"/>'.format(nsdecls('w')))
        cell._tc.get_or_add_tcPr().append(shd)

        borders = parse_xml(
            r'<w:tcBorders {}><w:left w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>'
            r'<w:top w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>'
            r'<w:right w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>'
            r'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/></w:tcBorders>'.format(nsdecls('w'))
        )
        cell._tc.get_or_add_tcPr().append(borders)

        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(code)
        run.font.name = "Courier New"
        run.font.size = Pt(9)

        doc.add_paragraph()  # spacer
