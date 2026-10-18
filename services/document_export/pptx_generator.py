# services/document_export/pptx_generator.py
"""
PowerPoint (.pptx) Generator for HEXA Enterprise Document Export System.
Uses python-pptx to generate clean, executive presentation decks from DocumentModel.
"""

import os
from typing import List, Optional
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from services.document_export.models import DocumentModel, PPTXExportOptions


class PPTXGenerator:
    """Generates PowerPoint presentations from DocumentModel."""

    TEAL_RGB = RGBColor(0, 147, 150)       # #0a9396
    NAVY_RGB = RGBColor(0, 95, 115)        # #005f73
    CHARCOAL_RGB = RGBColor(43, 45, 66)    # #2b2d42
    GRAY_RGB = RGBColor(108, 117, 125)     # #6c757d
    WHITE_RGB = RGBColor(255, 255, 255)

    @classmethod
    def generate(cls, doc_model: DocumentModel, options: PPTXExportOptions, output_path: str) -> str:
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

        prs = Presentation()
        prs.slide_width = Inches(13.333)   # 16:9 widescreen
        prs.slide_height = Inches(7.5)

        blank_layout = prs.slide_layouts[6]

        slide_count = 0
        max_slides = options.max_slides

        # 1. Title Slide
        if options.include_title_slide and slide_count < max_slides:
            cls._create_title_slide(prs, blank_layout, doc_model)
            slide_count += 1

        # 2. Executive Summary Slide
        if options.include_summary_slide and doc_model.summary and slide_count < max_slides:
            cls._create_summary_slide(prs, blank_layout, doc_model.summary)
            slide_count += 1

        # 3. Content Sections & Tables
        for sec in doc_model.sections:
            if slide_count >= max_slides:
                break

            # Section text/bullet points
            bullets = cls._collect_bullets(sec)
            if bullets:
                # Split bullets into chunks of max 5 per slide
                chunk_size = 5
                for i in range(0, len(bullets), chunk_size):
                    if slide_count >= max_slides:
                        break
                    chunk = bullets[i : i + chunk_size]
                    title_suffix = f" ({i // chunk_size + 1})" if i > 0 else ""
                    cls._create_bullet_slide(
                        prs,
                        blank_layout,
                        title=f"{sec.title}{title_suffix}" if sec.title else "Analysis",
                        bullets=chunk,
                    )
                    slide_count += 1

            # Tables in section
            for tbl in sec.tables:
                if slide_count >= max_slides:
                    break
                cls._create_table_slide(
                    prs,
                    blank_layout,
                    title=f"{sec.title} - Data Table" if sec.title else "Data Table",
                    tbl_model=tbl,
                )
                slide_count += 1

        # Standalone Global Tables
        for tbl in doc_model.tables:
            if slide_count >= max_slides:
                break
            if not any(tbl in sec.tables for sec in doc_model.sections):
                cls._create_table_slide(
                    prs,
                    blank_layout,
                    title="Data Analysis Table",
                    tbl_model=tbl,
                )
                slide_count += 1

        # 4. References Slide
        if options.include_source_slide and doc_model.references and slide_count < max_slides:
            cls._create_references_slide(prs, blank_layout, doc_model.references)
            slide_count += 1

        prs.save(output_path)
        return output_path

    @classmethod
    def _create_title_slide(cls, prs: Presentation, layout, doc_model: DocumentModel):
        slide = prs.slides.add_slide(layout)

        # Background banner / card shape
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5)
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(248, 250, 252)
        shape.line.fill.background()

        # Top Accent Line
        accent = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.25)
        )
        accent.fill.solid()
        accent.fill.fore_color.rgb = cls.TEAL_RGB
        accent.line.fill.background()

        # Title Box
        tb = slide.shapes.add_textbox(Inches(1.2), Inches(2.2), Inches(10.9), Inches(3.5))
        tf = tb.text_frame
        tf.word_wrap = True

        p0 = tf.paragraphs[0]
        p0.text = "HEXA AI WORKBENCH"
        p0.font.name = "Arial"
        p0.font.size = Pt(14)
        p0.font.bold = True
        p0.font.color.rgb = cls.TEAL_RGB
        p0.space_after = Pt(14)

        p1 = tf.add_paragraph()
        p1.text = doc_model.title
        p1.font.name = "Arial"
        p1.font.size = Pt(36)
        p1.font.bold = True
        p1.font.color.rgb = cls.NAVY_RGB
        p1.space_after = Pt(20)

        p2 = tf.add_paragraph()
        p2.text = f"Prepared by: {doc_model.author}  |  Date: {doc_model.created_at}  |  Source: {doc_model.source_type}"
        p2.font.name = "Arial"
        p2.font.size = Pt(13)
        p2.font.color.rgb = cls.GRAY_RGB

    @classmethod
    def _create_summary_slide(cls, prs: Presentation, layout, summary: str):
        slide = prs.slides.add_slide(layout)
        cls._add_header_footer(slide, "Executive Summary")

        tb = slide.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.333), Inches(4.8))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = summary
        p.font.name = "Arial"
        p.font.size = Pt(18)
        p.font.color.rgb = cls.CHARCOAL_RGB
        p.line_spacing = 1.3

    @classmethod
    def _create_bullet_slide(cls, prs: Presentation, layout, title: str, bullets: List[str]):
        slide = prs.slides.add_slide(layout)
        cls._add_header_footer(slide, title)

        tb = slide.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.333), Inches(4.8))
        tf = tb.text_frame
        tf.word_wrap = True

        for idx, bullet in enumerate(bullets):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            p.text = f"•  {bullet}"
            p.font.name = "Arial"
            p.font.size = Pt(16)
            p.font.color.rgb = cls.CHARCOAL_RGB
            p.space_after = Pt(14)
            p.line_spacing = 1.25

    @classmethod
    def _create_table_slide(cls, prs: Presentation, layout, title: str, tbl_model):
        slide = prs.slides.add_slide(layout)
        cls._add_header_footer(slide, title)

        all_rows = []
        if tbl_model.headers:
            all_rows.append(tbl_model.headers)
        all_rows.extend(tbl_model.rows)

        if not all_rows:
            return

        # Cap visible rows to 8 per slide
        visible_rows = all_rows[:8]
        rows_count = len(visible_rows)
        cols_count = max(len(r) for r in visible_rows)

        left = Inches(1.0)
        top = Inches(1.8)
        width = Inches(11.333)
        height = Inches(4.8)

        table_shape = slide.shapes.add_table(rows_count, cols_count, left, top, width, height)
        table = table_shape.table

        for r_idx, row_data in enumerate(visible_rows):
            is_header = (r_idx == 0 and bool(tbl_model.headers))
            for c_idx in range(cols_count):
                val = str(row_data[c_idx]) if c_idx < len(row_data) else ""
                cell = table.cell(r_idx, c_idx)
                cell.text = val

                p = cell.text_frame.paragraphs[0]
                p.font.name = "Arial"
                p.font.size = Pt(12 if is_header else 10.5)

                if is_header:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = cls.TEAL_RGB
                    p.font.bold = True
                    p.font.color.rgb = cls.WHITE_RGB
                    p.alignment = PP_ALIGN.CENTER
                else:
                    if r_idx % 2 == 0:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = RGBColor(248, 250, 252)
                    p.font.color.rgb = cls.CHARCOAL_RGB

    @classmethod
    def _create_references_slide(cls, prs: Presentation, layout, references):
        slide = prs.slides.add_slide(layout)
        cls._add_header_footer(slide, "References & Sources")

        tb = slide.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.333), Inches(4.8))
        tf = tb.text_frame
        tf.word_wrap = True

        for idx, ref in enumerate(references[:8]):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            r_str = f"[{idx + 1}] {ref.source_name}"
            if ref.page:
                r_str += f" (Page {ref.page})"
            if ref.url:
                r_str += f" - {ref.url}"
            p.text = r_str
            p.font.name = "Arial"
            p.font.size = Pt(13)
            p.font.color.rgb = cls.CHARCOAL_RGB
            p.space_after = Pt(10)

    @classmethod
    def _add_header_footer(cls, slide, title_text: str):
        # Background
        bg = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5)
        )
        bg.fill.solid()
        bg.fill.fore_color.rgb = RGBColor(255, 255, 255)
        bg.line.fill.background()

        # Header Box
        tb = slide.shapes.add_textbox(Inches(1.0), Inches(0.5), Inches(11.333), Inches(1.0))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = title_text
        p.font.name = "Arial"
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = cls.NAVY_RGB

        # Divider line
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(1.4), Inches(11.333), Inches(0.04)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = cls.TEAL_RGB
        line.line.fill.background()

        # Footer
        ftb = slide.shapes.add_textbox(Inches(1.0), Inches(6.8), Inches(11.333), Inches(0.5))
        fp = ftb.text_frame.paragraphs[0]
        fp.text = "HEXA AI Workbench  |  Confidential & Proprietary"
        fp.font.name = "Arial"
        fp.font.size = Pt(9)
        fp.font.color.rgb = cls.GRAY_RGB

    @classmethod
    def _collect_bullets(cls, section) -> List[str]:
        bullets = []
        for p in section.paragraphs:
            if len(p) > 10:
                bullets.append(p)
        for b_list in section.bullet_lists:
            bullets.extend(b_list)
        for n_list in section.numbered_lists:
            bullets.extend(n_list)
        return bullets
