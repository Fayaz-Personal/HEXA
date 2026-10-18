# services/document_export/parser.py
"""
AI Content & Markdown Parser for HEXA Enterprise Document Export System.
Converts Markdown, HTML, and structured data into normalized DocumentModel.
"""

import re
from typing import Any, Dict, List, Optional
from services.document_export.models import (
    DocumentCodeBlock,
    DocumentModel,
    DocumentReference,
    DocumentSection,
    DocumentTable,
)


class ContentParser:
    """Parses raw AI markdown text and structured data into a DocumentModel."""

    @classmethod
    def parse(
        cls,
        content: str,
        title: Optional[str] = None,
        author: str = "HEXA AI",
        source_type: str = "AI Response",
        structured_data: Optional[Dict[str, Any]] = None,
    ) -> DocumentModel:
        content = content or ""
        lines = content.split("\n")

        extracted_title = None
        executive_summary = ""
        sections: List[DocumentSection] = []
        global_tables: List[DocumentTable] = []
        global_code_blocks: List[DocumentCodeBlock] = []
        references: List[DocumentReference] = []

        current_section = DocumentSection(title="Overview", level=1)
        in_code_block = False
        code_lang = "text"
        code_lines: List[str] = []
        in_table = False
        table_lines: List[str] = []
        current_bullet_list: List[str] = []
        current_numbered_list: List[str] = []

        def flush_lists():
            nonlocal current_bullet_list, current_numbered_list
            if current_bullet_list:
                current_section.bullet_lists.append(list(current_bullet_list))
                current_bullet_list = []
            if current_numbered_list:
                current_section.numbered_lists.append(list(current_numbered_list))
                current_numbered_list = []

        def flush_table():
            nonlocal in_table, table_lines
            if not in_table or not table_lines:
                in_table = False
                table_lines = []
                return

            tbl = cls._parse_markdown_table(table_lines)
            if tbl:
                current_section.tables.append(tbl)
                global_tables.append(tbl)
            in_table = False
            table_lines = []

        idx = 0
        while idx < len(lines):
            line = lines[idx]
            stripped = line.strip()

            # 1. Code Block Parsing
            if stripped.startswith("```"):
                if in_code_block:
                    # End code block
                    code_str = "\n".join(code_lines)
                    block = DocumentCodeBlock(language=code_lang, code=code_str)
                    current_section.code_blocks.append(block)
                    global_code_blocks.append(block)
                    in_code_block = False
                    code_lines = []
                    code_lang = "text"
                else:
                    flush_lists()
                    flush_table()
                    in_code_block = True
                    code_lang = stripped[3:].strip() or "text"
                    code_lines = []
                idx += 1
                continue

            if in_code_block:
                code_lines.append(line)
                idx += 1
                continue

            # 2. Table Parsing
            if "|" in line and (line.count("|") >= 2):
                flush_lists()
                in_table = True
                table_lines.append(line)
                idx += 1
                continue
            else:
                if in_table:
                    flush_table()

            # Empty lines
            if not stripped:
                flush_lists()
                idx += 1
                continue

            # 3. Headings
            heading_match = re.match(r"^(#{1,6})\s+(.+)$", stripped)
            if heading_match:
                flush_lists()
                flush_table()
                hashes, heading_text = heading_match.groups()
                level = len(hashes)
                heading_clean = heading_text.strip()

                if level == 1 and not extracted_title:
                    extracted_title = heading_clean

                # Check if it's executive summary section
                if heading_clean.lower() in ["executive summary", "summary"]:
                    current_section = DocumentSection(title=heading_clean, level=level)
                    sections.append(current_section)
                elif heading_clean.lower() in ["references", "sources", "citations"]:
                    current_section = DocumentSection(title=heading_clean, level=level)
                    sections.append(current_section)
                else:
                    if current_section.paragraphs or current_section.bullet_lists or current_section.tables or current_section.code_blocks:
                        pass
                    current_section = DocumentSection(title=heading_clean, level=level)
                    sections.append(current_section)

                idx += 1
                continue

            # 4. Bullet lists
            bullet_match = re.match(r"^[\*\-\+]\s+(.+)$", stripped)
            if bullet_match:
                if current_numbered_list:
                    flush_lists()
                current_bullet_list.append(bullet_match.group(1).strip())
                idx += 1
                continue

            # 5. Numbered lists
            num_match = re.match(r"^\d+[\.\)]\s+(.+)$", stripped)
            if num_match:
                if current_bullet_list:
                    flush_lists()
                current_numbered_list.append(num_match.group(1).strip())
                idx += 1
                continue

            # 6. References / Sources check in text
            ref_match = re.match(r"^(?:Source|Reference|\d+[\.\)])\s*:\s*(.+)$", stripped, re.IGNORECASE)
            if ref_match and current_section.title.lower() in ["references", "sources"]:
                references.append(DocumentReference(source_name=ref_match.group(1).strip()))
                idx += 1
                continue

            # Default paragraph
            flush_lists()
            current_section.paragraphs.append(stripped)
            idx += 1

        # Flush any remaining buffers
        flush_lists()
        flush_table()

        if current_section not in sections and (
            current_section.paragraphs or current_section.bullet_lists or current_section.tables or current_section.code_blocks
        ):
            sections.append(current_section)

        # Parse structured data if provided
        if structured_data and isinstance(structured_data, dict):
            cls._merge_structured_data(structured_data, sections, global_tables, references)

        # Extract executive summary
        for sec in sections:
            if sec.title.lower() in ["executive summary", "summary"] and sec.paragraphs:
                executive_summary = " ".join(sec.paragraphs)
                break
        if not executive_summary and sections and sections[0].paragraphs:
            executive_summary = sections[0].paragraphs[0]

        # Extract references from reference sections
        for sec in sections:
            if sec.title.lower() in ["references", "sources", "citations"]:
                for bullet in sec.bullet_lists:
                    for item in bullet:
                        references.append(DocumentReference(source_name=item))
                for num_list in sec.numbered_lists:
                    for item in num_list:
                        references.append(DocumentReference(source_name=item))
                for p in sec.paragraphs:
                    references.append(DocumentReference(source_name=p))

        final_title = title if (title and title.strip() and title != "HEXA AI Generated Report") else (extracted_title or title or "HEXA AI Generated Report")

        return DocumentModel(
            title=final_title,
            author=author or "HEXA AI",
            source_type=source_type or "AI Response",
            summary=executive_summary[:500] if executive_summary else "",
            sections=sections,
            tables=global_tables,
            code_blocks=global_code_blocks,
            references=references,
        )

    @classmethod
    def _parse_markdown_table(cls, lines: List[str]) -> Optional[DocumentTable]:
        if len(lines) < 2:
            return None

        parsed_rows: List[List[str]] = []
        for line in lines:
            # Skip delimiter line (| --- | --- |)
            if re.match(r"^\|?\s*:?-+:?\s*\|", line):
                continue
            cells = [c.strip() for c in line.split("|")]
            # Remove leading/trailing empty cells if line started/ended with |
            if line.strip().startswith("|") and len(cells) > 0 and cells[0] == "":
                cells.pop(0)
            if line.strip().endswith("|") and len(cells) > 0 and cells[-1] == "":
                cells.pop()
            if cells:
                parsed_rows.append(cells)

        if not parsed_rows:
            return None

        headers = parsed_rows[0]
        rows = parsed_rows[1:] if len(parsed_rows) > 1 else []
        return DocumentTable(headers=headers, rows=rows)

    @classmethod
    def _merge_structured_data(
        cls,
        data: Dict[str, Any],
        sections: List[DocumentSection],
        global_tables: List[DocumentTable],
        references: List[DocumentReference],
    ):
        """Merges extra structured JSON / table inputs into sections and tables."""
        if "table" in data or "tables" in data:
            tbls = data.get("tables") or [data.get("table")]
            for t in tbls:
                if isinstance(t, dict) and "headers" in t and "rows" in t:
                    tbl = DocumentTable(headers=t["headers"], rows=t["rows"], caption=t.get("caption"))
                    global_tables.append(tbl)
                    if sections:
                        sections[0].tables.append(tbl)
                elif isinstance(t, list) and len(t) > 0:
                    headers = [str(x) for x in t[0]]
                    rows = [[str(cell) for cell in r] for r in t[1:]]
                    tbl = DocumentTable(headers=headers, rows=rows)
                    global_tables.append(tbl)
                    if sections:
                        sections[0].tables.append(tbl)

        if "sources" in data or "references" in data:
            refs = data.get("sources") or data.get("references") or []
            for r in refs:
                if isinstance(r, str):
                    references.append(DocumentReference(source_name=r))
                elif isinstance(r, dict):
                    references.append(
                        DocumentReference(
                            source_name=r.get("name") or r.get("title") or "Source",
                            page=str(r.get("page")) if r.get("page") else None,
                            url=r.get("url"),
                            citation=r.get("citation"),
                        )
                    )
