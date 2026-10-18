# tests/test_export_system.py
"""
Automated unit & API tests for HEXA Enterprise Document Export System.
"""

import os
import pytest
from fastapi.testclient import TestClient

from app import app
from services.document_export import (
    ContentParser,
    DocxGenerator,
    ExportRequest,
    ExportService,
    PDFGenerator,
    PPTXGenerator,
    XLSXGenerator,
)

SAMPLE_MARKDOWN = """# Industrial Equipment Analysis

## Executive Summary

The equipment is operating within normal temperature limits.

## Measurements

| Parameter   | Value | Unit | Status  |
| ----------- | ----: | ---- | ------- |
| Temperature |    72 | °C   | Normal  |
| Pressure    |   5.2 | bar  | Normal  |
| Vibration   |   3.8 | mm/s | Warning |

## Recommendations

* Continue monitoring temperature.
* Inspect vibration levels.
* Schedule preventive maintenance.

## Conclusion

The equipment can continue operating under current conditions.

## References
1. Vibration Sensor Model V3 Spec Sheet
2. Temperature Gauge Log 2026-10-08
"""


def test_content_parser():
    doc = ContentParser.parse(SAMPLE_MARKDOWN, title="Industrial Equipment Analysis")
    assert doc.title == "Industrial Equipment Analysis"
    assert "operating within normal temperature" in doc.summary
    assert len(doc.sections) >= 3
    assert len(doc.tables) >= 1
    assert doc.tables[0].headers == ["Parameter", "Value", "Unit", "Status"]
    assert len(doc.tables[0].rows) == 3
    assert len(doc.references) >= 2


def test_pdf_generation(tmp_path):
    doc = ContentParser.parse(SAMPLE_MARKDOWN, title="Industrial Equipment Analysis")
    pdf_path = str(tmp_path / "test_report.pdf")
    res_path = PDFGenerator.generate(doc, req_options(), pdf_path)
    assert os.path.exists(res_path)
    assert os.path.getsize(res_path) > 1000


def test_docx_generation(tmp_path):
    doc = ContentParser.parse(SAMPLE_MARKDOWN, title="Industrial Equipment Analysis")
    docx_path = str(tmp_path / "test_report.docx")
    res_path = DocxGenerator.generate(doc, word_options(), docx_path)
    assert os.path.exists(res_path)
    assert os.path.getsize(res_path) > 1000


def test_pptx_generation(tmp_path):
    doc = ContentParser.parse(SAMPLE_MARKDOWN, title="Industrial Equipment Analysis")
    pptx_path = str(tmp_path / "test_report.pptx")
    res_path = PPTXGenerator.generate(doc, pptx_options(), pptx_path)
    assert os.path.exists(res_path)
    assert os.path.getsize(res_path) > 1000


def test_xlsx_generation(tmp_path):
    doc = ContentParser.parse(SAMPLE_MARKDOWN, title="Industrial Equipment Analysis")
    xlsx_path = str(tmp_path / "test_report.xlsx")
    res_path = XLSXGenerator.generate(doc, excel_options(), xlsx_path)
    assert os.path.exists(res_path)
    assert os.path.getsize(res_path) > 1000


def test_export_service_all():
    req = ExportRequest(
        title="Industrial Equipment Analysis",
        author="HEXA Test Runner",
        content=SAMPLE_MARKDOWN,
    )
    result = ExportService.export_all(req, owner_id="test_user")
    assert result["status"] == "ok"
    assert result["format"] == "zip"
    assert os.path.exists(os.path.join(result["download_url"])) is False or "export_id" in result
    assert len(result["files"]) == 4


def test_export_api_endpoints():
    client = TestClient(app)

    payload = {
        "title": "Industrial Equipment Analysis API Test",
        "author": "HEXA AI",
        "content": SAMPLE_MARKDOWN,
    }

    # PDF Endpoint
    resp = client.post("/api/export/pdf", json=payload)
    assert resp.status_code == 200
    data = resp.json()
    assert data["status"] == "ok"
    download_url = data["download_url"]

    # Download PDF
    dl_resp = client.get(download_url)
    assert dl_resp.status_code == 200
    assert dl_resp.headers["content-type"] == "application/pdf"

    # DOCX Endpoint
    resp = client.post("/api/export/docx", json=payload)
    assert resp.status_code == 200
    assert resp.json()["status"] == "ok"

    # PPTX Endpoint
    resp = client.post("/api/export/pptx", json=payload)
    assert resp.status_code == 200
    assert resp.json()["status"] == "ok"

    # XLSX Endpoint
    resp = client.post("/api/export/xlsx", json=payload)
    assert resp.status_code == 200
    assert resp.json()["status"] == "ok"

    # ALL (ZIP) Endpoint
    resp = client.post("/api/export/all", json=payload)
    assert resp.status_code == 200
    data_all = resp.json()
    assert data_all["status"] == "ok"
    assert data_all["format"] == "zip"

    # Download ZIP
    zip_dl = client.get(data_all["download_url"])
    assert zip_dl.status_code == 200
    assert zip_dl.headers["content-type"] == "application/zip"


def test_security_path_traversal():
    client = TestClient(app)
    # Attempt path traversal in download
    bad_resp = client.get("/api/export/download/../../etc/passwd")
    assert bad_resp.status_code in [404, 403, 422]


def req_options():
    from services.document_export.models import PDFExportOptions
    return PDFExportOptions()

def word_options():
    from services.document_export.models import WordExportOptions
    return WordExportOptions()

def pptx_options():
    from services.document_export.models import PPTXExportOptions
    return PPTXExportOptions()

def excel_options():
    from services.document_export.models import ExcelExportOptions
    return ExcelExportOptions()
