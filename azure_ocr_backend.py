#!/usr/bin/env python3
"""Azure Document Intelligence OCR backend for HSBC statement extraction.

Replaces RapidOCR with Azure prebuilt-layout. Converts Azure table output into
OcrLine objects that the existing hsbc_statement_to_excel parser expects.
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

# ---------------------------------------------------------------------------
# Azure credentials — read from environment or a local config file
# ---------------------------------------------------------------------------

AZURE_ENDPOINT = os.environ.get(
    "AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT",
    "https://documentbank.cognitiveservices.azure.com/"
)
AZURE_KEY = os.environ.get("AZURE_DOCUMENT_INTELLIGENCE_KEY", "")

# Fallback: read from a JSON file next to this module if env vars not set
if not AZURE_KEY:
    _cfg_path = Path(__file__).with_name("azure_config.json")
    if _cfg_path.exists():
        import json

        _cfg = json.loads(_cfg_path.read_text())
        AZURE_ENDPOINT = _cfg.get("endpoint", AZURE_ENDPOINT)
        AZURE_KEY = _cfg.get("key", "")

# Production builds should NOT include test_azure.py —
# the key must come from env var or azure_config.json only.
# The dev fallback below is disabled by default; set CARMEN_DEV=1 to enable.
if not AZURE_KEY and os.environ.get("CARMEN_DEV") == "1":
    try:
        import importlib.util
        _test_azure_path = Path(__file__).parent / "test_azure.py"
        if _test_azure_path.exists():
            _spec = importlib.util.spec_from_file_location("test_azure", _test_azure_path)
            _mod = importlib.util.module_from_spec(_spec)
            _spec.loader.exec_module(_mod)
            AZURE_KEY = getattr(_mod, "KEY", "")
            AZURE_ENDPOINT = getattr(_mod, "ENDPOINT", AZURE_ENDPOINT)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Public types (mirrors hsbc_statement_to_excel)
# ---------------------------------------------------------------------------

@dataclass
class OcrLine:
    page: int
    y: float
    date_col: str
    detail_col: str
    deposit_col: str
    withdrawal_col: str
    balance_col: str


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _normalize_spaces(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "")).strip()


def _extract_year_from_text(text: str) -> int | None:
    """Pull a year like 2025 from statement text."""
    for m in re.finditer(r"\b(20\d{2})\b", text):
        year = int(m.group(1))
        if 2020 <= year <= 2035:
            return year
    return None


def _is_transaction_table(table: list[list[str]]) -> bool:
    """True if the table header smells like a transaction grid."""
    if len(table) < 3 or len(table[0]) < 4:
        return False
    header = " ".join(table[0]).lower()
    # Must have Date + at least one of deposit/withdrawal/balance
    has_date = "date" in header or "日期" in header
    has_amount_col = any(k in header for k in ["deposit", "withdrawal", "balance", "存入", "支出", "結餘"])
    return has_date and has_amount_col


def _build_ocr_lines_from_azure_tables(tables: list[list[list[str]]]) -> list[OcrLine]:
    """Convert Azure table grids into OcrLine objects the existing parser can eat."""
    lines: list[OcrLine] = []

    for table in tables:
        if not _is_transaction_table(table):
            continue

        # Identify header row and column indices
        header = [_normalize_spaces(h).lower() for h in table[0]]
        header_text = " ".join(header)

        # Some tables have a page-header row before the real table header
        header_row_idx = 0
        if "date" not in header_text and "transaction" not in header_text and "details" not in header_text:
            if len(table) > 1:
                header = [_normalize_spaces(h).lower() for h in table[1]]
                header_row_idx = 1

        date_idx = next((i for i, h in enumerate(header) if "date" in h), None)
        desc_idx = next(
            (i for i, h in enumerate(header) if "detail" in h or "description" in h or "進支詳情" in h), None
        )
        deposit_idx = next((i for i, h in enumerate(header) if "deposit" in h or "存入" in h), None)
        withdrawal_idx = next((i for i, h in enumerate(header) if "withdrawal" in h or "支出" in h), None)
        balance_idx = next((i for i, h in enumerate(header) if "balance" in h or "結餘" in h), None)

        # If we can't map columns, skip
        if date_idx is None and desc_idx is None:
            continue

        # Fake a y-position so rows sort in order.  We use row index * 100.
        for row_offset, row in enumerate(table[header_row_idx + 1 :], start=1):
            if len(row) < 3:
                continue

            def cell(idx: int | None) -> str:
                if idx is None or idx >= len(row):
                    return ""
                return _normalize_spaces(row[idx])

            date_val = cell(date_idx)
            desc_val = cell(desc_idx)
            deposit_val = cell(deposit_idx)
            withdrawal_val = cell(withdrawal_idx)
            balance_val = cell(balance_idx)

            # Skip obvious non-data rows
            joined = f"{date_val} {desc_val} {deposit_val} {withdrawal_val} {balance_val}".strip().lower()
            if "b/f balance" in joined or "balance b/f" in joined:
                # Still emit — the parser uses B/F to detect opening balance
                pass
            if "total" in joined and not deposit_val and not withdrawal_val:
                continue

            lines.append(
                OcrLine(
                    page=1,  # Azure merges pages; we lose page granularity but parser doesn't care much
                    y=float(row_offset * 100),
                    date_col=date_val,
                    detail_col=desc_val,
                    deposit_col=deposit_val,
                    withdrawal_col=withdrawal_val,
                    balance_col=balance_val,
                )
            )

    return lines


# ---------------------------------------------------------------------------
# Azure DI call
# ---------------------------------------------------------------------------

def _azure_extract_tables(pdf_path: Path) -> list[list[list[str]]]:
    """Send PDF to Azure prebuilt-layout and return tables as nested lists."""
    from azure.ai.formrecognizer import DocumentAnalysisClient
    from azure.core.credentials import AzureKeyCredential

    client = DocumentAnalysisClient(AZURE_ENDPOINT, AzureKeyCredential(AZURE_KEY))

    with pdf_path.open("rb") as f:
        poller = client.begin_analyze_document("prebuilt-layout", f)
        result = poller.result()

    tables: list[list[list[str]]] = []
    for table in result.tables:
        grid: dict[tuple[int, int], str] = {}
        for cell in table.cells:
            grid[(cell.row_index, cell.column_index)] = cell.content.strip()

        rows: list[list[str]] = []
        for r in range(table.row_count):
            row = [grid.get((r, c), "") for c in range(table.column_count)]
            rows.append(row)
        tables.append(rows)

    return tables


# ---------------------------------------------------------------------------
# Public API — drop-in replacement for ocr_pdf_to_lines()
# ---------------------------------------------------------------------------

def azure_ocr_pdf_to_lines(pdf_path: Path) -> list[OcrLine]:
    """Primary OCR path: Azure Document Intelligence."""
    tables = _azure_extract_tables(pdf_path)
    return _build_ocr_lines_from_azure_tables(tables)


def rapidocr_pdf_to_lines(pdf_path: Path, scale_dpi: int = 450) -> list[OcrLine]:
    """Fallback OCR path: local RapidOCR (original implementation)."""
    # Import inline so missing deps don't break import of this module
    try:
        from rapidocr_onnxruntime import RapidOCR
        import pypdfium2 as pdfium
        import numpy as np
    except Exception:
        return []

    doc = pdfium.PdfDocument(str(pdf_path))
    ocr = RapidOCR()
    lines: list[OcrLine] = []
    dpi_scale = scale_dpi / 300.0
    x_date_max = 330 * dpi_scale
    x_detail_max = 1360 * dpi_scale
    x_deposit_max = 1680 * dpi_scale
    x_withdrawal_max = 1940 * dpi_scale

    for page_idx in range(len(doc)):
        pil_image = doc[page_idx].render(scale=scale_dpi / 72).to_pil()
        img_array = np.array(pil_image)
        result, _ = ocr(img_array)
        points = []
        for box, text, _confidence in (result or []):
            x = min(point[0] for point in box)
            y = min(point[1] for point in box)
            points.append((y, x, text))
        points.sort()

        grouped: list[tuple[float, list[tuple[float, str]]]] = []
        for y, x, text in points:
            if not grouped or abs(y - grouped[-1][0]) > 25:
                grouped.append((y, [(x, text)]))
            else:
                grouped[-1][1].append((x, text))

        for y, cells in grouped:
            ordered = sorted(cells)
            date_col = " ".join(text for x, text in ordered if x < x_date_max)
            detail_col = " ".join(text for x, text in ordered if x_date_max <= x < x_detail_max)
            deposit_col = " ".join(text for x, text in ordered if x_detail_max <= x < x_deposit_max)
            withdrawal_col = " ".join(text for x, text in ordered if x_deposit_max <= x < x_withdrawal_max)
            balance_col = " ".join(text for x, text in ordered if x_withdrawal_max <= x)
            lines.append(
                OcrLine(
                    page=page_idx + 1,
                    y=y,
                    date_col=_normalize_spaces(date_col),
                    detail_col=_normalize_spaces(detail_col),
                    deposit_col=_normalize_spaces(deposit_col),
                    withdrawal_col=_normalize_spaces(withdrawal_col),
                    balance_col=_normalize_spaces(balance_col),
                )
            )

    return lines


# ---------------------------------------------------------------------------
# Unified entry point
# ---------------------------------------------------------------------------

def ocr_pdf_to_lines(pdf_path: Path, prefer_azure: bool = True) -> list[OcrLine]:
    """Try Azure first, fall back to RapidOCR if Azure unavailable or returns nothing."""
    if prefer_azure and AZURE_KEY:
        try:
            lines = azure_ocr_pdf_to_lines(pdf_path)
            if lines:
                return lines
        except Exception:
            pass  # fall through

    # Fallback or disabled
    if os.environ.get("HSBC_DISABLE_RAPIDOCR", "").strip() == "1":
        return []
    return rapidocr_pdf_to_lines(pdf_path)


# ---------------------------------------------------------------------------
# CLI test
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python azure_ocr_backend.py <pdf_path>")
        sys.exit(1)

    pdf = Path(sys.argv[1])
    lines = ocr_pdf_to_lines(pdf)
    print(f"Extracted {len(lines)} OCR lines from {pdf.name}")
    for line in lines[:20]:
        print(f"  y={line.y:6.1f} | {line.date_col:12s} | {line.detail_col:40s} | {line.deposit_col:14s} | {line.withdrawal_col:14s} | {line.balance_col:14s}")
