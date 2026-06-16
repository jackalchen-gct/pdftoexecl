from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import Iterable

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")


def clean_cell(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\n", " ")).strip()


def normalize_table(table: list[list[object]]) -> list[list[str]]:
    rows = [[clean_cell(cell) for cell in row] for row in table]
    return [row for row in rows if any(cell for cell in row)]


def iter_tables(pdf_path: Path) -> Iterable[tuple[int, int, list[list[str]]]]:
    with pdfplumber.open(pdf_path) as pdf:
        for page_index, page in enumerate(pdf.pages, start=1):
            for table_index, table in enumerate(page.extract_tables() or [], start=1):
                normalized = normalize_table(table)
                if normalized and table_has_quote_data(normalized):
                    yield page_index, table_index, normalized


def table_has_quote_data(rows: list[list[str]]) -> bool:
    joined = " ".join(" ".join(row) for row in rows).lower()
    quote_markers = ["product", "po price", "單價", "總計", "qty", "quotation", "nvidia", "beehe"]
    non_empty_cells = sum(1 for row in rows for cell in row if cell)
    return non_empty_cells >= 4 and any(marker in joined for marker in quote_markers)


def write_table(ws, start_row: int, title: str, rows: list[list[str]]) -> int:
    max_cols = max((len(row) for row in rows), default=1)
    ws.cell(start_row, 1, title)
    ws.cell(start_row, 1).font = Font(bold=True, size=13)
    start_row += 1

    for row_offset, row in enumerate(rows):
        for col_index in range(1, max_cols + 1):
            cell = ws.cell(start_row + row_offset, col_index)
            cell.value = row[col_index - 1] if col_index <= len(row) else ""
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if row_offset == 0:
                cell.font = Font(bold=True)
                cell.fill = PatternFill("solid", fgColor="E8DAC0")
    return start_row + len(rows) + 2


def autosize(ws) -> None:
    for column in ws.columns:
        letter = get_column_letter(column[0].column)
        width = max(len(str(cell.value or "")) for cell in column)
        ws.column_dimensions[letter].width = min(max(width + 2, 10), 42)


def write_text_fallback(wb: Workbook, pdf_path: Path) -> None:
    ws = wb.create_sheet("Text")
    ws.append(["Page", "Text"])
    with pdfplumber.open(pdf_path) as pdf:
        for page_index, page in enumerate(pdf.pages, start=1):
            ws.append([page_index, page.extract_text() or ""])
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 100
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"


def convert(pdf_path: Path, output_path: Path) -> None:
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    wb = Workbook()
    ws = wb.active
    ws.title = "Tables"

    row_cursor = 1
    table_count = 0
    for page_index, table_index, rows in iter_tables(pdf_path):
        table_count += 1
        row_cursor = write_table(ws, row_cursor, f"Page {page_index} Table {table_index}", rows)

    if table_count == 0:
        ws.append(["No table-like quotation data detected. See Text sheet."])

    autosize(ws)
    ws.freeze_panes = "A2"
    write_text_fallback(wb, pdf_path)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def main() -> int:
    if len(sys.argv) != 3:
        print("Usage: convert_pdf.py <input.pdf> <output.xlsx>", file=sys.stderr)
        return 2
    try:
        convert(Path(sys.argv[1]), Path(sys.argv[2]))
    except Exception as error:
        print(str(error), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
