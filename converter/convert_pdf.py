from __future__ import annotations

import base64
import json
import re
import sys
from pathlib import Path
from typing import Iterable

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    import fitz  # type: ignore[import-not-found]
except Exception:  # pragma: no cover - optional dependency
    fitz = None

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

LOGS: list[str] = []


def log(msg: str) -> None:
    LOGS.append(msg)
    print(msg, file=sys.stderr)


def clean_cell(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\n", " ")).strip()


def merge_rows(rows: list[list[str]]) -> list[list[str]]:
    if not rows:
        return []

    merged_rows = []
    current_row = rows[0]

    for next_row in rows[1:]:
        has_content = any(cell.strip() for cell in next_row)
        # A row is a continuation if:
        # 1. Its first column is empty.
        # 2. It has at least one cell with content.
        # 3. It has a high proportion of empty cells (more than half).
        is_continuation = (
            has_content and
            (not next_row[0].strip()) and
            (sum(1 for cell in next_row if not cell.strip()) > len(next_row) / 2)
        )

        if is_continuation:
            # Merge cell values column by column
            for col_idx in range(len(current_row)):
                if col_idx < len(next_row):
                    next_val = next_row[col_idx].strip()
                    if next_val:
                        if current_row[col_idx].strip():
                            current_row[col_idx] = current_row[col_idx] + " " + next_val
                        else:
                            current_row[col_idx] = next_val
        else:
            merged_rows.append(current_row)
            current_row = next_row

    merged_rows.append(current_row)
    return merged_rows


def normalize_table(table: list[list[object]]) -> list[list[str]]:
    rows = [[clean_cell(cell) for cell in row] for row in table]
    non_empty_rows = [row for row in rows if any(cell for cell in row)]
    return merge_rows(non_empty_rows)


def table_has_quote_data(rows: list[list[str]]) -> bool:
    joined = " ".join(" ".join(row) for row in rows).lower()
    quote_markers = ["product", "po price", "單價", "總計", "qty", "quotation", "nvidia", "beehe"]
    non_empty_cells = sum(1 for row in rows for cell in row if cell)
    return non_empty_cells >= 4 and any(marker in joined for marker in quote_markers)


def find_header_row_index(rows: list[list[str]]) -> int:
    # Rule 1: Find first row starting with a sequence digit "1", "01", or "1."
    # and return the index of the row before it.
    for idx in range(1, len(rows)):
        first_cell = rows[idx][0].strip()
        if first_cell in ("1", "01", "1."):
            prev_cell = rows[idx - 1][0].strip()
            if not prev_cell.isdigit():
                return idx - 1

    # Rule 2: Search for typical header keywords.
    header_keywords = {"no.", "no", "item", "product", "part number", "description", "專案", "內容", "單價"}
    for idx, row in enumerate(rows):
        for cell in row:
            val = cell.lower().strip()
            if any(kw in val for kw in header_keywords):
                return idx
    return 0


def is_footer_row(row: list[str]) -> bool:
    if not row:
        return True
    non_empty = [c.strip() for c in row if c.strip()]
    if not non_empty:
        return True
    if len(non_empty) == 1:
        val = non_empty[0].lower()
        footer_kws = ["備註", "remark", "note", "付款", "交易", "保固", "有效期限", "sign", "signature", "址", "tel", "fax"]
        if any(kw in val for kw in footer_kws) or len(val) > 40:
            return True
    return False


def clean_quote_table(rows: list[list[str]]) -> list[list[str]]:
    if not rows:
        return []
    header_idx = find_header_row_index(rows)
    table_rows = rows[header_idx:]
    if not table_rows:
        return []

    cleaned_rows = [table_rows[0]]  # Start with the header row
    for row in table_rows[1:]:
        first_cell = row[0].strip()
        row_joined = " ".join(row).lower()

        # Check if row looks like a total row
        if "ttl" in first_cell.lower() or "total" in first_cell.lower() or "ttl" in row_joined or "total" in row_joined:
            joined_all = " ".join(c.strip() for c in row if c.strip())
            ttl_match = re.search(r"(?i)(ttl|total|grand\s*total)\s*(\([^)]+\))?\s*([\d,.]+)", joined_all)
            if ttl_match:
                new_row = [""] * len(row)
                label = ttl_match.group(1)
                if ttl_match.group(2):
                    label += " " + ttl_match.group(2)
                val = ttl_match.group(3)
                new_row[-1] = val
                label_idx = -2 if len(row) < 6 else -3
                new_row[label_idx] = label
                cleaned_rows.append(new_row)
            else:
                cleaned_rows.append(row)
            break  # Discard everything after total

        if is_footer_row(row):
            break  # Discard this footer row and everything after it

        cleaned_rows.append(row)

    return cleaned_rows



def render_page_thumbnail(doc: object | None, page_index: int) -> str:
    if fitz is None or doc is None:
        return ""

    page = doc.load_page(page_index - 1)
    target_width = 1200.0
    scale = min(3.0, target_width / max(float(page.rect.width), 1.0))
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    encoded = base64.b64encode(pix.tobytes("jpg")).decode("ascii")
    return f"data:image/jpeg;base64,{encoded}"


def parse_amd_text_table(text: str) -> list[list[str]] | None:
    log("Running AMD fallback text-based table parser...")
    lines = [line.strip() for line in text.split("\n")]
    
    header_idx = -1
    for idx, line in enumerate(lines):
        if "Product" in line and "Start Date" in line and "End Date" in line and "Qty" in line and "Pricing Term" in line:
            header_idx = idx
            break
            
    if header_idx == -1:
        log("AMD fallback parser: Required header keywords not found in text.")
        return None
        
    log(f"AMD fallback parser: Found header keywords at line {header_idx}: '{lines[header_idx]}'")
    headers = ["Product", "Start Date", "End Date", "Qty", "Pricing Term"]
    rows = [headers]
    
    idx = header_idx + 1
    while idx < len(lines):
        line = lines[idx]
        if not line or line.startswith("___") or "Eligible Incentives" in line:
            log(f"AMD fallback parser: Stopping extraction at line {idx} due to stop delimiter: '{line}'")
            break
            
        # Match dates in the line to robustly split the columns even with spaces in the product name
        date_matches = list(re.finditer(r"\b\d{2}/\d{2}/\d{4}\b", line))
        if len(date_matches) >= 2:
            start_date = date_matches[0].group(0)
            end_date = date_matches[1].group(0)
            
            # Everything before the first date is the product name
            product = line[:date_matches[0].start()].strip()
            
            # Everything after the second date contains Qty and Pricing Term
            rest = line[date_matches[1].end():].strip()
            rest_parts = rest.split()
            if len(rest_parts) >= 2:
                qty = rest_parts[0]
                price = rest_parts[1]
                term = " ".join(rest_parts[2:])
                
                # Check if next line contains the wrapped decimal digit (e.g. '$600.0' + '0')
                if idx + 1 < len(lines):
                    next_line = lines[idx + 1].strip()
                    if next_line.isdigit() and len(next_line) == 1:
                        price = price + next_line
                        idx += 1
                        
                rows.append([product, start_date, end_date, qty, f"{price} {term}".strip()])
                log(f"AMD fallback parser: Extracted row: {rows[-1]}")
        idx += 1
        
    if len(rows) > 1:
        log(f"AMD fallback parser: Successfully extracted {len(rows) - 1} data rows.")
        return rows
    log("AMD fallback parser: No data rows were extracted.")
    return None


def extract_document(pdf_path: Path, target_pages: set[int] | None = None) -> tuple[list[dict[str, object]], list[dict[str, object]]]:
    log(f"Extracting tables from document: {pdf_path.name}")
    tables: list[dict[str, object]] = []
    pages: list[dict[str, object]] = []
    pdf_doc = fitz.open(pdf_path) if fitz is not None else None

    with pdfplumber.open(pdf_path) as pdf:
        for page_index, page in enumerate(pdf.pages, start=1):
            if target_pages is not None and page_index not in target_pages:
                log(f"Page {page_index}: Skipped (not in target pages).")
                continue
            text = page.extract_text() or ""
            log(f"Page {page_index}: Extracted text (length={len(text)}).")
            pages.append(
                {
                    "page": page_index,
                    "text": text,
                    "thumbnail": render_page_thumbnail(pdf_doc, page_index),
                }
            )
            
            page_tables = []
            
            # Try extraction with normal tables
            raw_tables = page.extract_tables() or []
            log(f"Page {page_index}: pdfplumber detected {len(raw_tables)} raw tables.")
            for table_idx, table in enumerate(raw_tables, start=1):
                normalized = normalize_table(table)
                if normalized and table_has_quote_data(normalized):
                    cleaned = clean_quote_table(normalized)
                    # Check that the table has more than just the header row
                    if cleaned and len(cleaned) > 1:
                        page_tables.append(cleaned)
                        log(f"Page {page_index} Table {table_idx}: Accepted grid table with {len(cleaned)} rows (including header).")
                    else:
                        log(f"Page {page_index} Table {table_idx}: Ignored grid table with {len(cleaned) if cleaned else 0} rows (no data).")
                else:
                    log(f"Page {page_index} Table {table_idx}: Ignored (does not match quote signature or too small).")
            
            # Fallback if no tables found and it looks like an AMD quote
            if not page_tables and ("AMD" in text or "Advanced Micro Devices" in text):
                log(f"Page {page_index}: No grid tables accepted, but AMD detected. Triggering AMD fallback parser.")
                amd_table = parse_amd_text_table(text)
                if amd_table:
                    page_tables.append(amd_table)
            
            # Register page tables
            for table_idx, normalized in enumerate(page_tables, start=1):
                tables.append(
                    {
                        "page": page_index,
                        "index": table_idx,
                        "title": f"Page {page_index} Table {table_idx}",
                        "rows": normalized,
                    }
                )
                log(f"Page {page_index}: Registered table {table_idx} with {len(normalized)} rows.")
                
    if pdf_doc is not None:
        pdf_doc.close()
    return tables, pages


def write_table(ws, start_row: int, title: str, rows: list[list[str]]) -> int:
    max_cols = max((len(row) for row in rows), default=1)

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


def write_text_fallback(wb: Workbook, pdf_path: Path, target_pages: set[int] | None = None) -> None:
    ws = wb.create_sheet("Text")
    ws.append(["Page", "Text"])
    with pdfplumber.open(pdf_path) as pdf:
        for page_index, page in enumerate(pdf.pages, start=1):
            if target_pages is not None and page_index not in target_pages:
                continue
            ws.append([page_index, page.extract_text() or ""])
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 100
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"


def convert(pdf_path: Path, output_path: Path, target_pages: set[int] | None = None) -> None:
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    tables, pages = extract_document(pdf_path, target_pages)
    wb = Workbook()
    ws = wb.active
    ws.title = "Tables"

    row_cursor = 1
    for table in tables:
        row_cursor = write_table(ws, row_cursor, table["title"], table["rows"])

    if not tables:
        ws.append(["No table-like quotation data detected. See Text sheet."])

    autosize(ws)
    ws.freeze_panes = "A2"
    write_text_fallback(wb, pdf_path, target_pages)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)

    payload = {
        "input": str(pdf_path),
        "output": str(output_path),
        "table_count": len(tables),
        "tables": tables,
        "pages": pages,
        "logs": LOGS,
    }
    print(json.dumps(payload, ensure_ascii=False))


def get_thumbnails(pdf_path: Path) -> None:
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    pages = []
    pdf_doc = fitz.open(pdf_path) if fitz is not None else None
    if pdf_doc is not None:
        for page_index in range(1, len(pdf_doc) + 1):
            pages.append(
                {
                    "page": page_index,
                    "text": "",
                    "thumbnail": render_page_thumbnail(pdf_doc, page_index),
                }
            )
        pdf_doc.close()
    else:
        # Fallback if fitz is not available
        with pdfplumber.open(pdf_path) as pdf:
            for page_index in range(1, len(pdf.pages) + 1):
                pages.append(
                    {
                        "page": page_index,
                        "text": "",
                        "thumbnail": "",
                    }
                )

    payload = {
        "input": str(pdf_path),
        "output": "",
        "table_count": 0,
        "tables": [],
        "pages": pages,
        "logs": LOGS,
    }
    print(json.dumps(payload, ensure_ascii=False))


def main() -> int:
    if len(sys.argv) < 3:
        print("Usage: convert_pdf.py <input.pdf> <output.xlsx | --thumbnails-only> [--pages 1,2,3]", file=sys.stderr)
        return 2
    try:
        input_path = Path(sys.argv[1])
        output_arg = sys.argv[2]
        
        target_pages = None
        if "--pages" in sys.argv:
            try:
                pages_idx = sys.argv.index("--pages")
                pages_str = sys.argv[pages_idx + 1]
                target_pages = {int(p) for p in pages_str.split(",") if p.strip()}
            except Exception as e:
                print(f"Error parsing --pages: {e}", file=sys.stderr)
                return 2

        if output_arg == "--thumbnails-only":
            get_thumbnails(input_path)
        else:
            convert(input_path, Path(output_arg), target_pages)
    except Exception as error:
        print(str(error), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
