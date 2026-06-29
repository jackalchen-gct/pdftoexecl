from __future__ import annotations

import base64
import json
import re
import sys
from pathlib import Path
from typing import Iterable

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image

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
    lines = [re.sub(r"[ \t]+", " ", line).strip() for line in str(value).split("\n")]
    return "\n".join(lines).strip()


def parse_price(val: str) -> float:
    if not val:
        return 0.0
    cleaned = val.replace(",", "").replace("USD", "").replace("$", "").replace("-", "").strip()
    if not cleaned:
        return 0.0
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def style_range(ws, cell_range: str, border: Border = None, fill: PatternFill = None, font: Font = None, alignment: Alignment = None):
    for row in ws[cell_range]:
        for cell in row:
            if fill is not None:
                cell.fill = fill
            if font is not None:
                cell.font = font
            if alignment is not None:
                cell.alignment = alignment
            if border is not None:
                cell.border = border

def style_merged_range_borders(ws, cell_range: str, border_side: Side):
    from openpyxl.utils import range_boundaries
    min_col, min_row, max_col, max_row = range_boundaries(cell_range)
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            cell = ws.cell(r, c)
            left = border_side if c == min_col else None
            right = border_side if c == max_col else None
            top = border_side if r == min_row else None
            bottom = border_side if r == max_row else None
            cell.border = Border(left=left, right=right, top=top, bottom=bottom)

def parse_customer_info(text: str) -> dict[str, str]:
    info = {"customer": "", "contact": "", "phone": "", "date": "", "project": ""}
    if not text:
        return info
    lines = [line.strip() for line in text.split("\n") if line.strip()]
    for line in lines:
        cust_match = re.search(r'(?:客戶|客户)\s*:\s*(.*?)(?=(?:聯繫|联系|聯絡人|報價日期|报价日期|電話|电话|$))', line)
        if cust_match:
            info["customer"] = cust_match.group(1).strip()
            
        date_match = re.search(r'(?:報價日期|报价日期)\s*:\s*(\d{4}[./-]\d{2}[./-]\d{2})', line)
        if date_match:
            info["date"] = date_match.group(1).strip()
            
        contact_match = re.search(r'(?:聯繫|联系|聯絡人)\s*:\s*(.*?)(?=(?:客戶|客户|報價日期|报价日期|電話|电话|$))', line)
        if contact_match:
            info["contact"] = contact_match.group(1).strip()
            
        phone_match = re.search(r'(?:電話|电话).+?(\+?[\d\s-]{9,})', line)
        if phone_match:
            info["phone"] = phone_match.group(1).strip()
        else:
            phone_match = re.search(r'(?:電話|电话)\s*:\s*([\d\s+-]*)', line)
            if phone_match and phone_match.group(1).strip():
                info["phone"] = phone_match.group(1).strip()
            
        if re.match(r'^\+?[\d-]+$', line):
            info["phone"] = line
            
    project_match = re.search(r'([^\s\n]*(?:報價|报价)(?!日期)[^\s\n]*)', text)
    if project_match:
        info["project"] = project_match.group(1).strip()
        info["project"] = re.sub(r'\s*No\.\s*\d+', '', info["project"])
        
    return info

def split_remarks(text: str) -> list[str]:
    if not text:
        return []
    raw_lines = [line.strip() for line in text.split("\n") if line.strip()]
    final_lines = []
    for line in raw_lines:
        matches = list(re.finditer(r'(\d+\))', line))
        if not matches:
            final_lines.append(line)
            continue
        first_part = line[:matches[0].start()].strip()
        if first_part:
            final_lines.append(first_part)
        for i in range(len(matches)):
            start_pos = matches[i].start()
            end_pos = matches[i+1].start() if i + 1 < len(matches) else len(line)
            part = line[start_pos:end_pos].strip()
            if part:
                final_lines.append(part)
    return final_lines


def extract_beehe_document(pdf_path: Path, target_pages: set[int] | None = None) -> list[dict[str, object]]:
    log(f"Extracting Beehe tables from document: {pdf_path.name}")
    tables = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page_index, page in enumerate(pdf.pages, start=1):
            if target_pages is not None and page_index not in target_pages:
                continue
            
            find_tables = page.find_tables() or []
            if not find_tables:
                continue
                
            for table_idx, t in enumerate(find_tables, start=1):
                raw_table = t.extract()
                grid = [[clean_cell(cell) for cell in row] for row in raw_table]
                
                header_idx = -1
                for idx, r in enumerate(grid):
                    row_joined = " ".join(r).lower()
                    if "no" in row_joined and ("專案" in row_joined or "項目" in row_joined or "內容" in row_joined):
                        header_idx = idx
                        break
                        
                if header_idx == -1:
                    log(f"Page {page_index} Table {table_idx}: Header row with 'No.' not found.")
                    continue
                    
                is_split_header = False
                if header_idx > 0:
                    prev_row_vals = grid[header_idx - 1]
                    if not prev_row_vals[0].strip() and any(prev_row_vals):
                        is_split_header = True
                
                if is_split_header:
                    top_y = t.rows[header_idx - 1].bbox[1]
                else:
                    top_y = t.rows[header_idx].bbox[1]
                    
                # Parse customer info
                cust_text = grid[1][0] if len(grid) > 1 and grid[1] else ""
                customer_info = parse_customer_info(cust_text)
                
                headers = []
                for col_idx in range(len(grid[header_idx])):
                    parts = []
                    for r_idx in (header_idx - 1, header_idx, header_idx + 1):
                        if 0 <= r_idx < len(grid):
                            val = grid[r_idx][col_idx].strip()
                            if val and val not in parts:
                                parts.append(val)
                    headers.append(" ".join(parts))
                    
                col_no = 0
                col_item = 1
                col_desc = 2
                col_price = 3
                col_warranty = 4
                col_qty = 5
                col_unit = 6
                col_total = 7
                
                for i, h in enumerate(headers):
                    h_clean = h.lower().replace(" ", "")
                    if "no" in h_clean:
                        col_no = i
                    elif "專案" in h_clean or "項目" in h_clean:
                        col_item = i
                    elif "內容" in h_clean:
                        col_desc = i
                    elif "單價(usd)/台" in h_clean or "保固單價" in h_clean or (i == 4 and "保固" in h_clean):
                        col_warranty = i
                    elif "單價" in h_clean:
                        col_price = i
                    elif "qty" in h_clean or "數量" in h_clean:
                        col_qty = i
                        if i + 1 < len(headers) and not headers[i + 1]:
                            col_unit = i + 1
                    elif "總計" in h_clean or "ttl" in h_clean or "total" in h_clean:
                        col_total = i
                
                has_remark_col = False
                remark_col_idx = -1
                for i, h in enumerate(headers):
                    if "備註" in h or "remark" in h.lower():
                        has_remark_col = True
                        remark_col_idx = i
                        break
                
                col_order = [col_no, col_item, col_desc, col_qty, col_unit, col_price, col_warranty, col_total]
                if has_remark_col:
                    col_order.append(remark_col_idx)
                    
                new_headers = [headers[idx] for idx in col_order]
                
                data_rows = []
                total_row_raw = None
                total_row_idx = -1
                
                for r_idx, r in enumerate(grid[header_idx + 1:], start=header_idx + 1):
                    if not r or not any(c.strip() for c in r if c is not None):
                        continue
                    first_cell = r[0].strip()
                    if re.match(r'^\d+$', first_cell):
                        data_rows.append(r)
                    elif first_cell.lower().startswith("ttl") or first_cell.lower().startswith("total") or re.search(r'\b(ttl|total)\b', " ".join(r).lower()):
                        total_row_raw = r
                        total_row_idx = r_idx
                
                if total_row_idx != -1:
                    bottom_y = t.rows[total_row_idx].bbox[3]
                else:
                    bottom_y = t.rows[-1].bbox[3]
                    
                # Parse remarks
                remarks_text_parts = []
                if total_row_idx != -1:
                    for r_rem in grid[total_row_idx + 1:]:
                        if r_rem and r_rem[0]:
                            remarks_text_parts.append(r_rem[0])
                remarks_text = "\n".join(remarks_text_parts)
                remarks_lines = split_remarks(remarks_text)
                
                reordered_data_rows = []
                for r in data_rows:
                    reordered_row = [r[idx] if idx < len(r) else "" for idx in col_order]
                    reordered_data_rows.append(reordered_row)
                
                r_col_no = 0
                r_col_item = 1
                r_col_desc = 2
                r_col_qty = 3
                r_col_unit = 4
                r_col_price = 5
                r_col_warranty = 6
                r_col_total = 7
                
                warranty_row_idx = -1
                warranty_years = 3
                for idx, r in enumerate(reordered_data_rows):
                    item_name = r[r_col_item].strip()
                    desc_text = r[r_col_desc].strip()
                    
                    match = None
                    if item_name:
                        match = re.search(r'^(\d+)年', item_name)
                    if not match:
                        match = re.search(r'^(\d+)年', desc_text)
                        
                    if match:
                        combined_text = (item_name + " " + desc_text).lower()
                        if any(kw in combined_text for kw in ("保固", "7x24", "小時", "到場", "維護", "維保", "教育訓練")):
                            warranty_row_idx = idx
                            warranty_years = int(match.group(1))
                            break
                
                other_items_sum = 0.0
                ratio = 0.35
                calculated_warranty_price = 0.0
                
                if warranty_row_idx != -1:
                    for idx, r in enumerate(reordered_data_rows):
                        if idx != warranty_row_idx:
                            other_items_sum += parse_price(r[r_col_total])
                            
                    ratios = {1: 0.35, 2: 0.35, 3: 0.35, 4: 0.45, 5: 0.55, 6: 0.65}
                    ratio = ratios.get(warranty_years, 0.35)
                    calculated_warranty_price = other_items_sum * ratio
                    
                    warranty_price_str = f"{calculated_warranty_price:,.0f}"
                    reordered_data_rows[warranty_row_idx][r_col_price] = warranty_price_str
                    reordered_data_rows[warranty_row_idx][r_col_total] = warranty_price_str
                
                grand_total = sum(parse_price(r[r_col_total]) for r in reordered_data_rows)
                
                warranty_header_name = f"{warranty_years}年保固 單價(USD)/台"
                clean_headers = ["No.", "專案", "內容", "Qty", "單位", "單價(USD)", warranty_header_name, "總計(USD)"]
                if has_remark_col:
                    clean_headers.append("備註")
                
                final_rows = [clean_headers]
                final_rows.extend(reordered_data_rows)
                
                if total_row_raw:
                    total_label = "TTL (USD)"
                    total_row = [""] * len(clean_headers)
                    total_row[r_col_warranty] = total_label
                    total_row[r_col_total] = f"{grand_total:,.0f}"
                    if has_remark_col and len(total_row_raw) > remark_col_idx:
                        total_row.append(total_row_raw[remark_col_idx])
                    final_rows.append(total_row)
                    
                table_entry = {
                    "page": page_index,
                    "index": table_idx,
                    "title": f"Page {page_index} Table {table_idx}",
                    "rows": final_rows,
                    "top_y": top_y,
                    "bottom_y": bottom_y,
                    "customer_info": customer_info,
                    "remarks_lines": remarks_lines
                }
                if warranty_row_idx != -1:
                    table_entry["ratio"] = f"{ratio * 100:.0f}%"
                    table_entry["formula"] = f"保固計算公式: 其他項目合計 {other_items_sum:,.0f} * {ratio * 100:.0f}% = {calculated_warranty_price:,.0f} (USD)"
                tables.append(table_entry)
                
    return tables


def merge_rows(rows: list[list[str]]) -> list[list[str]]:
    if not rows:
        return []

    merged_rows = []
    current_row = rows[0]

    for next_row in rows[1:]:
        has_content = any(cell.strip() for cell in next_row)
        is_continuation = (
            has_content and
            (not next_row[0].strip()) and
            (sum(1 for cell in next_row if not cell.strip()) > len(next_row) / 2)
        )

        if is_continuation:
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
    for idx in range(1, len(rows)):
        first_cell = rows[idx][0].strip()
        if first_cell in ("1", "01", "1."):
            prev_cell = rows[idx - 1][0].strip()
            if not prev_cell.isdigit():
                return idx - 1

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

    cleaned_rows = [table_rows[0]]
    for row in table_rows[1:]:
        first_cell = row[0].strip()
        row_joined = " ".join(row).lower()

        if (first_cell.lower().startswith("ttl") or 
            first_cell.lower().startswith("total") or 
            re.search(r'\b(ttl|total)\b', row_joined)):
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
            break

        if is_footer_row(row):
            break

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
            
        date_matches = list(re.finditer(r"\b\d{2}/\d{2}/\d{4}\b", line))
        if len(date_matches) >= 2:
            start_date = date_matches[0].group(0)
            end_date = date_matches[1].group(0)
            product = line[:date_matches[0].start()].strip()
            
            rest = line[date_matches[1].end():].strip()
            rest_parts = rest.split()
            if len(rest_parts) >= 2:
                qty = rest_parts[0]
                price = rest_parts[1]
                term = " ".join(rest_parts[2:])
                
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
        is_beehe = False
        for page in pdf.pages:
            text = page.extract_text() or ""
            if "beehe" in text.lower() or "比赫" in text:
                is_beehe = True
                break
                
        if is_beehe:
            beehe_tables = extract_beehe_document(pdf_path, target_pages)
            for page_index, page in enumerate(pdf.pages, start=1):
                if target_pages is not None and page_index not in target_pages:
                    continue
                text = page.extract_text() or ""
                pages.append({
                    "page": page_index,
                    "text": text,
                    "thumbnail": render_page_thumbnail(pdf_doc, page_index),
                })
            if pdf_doc is not None:
                pdf_doc.close()
            return beehe_tables, pages

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
            raw_tables = page.extract_tables() or []
            log(f"Page {page_index}: pdfplumber detected {len(raw_tables)} raw tables.")
            for table_idx, table in enumerate(raw_tables, start=1):
                normalized = normalize_table(table)
                if normalized and table_has_quote_data(normalized):
                    cleaned = clean_quote_table(normalized)
                    if cleaned and len(cleaned) > 1:
                        page_tables.append(cleaned)
                        log(f"Page {page_index} Table {table_idx}: Accepted grid table with {len(cleaned)} rows (including header).")
                    else:
                        log(f"Page {page_index} Table {table_idx}: Ignored grid table with {len(cleaned) if cleaned else 0} rows (no data).")
                else:
                    log(f"Page {page_index} Table {table_idx}: Ignored (does not match quote signature or too small).")
            
            if not page_tables and ("AMD" in text or "Advanced Micro Devices" in text):
                log(f"Page {page_index}: No grid tables accepted, but AMD detected. Triggering AMD fallback parser.")
                amd_table = parse_amd_text_table(text)
                if amd_table:
                    page_tables.append(amd_table)
            
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

    # Check if this was a Beehe PDF
    is_beehe = False
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if "beehe" in text.lower() or "比赫" in text:
                is_beehe = True
                break

    tables, pages = extract_document(pdf_path, target_pages)
    wb = Workbook()
    ws = wb.active
    ws.title = "Tables"

    if is_beehe and fitz is not None:
        log("Applying custom Beehe layout with native styled cells and Logo crop.")
        pdf_doc = fitz.open(pdf_path)
        row_cursor = 1
        temp_images: list[Path] = []
        
        for table in tables:
            page_index = table["page"]
            page_fitz = pdf_doc.load_page(page_index - 1)
            width = page_fitz.rect.width
            height = page_fitz.rect.height
            
            # Crop Logo icon (from coords: x0=48.20, x1=92.32, top=65.14, bottom=106.14)
            logo_img_path = output_path.parent / f"temp_{output_path.stem}_logo_{page_index}.png"
            pix_logo = page_fitz.get_pixmap(clip=fitz.Rect(45.0, 60.0, 95.0, 110.0), matrix=fitz.Matrix(3.0, 3.0))
            pix_logo.save(str(logo_img_path))
            temp_images.append(logo_img_path)
            
            # Write layout blocks
            customer_info = table.get("customer_info", {"customer": "", "contact": "", "phone": "", "date": "", "project": ""})
            remarks = table.get("remarks_lines", [])
            
            rem_header = "* 備註"
            rem_body = list(remarks)
            if rem_body and any(kw in rem_body[0] for kw in ("* 備註", "備註", "remark", "Remark")):
                rem_header = rem_body.pop(0)
                
            rows = table["rows"]
            start_row = row_cursor + 8
            table.setdefault("table_start_row", start_row)
            
            # The last row of remarks will be at:
            end_row = start_row + len(rows) + 2 + len(rem_body)
            
            # Apply solid white background to the entire page block (A{row_cursor} to H{end_row})
            white_fill = PatternFill(fill_type="solid", fgColor="FFFFFF")
            style_range(ws, f"A{row_cursor}:H{end_row}", fill=white_fill)
            
            # 1. Company Header Area
            ws.row_dimensions[row_cursor].height = 18
            ws.row_dimensions[row_cursor + 1].height = 16
            ws.row_dimensions[row_cursor + 2].height = 16
            ws.row_dimensions[row_cursor + 3].height = 25
            
            ws.merge_cells(f"A{row_cursor}:B{row_cursor+2}")
            ws.merge_cells(f"C{row_cursor}:H{row_cursor}")
            ws.merge_cells(f"C{row_cursor+1}:H{row_cursor+1}")
            ws.merge_cells(f"C{row_cursor+2}:H{row_cursor+2}")
            ws.merge_cells(f"A{row_cursor+3}:H{row_cursor+3}")
            
            ws.cell(row_cursor, 3).value = "Beehe Electric (Taicang) Co., Ltd"
            ws.cell(row_cursor + 1, 3).value = "ADD: No.5-1 Shuanghu Road, Taicang City, Jiangsu Province,China"
            ws.cell(row_cursor + 2, 3).value = "TEL：0086-512-53983090 etx.8910"
            ws.cell(row_cursor + 3, 1).value = "Quotation"
            
            # Style text
            style_range(ws, f"C{row_cursor}:H{row_cursor}", font=Font(name="Calibri", size=12, bold=True), alignment=Alignment(vertical="center", horizontal="left"))
            style_range(ws, f"C{row_cursor+1}:H{row_cursor+2}", font=Font(name="Calibri", size=9), alignment=Alignment(vertical="center", horizontal="left"))
            style_range(ws, f"A{row_cursor+3}:H{row_cursor+3}", font=Font(name="Calibri", size=16, bold=True), alignment=Alignment(vertical="center", horizontal="center"))
            
            # Insert logo image in A{row_cursor} (spans A1:B3)
            img_logo = Image(str(logo_img_path))
            img_logo.width = int(50 * 96 / 72)
            img_logo.height = int(50 * 96 / 72)
            ws.add_image(img_logo, f"A{row_cursor}")
            
            # 2. Customer Info Area (Row 5 to Row 8)
            ws.row_dimensions[row_cursor + 4].height = 18
            ws.row_dimensions[row_cursor + 5].height = 18
            ws.row_dimensions[row_cursor + 6].height = 18
            ws.row_dimensions[row_cursor + 7].height = 18
            
            ws.cell(row_cursor + 4, 1).value = "客戶 : " + customer_info["customer"]
            ws.cell(row_cursor + 5, 1).value = "聯繫 : " + customer_info["contact"]
            ws.cell(row_cursor + 5, 6).value = "報價日期 : " + customer_info["date"]
            ws.cell(row_cursor + 6, 1).value = "電話 : " + customer_info["phone"]
            ws.cell(row_cursor + 7, 1).value = ("專案 : " + customer_info["project"]) if customer_info.get("project") else ""
            
            # Apply styles and borders before merging to ensure openpyxl serializes outer borders correctly
            thin_side = Side(style="thin", color="000000")
            thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
            
            for r in range(row_cursor + 4, row_cursor + 8):
                for c in range(1, 9):
                    cell = ws.cell(r, c)
                    cell.font = Font(name="Calibri", size=10)
                    cell.alignment = Alignment(vertical="center", horizontal="left", indent=1)
                    
                    # Compute outer border box
                    left = thin_side if c == 1 else None
                    right = thin_side if c == 8 else None
                    top = thin_side if r == row_cursor + 4 else None
                    bottom = thin_side if r == row_cursor + 7 else None
                    cell.border = Border(left=left, right=right, top=top, bottom=bottom)
            
            # Special font styling for the project cell
            ws.cell(row_cursor + 7, 1).font = Font(name="Calibri", size=10, bold=True, color="FF0000")
            
            # Merge cells after applying borders
            ws.merge_cells(f"A{row_cursor+4}:H{row_cursor+4}")
            ws.merge_cells(f"A{row_cursor+5}:E{row_cursor+5}")
            ws.merge_cells(f"F{row_cursor+5}:H{row_cursor+5}")
            ws.merge_cells(f"A{row_cursor+6}:H{row_cursor+6}")
            ws.merge_cells(f"A{row_cursor+7}:H{row_cursor+7}")
            
            # 3. Write Quotation Table starting at row_cursor + 8 (Row 9)
            max_cols = max((len(row) for row in rows), default=1)
            for row_offset, row in enumerate(rows):
                r_idx = start_row + row_offset
                for col_index in range(1, max_cols + 1):
                    cell = ws.cell(r_idx, col_index)
                    cell.value = row[col_index - 1] if col_index <= len(row) else ""
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
                    cell.border = thin_border
                    if row_offset == 0:
                        cell.font = Font(bold=True)
                        cell.fill = PatternFill("solid", fgColor="D9E1F2")
                        
            # Apply blackout to data rows in this table
            black_fill = PatternFill(fill_type="solid", fgColor="000000")
            black_font = Font(color="000000")
            
            blackout_cols = []
            for c_idx in range(1, max_cols + 1):
                col_header = str(ws.cell(start_row, c_idx).value or "").strip().lower()
                if any(kw in col_header for kw in ("單價", "price", "保固", "warranty", "總計", "total", "ttl")) and "內容" not in col_header and "備註" not in col_header and "remark" not in col_header:
                    blackout_cols.append(c_idx)
                    
            curr_row = start_row + 1
            while curr_row < start_row + len(rows) - 1: # exclude header and total row
                for c_idx in blackout_cols:
                    ws.cell(curr_row, c_idx).fill = black_fill
                    ws.cell(curr_row, c_idx).font = black_font
                curr_row += 1
                
            # Spacer row
            ws.row_dimensions[curr_row + 1].height = 15
            
            # 4. Remarks Section (A18 onwards)
            ws.row_dimensions[curr_row + 2].height = 18
            ws.cell(curr_row + 2, 1).value = rem_header
            ws.cell(curr_row + 2, 1).font = Font(name="Calibri", size=10, bold=True)
            ws.merge_cells(f"A{curr_row+2}:H{curr_row+2}")
            style_range(ws, f"A{curr_row+2}:H{curr_row+2}", alignment=Alignment(vertical="center", horizontal="left"))
            
            r_rem = curr_row + 3
            for line in rem_body:
                ws.row_dimensions[r_rem].height = 18
                ws.cell(r_rem, 1).value = line
                ws.merge_cells(f"A{r_rem}:H{r_rem}")
                style_range(ws, f"A{r_rem}:H{r_rem}", font=Font(name="Calibri", size=9), alignment=Alignment(vertical="center", horizontal="left", wrap_text=True))
                r_rem += 1
                
            # Spacer row
            ws.row_dimensions[r_rem].height = 25
            
            # Move cursor
            row_cursor = r_rem + 1
            
        pdf_doc.close()
        
        # Auto fit columns (excluding image/remarks rows width calculations)
        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            max_len = 0
            for table in tables:
                t_start = table.get("table_start_row", 9)
                t_end = t_start + len(table["rows"])
                for r_idx in range(t_start, t_end):
                    val = str(ws.cell(r_idx, col_idx).value or "")
                    if val and len(val) > max_len:
                        max_len = len(val)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 42)
            
        ws.freeze_panes = None
        
    else:
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

    # Clean up temp images
    if is_beehe:
        for p in temp_images:
            try:
                p.unlink()
            except Exception:
                pass

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
