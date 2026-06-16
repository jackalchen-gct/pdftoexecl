from __future__ import annotations

import json
import sys
from pathlib import Path

import pdfplumber

sys.stdout.reconfigure(encoding="utf-8")


def summarize_pdf(path: Path) -> dict:
    pages = []
    with pdfplumber.open(path) as pdf:
        for index, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            tables = page.extract_tables() or []
            table_shapes = []
            previews = []
            for table in tables:
                rows = len(table)
                cols = max((len(row) for row in table), default=0)
                table_shapes.append({"rows": rows, "cols": cols})
                previews.append(table[:3])
            pages.append(
                {
                    "page": index,
                    "text_chars": len(text),
                    "table_count": len(tables),
                    "table_shapes": table_shapes,
                    "table_previews": previews,
                    "text_preview": text[:500],
                }
            )
    return {"file": path.name, "pages": pages}


def main() -> None:
    pdfs = sorted(Path(".").glob("*.pdf"))
    summaries = [summarize_pdf(path) for path in pdfs]
    print(json.dumps(summaries, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
