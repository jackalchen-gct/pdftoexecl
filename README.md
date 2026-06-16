# PDF to Excel Quote Converter

Tauri 2 + React 18 desktop app for converting quotation PDFs into Excel files.

## Current Scope

- Supports table-like quotation PDFs first.
- Verified with:
  - `Nvidia GH200 Pricing ODM 2023.8.23.pdf`
  - `比赫- 土耳其 CDU IB1350P 20260317.pdf`
- AMD text-only quotation parsing is intentionally deferred to a dedicated parser.

## Architecture

- `src/`: React 18 UI.
- `src-tauri/`: Tauri 2 shell and Rust commands.
- `converter/convert_pdf.py`: PDF extraction and Excel writer.

The React UI calls the Rust `convert_pdfs` command. Rust invokes the Python converter and returns per-file conversion results.

## Development Requirements

- Node.js and npm.
- Rust toolchain with Cargo, required for Tauri development/build.
- Python with `pdfplumber` and `openpyxl`.

You can bootstrap the development environment with:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\setup-dev.ps1
```

This script installs missing prerequisites with `winget`, creates `.venv`, installs Python packages, and runs `npm install`.

During development, the Rust command uses `python` by default. Set `PDFTOEXECL_PYTHON` if you need to point to a specific Python executable:

```powershell
$env:PDFTOEXECL_PYTHON="C:\Path\To\python.exe"
```

If you use the local virtual environment created by the setup script, the path is:

```powershell
$env:PDFTOEXECL_PYTHON="$PWD\.venv\Scripts\python.exe"
```

## Useful Commands

```powershell
npm install
npm run tauri dev
```

Run the converter directly:

```powershell
python converter\convert_pdf.py "Nvidia GH200 Pricing ODM 2023.8.23.pdf" "Nvidia GH200 Pricing ODM 2023.8.23.xlsx"
```

## Output

For `input.pdf`, the app writes `input.xlsx` to the selected output folder.

Each workbook currently includes:

- `Tables`: extracted quotation table data.
- `Text`: raw page text for audit/debugging.
