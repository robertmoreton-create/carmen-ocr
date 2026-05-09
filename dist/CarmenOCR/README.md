# Carmen OCR

Extract structured data from HSBC and American Express PDF statements and export to Excel.

## What It Does

- **HSBC Statements**: Reads PDF bank statements, extracts transactions (deposits, withdrawals, balances), and writes them to a formatted Excel workbook matching your Advent AXYS template.
- **AMex Statements**: Reads PDF credit card statements, extracts transactions (merchant, foreign spend, HKD amount), appends them to the AMEX.xlsx workbook, and queues new merchants for categorization.
- **Review Queue**: Built-in GUI viewer for uncategorized merchants with dropdowns to assign Category and Belongs To.
- **Auto-categorization**: Known merchants are automatically categorized from the Category lookup sheet.

## Architecture

```
PDF → Azure Document Intelligence (OCR) → Parser → Excel
```

Azure DI `prebuilt-layout` model extracts tables and text. The app maps the output to your existing Excel templates without changing formats.

## Requirements

### For Users (Pre-built .exe)
- Windows 10/11 (64-bit)
- Azure Document Intelligence endpoint + key (provided by your administrator)
- No Python, no admin rights needed

### For Developers
- Python 3.9+
- Dependencies: `openpyxl`, `azure-ai-formrecognizer`, `tkinter` (usually bundled)

Install dev dependencies:
```bash
pip install openpyxl azure-ai-formrecognizer
```

Build Windows .exe:
```bash
pip install pyinstaller
python build_windows.py
```

## Quick Start

1. **Download** `CarmenOCR.exe` and place it in a folder on your Desktop or Documents
2. **Create subfolders**:
   ```
   CarmenOCR/
   ├── CarmenOCR.exe
   ├── Statements/
   │   ├── Incoming/     ← drop HSBC PDFs here
   │   ├── Output/       ← Excel files appear here
   │   └── Docs/
   │       └── Su - HSBC One #833 (PayMe).xlsx  ← template
   └── Amex/
       ├── AMEX.xlsx     ← AMex workbook
       └── *.pdf         ← AMex statements
   ```
3. **Run** `CarmenOCR.exe`
4. **Enter Azure credentials** in the Settings tab (first run only)
5. **Process statements**:
   - HSBC tab → click "Process All PDFs"
   - AMex tab → click "Process All PDFs"
6. **Review Queue** tab → categorize any new merchants via dropdowns

## Configuration

Settings are stored in `Documents\CarmenOCR\carmen_settings.json` and `azure_config.json`. The app auto-creates these on first run. You can also set via environment variables:

```bash
set AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT=https://documentbank.cognitiveservices.azure.com/
set AZURE_DOCUMENT_INTELLIGENCE_KEY=your-key-here
```

## Security

- Azure key is stored locally in a JSON file (not hardcoded)
- Key input is masked in the GUI
- No registry writes, no admin required
- Config falls back to user's Documents folder if the .exe location is read-only
- Dev fallback keys disabled in production builds (requires `CARMEN_DEV=1` env var)

## Files

| File | Purpose |
|------|---------|
| `carmen_gui.py` | Main GUI application |
| `hsbc_statement_to_excel.py` | HSBC extraction engine |
| `azure_di_pipeline.py` | Azure DI → HSBC Excel pipeline |
| `amex_pipeline.py` | Azure DI → AMex Excel pipeline |
| `azure_ocr_backend.py` | Azure DI client wrapper |
| `build_windows.py` | PyInstaller build script |
| `WINDOWS_DEPLOY.md` | Detailed Windows deployment guide |

## License

Internal use only — proprietary to your organization.
