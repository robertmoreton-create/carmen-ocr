# Carmen OCR — Windows Deployment Guide

## For Users (No Admin Required)

### What You Need
- Windows 10 or 11 (64-bit)
- Azure Document Intelligence API key (provided by IT)
- CarmenOCR.exe (single file, ~30MB)

### Setup (2 Minutes)

1. **Create a folder** on your Desktop:
   ```
   C:\Users\YourName\Desktop\CarmenOCR
   ```

2. **Copy CarmenOCR.exe** into that folder

3. **Copy your files** into subfolders:
   ```
   CarmenOCR/
   ├── CarmenOCR.exe          ← the app
   ├── Statements/
   │   ├── Incoming/          ← drop HSBC PDFs here
   │   ├── Output/            ← Excel files appear here
   │   └── Docs/
   │       └── Su - HSBC One #833 (PayMe).xlsx  ← template
   └── Amex/
       ├── AMEX.xlsx          ← AMex workbook
       └── *.pdf              ← AMex statements
   ```

4. **Double-click CarmenOCR.exe**

5. **Enter Azure credentials** in the Settings tab:
   - Endpoint: `https://documentbank.cognitiveservices.azure.com/`
   - Key: [provided by IT]
   - Click "Save Azure Credentials"

6. **Start processing**:
   - HSBC tab: click "Process All PDFs"
   - AMex tab: click "Process All PDFs"

### Where Files Go

| What | Where |
|------|-------|
| Settings | `Documents\CarmenOCR\carmen_settings.json` |
| Azure key | `Documents\CarmenOCR\azure_config.json` |
| HSBC output | `Desktop\CarmenOCR\Statements\Output\` |
| AMex output | appended to `Desktop\CarmenOCR\Amex\AMEX.xlsx` |

### No Admin? No Problem

- No installation wizard
- No registry changes
- No Program Files writes
- No Python needed on target machine
- Just a single .exe + your data folders

## For IT / Builder

### Build the .exe

On a Windows machine with Python:

```bash
pip install pyinstaller openpyxl azure-ai-formrecognizer
python build_windows.py
```

Output: `dist/CarmenOCR.exe`

### Size
- CarmenOCR.exe: ~25-35 MB (includes Python runtime + libraries)

### Dependencies Bundled
- Python 3.11+
- tkinter (GUI)
- openpyxl (Excel)
- azure-ai-formrecognizer (OCR)
- All project .py files

### Updating
To update: replace CarmenOCR.exe with new version. Settings and config files are preserved in Documents\CarmenOCR\.
