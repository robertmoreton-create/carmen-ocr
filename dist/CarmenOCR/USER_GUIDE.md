# Carmen OCR — User Guide

## What This App Does

Carmen OCR reads your bank and credit card PDF statements and puts the transactions into Excel spreadsheets automatically. No more manual typing.

### Supported Statements
- **HSBC One** (Hong Kong) — savings/current account statements
- **American Express** — credit card statements

### What You Get
- HSBC transactions in the same Excel format you already use for Advent AXYS
- AMex transactions appended to your existing AMEX.xlsx workbook
- New merchants automatically queued for categorization

## Before You Start

You need:
1. A Windows PC (Windows 10 or 11)
2. The `CarmenOCR.exe` file
3. An **Azure Document Intelligence key** from your IT team
4. **Two Excel spreadsheets** (the app cannot run without these):
   - `Su - HSBC One #833 (PayMe).xlsx` — your HSBC template
   - `AMEX.xlsx` — your AMex workbook

No Python, no installation, no admin rights needed.

## First-Time Setup (5 Minutes)

### 1. Create Your Folder Structure

Create a folder on your PC called `CarmenOCR`. Inside it, create these subfolders:

```
CarmenOCR/
├── CarmenOCR.exe                    ← the app
├── Statements/
│   ├── Incoming/                    ← put HSBC PDFs here
│   ├── Output/                      ← Excel files appear here
│   └── Docs/
│       └── Su - HSBC One #833 (PayMe).xlsx   ← your HSBC template
└── Amex/
    ├── AMEX.xlsx                    ← your AMex workbook
    └── (put AMex PDFs here)
```

**Important**: The app will not work without the two spreadsheets above. Place them in the correct folders before running the app.

### 2. Run the App

Double-click `CarmenOCR.exe`. You will see a window with 4 tabs.

### 3. Enter Azure Credentials

1. Click the **Settings** tab
2. Enter the **Endpoint** (from your IT team):
   ```
   https://documentbank.cognitiveservices.azure.com/
   ```
3. Enter the **Azure Key** (from your IT team)
4. Click **Save Azure Credentials**
5. Click **Test Connection** to verify it works

Your credentials are saved. You won't need to enter them again.

### 4. Set Your File Paths

Still in the **Settings** tab:

- **HSBC Input Folder**: browse to `CarmenOCR\Statements\Incoming`
- **HSBC Output Folder**: browse to `CarmenOCR\Statements\Output`
- **HSBC Template**: browse to `CarmenOCR\Statements\Docs\Su - HSBC One #833 (PayMe).xlsx`
- **AMex Workbook**: browse to `CarmenOCR\Amex\AMEX.xlsx`
- **AMex PDF Folder**: browse to `CarmenOCR\Amex`

Click **Save All Settings**.

## Daily Workflow

### Processing HSBC Statements

1. Drop your HSBC PDF into `Statements\Incoming`
2. Open Carmen OCR, click the **HSBC Statements** tab
3. Click **Refresh** to see the PDF in the list
4. Select the PDF and click **Process Selected PDF**
   (or click **Process All PDFs** to do them all at once)
5. The Excel file appears in `Statements\Output`

### Processing AMex Statements

1. Drop your AMex PDF into `Amex\`
2. Open Carmen OCR, click the **AMex Statements** tab
3. Click **Refresh** to see the PDF
4. Click **Process Selected PDF** or **Process All PDFs**
5. Transactions are appended to `AMEX.xlsx`

### Categorizing New Merchants

1. Click the **Review Queue** tab
2. You'll see a table of uncategorized merchants
3. **Double-click** the **Category** cell → a dropdown appears
4. Select the correct category (e.g. Dining out, Travel, Groceries)
5. **Double-click** the **Belongs To** cell → select who it belongs to
6. Click **Save Changes**
7. The categories flow through to your AMex Transactions sheet automatically

## Tips

- The app remembers your settings between runs
- If a merchant appears again in a future statement, it will be auto-categorized
- The Review Queue only shows merchants that need attention — not already categorized ones
- You can process multiple PDFs at once with **Process All PDFs**

## Troubleshooting

| Problem | Solution |
|---------|----------|
| "Azure connection failed" | Check your endpoint and key in Settings. Click Test Connection. |
| "No PDFs found" | Check the folder path in Settings. Make sure PDFs are in the right folder. |
| "Template not found" | Browse to the correct template file in Settings. |
| "AMEX.xlsx not found" | Make sure the AMex workbook is in the Amex folder. |
| App won't start | Make sure you're on Windows 10/11 64-bit. |
| Categories not saving | Make sure you click Save Changes in the Review Queue tab. |

## Where Your Data Is Stored

| File | Location | What It Contains |
|------|----------|-------------------|
| Settings | `CarmenOCR\Config\carmen_settings.json` | Folder paths, template locations |
| Azure Key | `CarmenOCR\Config\azure_config.json` | Your Azure endpoint and key |
| HSBC Output | `CarmenOCR\Statements\Output\` | Generated Excel files |
| AMex Workbook | `CarmenOCR\Amex\AMEX.xlsx` | Updated with new transactions |

## Getting Help

Contact your IT team for:
- Azure key issues
- New template formats
- App updates
