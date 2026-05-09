# Carmen OCR — Root Folder Installation Guide

## For IT Administrators

This app is designed to run from a root-level folder on a Windows PC (e.g., `C:\CarmenOCR\`). No admin rights are required for users after initial setup.

## Folder Structure

Create the following structure on the target PC:

```
C:\CarmenOCR\
├── CarmenOCR.exe              ← the application
├── Statements\
│   ├── Incoming\              ← users drop HSBC PDFs here
│   ├── Output\                ← Excel files generated here
│   └── Docs\
│       └── Su - HSBC One #833 (PayMe).xlsx  ← HSBC template
├── Amex\
│   ├── AMEX.xlsx             ← AMex master workbook
│   └── *.pdf                 ← AMex statements
└── Config\
    ├── azure_config.json      ← Azure credentials (auto-created)
    └── carmen_settings.json   ← app settings (auto-created)
```

## User Permissions

| Folder | User Permission | Purpose |
|--------|----------------|---------|
| `C:\CarmenOCR\` | Read | Application files |
| `C:\CarmenOCR\Statements\Incoming` | Read/Write | Drop PDFs |
| `C:\CarmenOCR\Statements\Output` | Read/Write | Retrieve Excel |
| `C:\CarmenOCR\Amex` | Read/Write | AMex PDFs + workbook |
| `C:\CarmenOCR\Config` | Read/Write | Settings + credentials |

## First-Time Setup

1. Copy `CarmenOCR.exe` and the folder structure to `C:\CarmenOCR\`
2. Provide the user with:
   - Azure Document Intelligence **endpoint**
   - Azure Document Intelligence **key**
3. User launches app, enters credentials in **Settings** tab
4. User clicks **Test Connection** to verify
5. User sets folder paths (pre-filled if structure matches above)

## Azure Firewall

If using IP allowlisting on Azure Document Intelligence:
1. Ask user to click **Show My IP** in Settings tab
2. Add that IP to the Azure resource firewall
3. User clicks **Test Connection** again

## Updating the App

To update:
1. Replace `CarmenOCR.exe` with new version
2. Config files in `Config\` are preserved
3. User settings remain intact

## Troubleshooting

| Issue | Cause | Fix |
|-------|-------|-----|
| "Access denied" | User lacks write permission | Grant write access to `Config\`, `Statements\`, `Amex\` |
| "Azure connection failed" | Wrong key or IP not allowlisted | Verify key, check IP firewall |
| "Template not found" | Path changed | Browse to correct template in Settings |

## Security Notes

- Azure key stored in `Config\azure_config.json` (plain text)
- Consider file-level permissions on `Config\` folder
- Key is masked in GUI input field
- No registry writes, no system changes
