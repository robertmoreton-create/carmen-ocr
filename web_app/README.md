# Carmen OCR - Web App

Azure-hosted web version of Carmen OCR for processing bank and credit card statements.

## Quick Start

```bash
cd web_app
pip install -r requirements.txt
python app.py
```

Open http://localhost:5000 in your browser.

## Azure Deployment

### 1. Create Azure Resources (Bicep)

```bash
az deployment group create \
  --resource-group carmen-ocr-rg \
  --template-file infrastructure/main.bicep \
  --parameters environment=production
```

### 2. Configure Environment Variables

```bash
export AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT=https://your-resource.cognitiveservices.azure.com/
export AZURE_DOCUMENT_INTELLIGENCE_KEY=your-key-here
```

### 3. Deploy to Azure App Service

```bash
az webapp up \
  --name carmen-ocr-app \
  --resource-group carmen-ocr-rg \
  --plan carmen-ocr-plan \
  --runtime "PYTHON:3.11"
```

## Features

- Upload PDF statements via drag & drop
- Process HSBC and AMex statements
- Download extracted transactions as Excel
- Azure Document Intelligence for OCR

## Architecture

```
Browser → Azure App Service → Azure DI → Excel Download
```
# Trigger CI/CD deployment
# Trigger CI/CD
# Deployment trigger Sat  9 May 2026 21:50:51 HKT
