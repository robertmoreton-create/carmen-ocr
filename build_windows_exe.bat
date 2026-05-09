@echo off
setlocal
cd /d "%~dp0"

REM Build a standalone Windows EXE (run once by IT/admin)
if not exist ".venv\Scripts\python.exe" (
  if exist ".venv" (
    echo Existing .venv is not a Windows virtual environment. Rebuilding it...
    rmdir /s /q ".venv"
  )
  python -m venv .venv
)

set "VENV_PY=.venv\Scripts\python.exe"
set "TEMPLATE_XLSX=Statements\Docs\Su - HSBC One #833 (PayMe).xlsx"

"%VENV_PY%" --version >nul 2>nul
if errorlevel 1 (
  echo Existing .venv points to a missing Python install. Rebuilding it...
  rmdir /s /q ".venv"
  python -m venv .venv
)

"%VENV_PY%" -m pip install --upgrade pip
if errorlevel 1 goto :fail

"%VENV_PY%" -m pip install -r requirements.txt
if errorlevel 1 goto :fail

if not exist "Statements\Incoming" mkdir "Statements\Incoming"
if not exist "Statements\Output" mkdir "Statements\Output"
if exist "%TEMPLATE_XLSX%" goto :template_ok
echo Missing template workbook: %TEMPLATE_XLSX%
goto :fail
:template_ok

"%VENV_PY%" -m PyInstaller --noconfirm --clean --windowed --name HSBCStatementOCR ^
  --collect-all rapidocr_onnxruntime ^
  --add-data "app_settings.json;." ^
  --add-data "%TEMPLATE_XLSX%;." ^
  hsbc_statement_desktop_app.py
if errorlevel 1 goto :fail

if not exist "dist\HSBCStatementOCR\Statements\Incoming" mkdir "dist\HSBCStatementOCR\Statements\Incoming"
if not exist "dist\HSBCStatementOCR\Statements\Output" mkdir "dist\HSBCStatementOCR\Statements\Output"
copy /Y "app_settings.json" "dist\HSBCStatementOCR\app_settings.json" >nul

echo.
echo Build complete.
echo EXE location: dist\HSBCStatementOCR\HSBCStatementOCR.exe
pause
goto :eof

:fail
echo.
echo Build failed.
pause
exit /b 1
