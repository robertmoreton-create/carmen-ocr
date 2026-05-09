@echo off
REM Build Carmen OCR for Windows (no admin required)
REM Run this on a Windows machine with Python installed

echo Checking Python...
python --version || (
    echo Python not found. Install from python.org first.
    exit /b 1
)

echo Installing dependencies...
pip install pyinstaller openpyxl azure-ai-formrecognizer

echo Building executable...
pyinstaller --onefile --windowed --name CarmenOCR --clean carmen_gui.py

echo.
echo Build complete!
echo Output: dist\CarmenOCR.exe
echo.
echo To distribute:
echo   1. Copy dist\CarmenOCR.exe to a USB stick or shared folder
echo   2. On target machine, place in e.g. C:\Users\%%USERNAME%%\Desktop\CarmenOCR\
echo   3. Double-click to run — no installation needed
echo   4. First run: enter Azure credentials in Settings tab
echo.
pause
