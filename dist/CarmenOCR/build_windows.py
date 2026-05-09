#!/usr/bin/env python3
"""Build script for Carmen OCR Windows executable.

Usage:
    python build_windows.py

Requirements:
    pip install pyinstaller openpyxl azure-ai-formrecognizer
"""

import subprocess
import sys
import shutil
from pathlib import Path


def main():
    print("Building Carmen OCR for Windows...")
    
    # Check pyinstaller
    if not shutil.which("pyinstaller"):
        print("Installing pyinstaller...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
    
    # Build
    cmd = [
        "pyinstaller",
        "--onefile",
        "--windowed",
        "--name", "CarmenOCR",
        "--clean",
        "--add-data", "azure_ocr_backend.py;.",
        "--add-data", "azure_di_pipeline.py;.",
        "--add-data", "amex_pipeline.py;.",
        "--add-data", "hsbc_statement_to_excel.py;.",
        "--add-data", "test_azure.py;.",
        "carmen_gui.py",
    ]
    
    print(f"Running: {' '.join(cmd)}")
    subprocess.run(cmd, check=True)
    
    print("\nBuild complete!")
    print("Output: dist/CarmenOCR.exe")
    print("\nTo distribute:")
    print("  1. Copy dist/CarmenOCR.exe to USB/shared folder")
    print("  2. Place on target machine in e.g. Desktop/CarmenOCR/")
    print("  3. Double-click to run — no admin, no installation")
    print("  4. First run: enter Azure credentials in Settings tab")


if __name__ == "__main__":
    main()
