# Installation for Developers

## macOS / Linux

```bash
cd CarmenOCR
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python carmen_gui.py
```

## Windows

```cmd
cd CarmenOCR
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python carmen_gui.py
```

## Build Windows .exe

```bash
pip install pyinstaller
python build_windows.py
```

Output: `dist/CarmenOCR.exe`
