#!/bin/zsh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

VENV_DIR="$SCRIPT_DIR/.venv-mac"

pick_python() {
  for candidate in \
    "/Library/Frameworks/Python.framework/Versions/3.12/bin/python3" \
    "/Library/Frameworks/Python.framework/Versions/3.11/bin/python3" \
    "/opt/homebrew/bin/python3" \
    "$(command -v python3 2>/dev/null || true)"
  do
    if [ -n "$candidate" ] && [ -x "$candidate" ]; then
      echo "$candidate"
      return 0
    fi
  done
  return 1
}

PYTHON_SRC="$(pick_python)"
if [ -z "${PYTHON_SRC:-}" ]; then
  echo "No usable python3 found."
  exit 1
fi

if [ ! -x "$VENV_DIR/bin/python3" ]; then
  echo "Creating macOS virtual environment using: $PYTHON_SRC"
  "$PYTHON_SRC" -m venv "$VENV_DIR"
fi

source "$VENV_DIR/bin/activate"

python -m pip install --upgrade pip >/dev/null
python -m pip install openpyxl pypdf flask >/dev/null

mkdir -p "$SCRIPT_DIR/Statements/Incoming" "$SCRIPT_DIR/Statements/Output"

export HSBC_DISABLE_RAPIDOCR=1

echo "Launching HSBC Statement OCR app..."
if python - <<'PY'
import tkinter as tk
root = tk.Tk()
root.withdraw()
root.update_idletasks()
root.destroy()
print("tk_ok")
PY
then
  python "$SCRIPT_DIR/hsbc_statement_desktop_app.py"
else
  echo "GUI runtime not compatible on this macOS/Python build."
  echo "Starting browser GUI fallback at http://127.0.0.1:8765 ..."
  (sleep 1; open "http://127.0.0.1:8765" >/dev/null 2>&1 || true) &
  python "$SCRIPT_DIR/hsbc_mac_web.py"
fi
