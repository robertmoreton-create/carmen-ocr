#!/usr/bin/env python3
from __future__ import annotations

import json
import os
import re
from datetime import date
from pathlib import Path

from flask import Flask, redirect, render_template_string, request, url_for

from hsbc_statement_to_excel import convert_statement

APP_DIR = Path(__file__).resolve().parent
SETTINGS_PATH = APP_DIR / "app_settings.json"

HTML = """
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>HSBC Statement OCR (Mac)</title>
  <style>
    body{font-family:Georgia,serif;margin:24px;background:#f8f6f2;color:#1d1b16}
    .card{background:white;max-width:920px;margin:auto;padding:20px;border:1px solid #d9d2c4;border-radius:10px}
    h1{margin-top:0}
    label{display:block;margin:10px 0 4px;font-weight:600}
    input,select{width:100%;padding:8px;border:1px solid #bfb7a8;border-radius:6px}
    button{margin-top:16px;padding:10px 14px;background:#164b7a;color:white;border:none;border-radius:6px;cursor:pointer}
    .ok{background:#eaf7ee;border:1px solid #9fd0ad;padding:10px;border-radius:6px;margin:12px 0}
    .err{background:#fdecec;border:1px solid #e2a5a5;padding:10px;border-radius:6px;margin:12px 0}
    .meta{font-size:12px;color:#5b5548;margin-top:8px}
  </style>
</head>
<body>
  <div class="card">
    <h1>HSBC Statement OCR (Mac)</h1>
    <form method="post" action="{{ url_for('run_convert') }}" enctype="multipart/form-data">
      <label>Input Folder</label>
      <div style="display:flex;gap:8px;align-items:center">
        <input type="text" name="input_folder" value="{{ settings['input_folder'] }}" />
        <a href="{{ browse_input_url }}"><button type="button">Select</button></a>
      </div>

      <label>Output Folder</label>
      <div style="display:flex;gap:8px;align-items:center">
        <input type="text" name="output_folder" value="{{ settings['output_folder'] }}" />
        <a href="{{ browse_output_url }}"><button type="button">Select</button></a>
      </div>

      <label>Template Workbook</label>
      <div style="display:flex;gap:8px;align-items:center">
        <input type="text" name="template_workbook" value="{{ settings['template_workbook'] }}" />
        <a href="{{ browse_template_url }}"><button type="button">Select</button></a>
      </div>

      <label>Template Sheet</label>
      <input type="text" name="template_sheet" value="{{ settings['template_sheet'] }}" />

      <label>Statement PDF from Input Folder</label>
      <select name="pdf_name">
        {% for name in pdfs %}
          <option value="{{ name }}">{{ name }}</option>
        {% endfor %}
      </select>

      <label>Or Choose a PDF File</label>
      <input type="file" name="pdf_upload" accept=".pdf,application/pdf" />

      <label>Or Enter Full PDF Path</label>
      <input type="text" name="pdf_path" value="" placeholder="/Users/robert/.../statement.pdf" />

      <label>Output Sheet Name</label>
      <input type="text" name="sheet_name" value="{{ default_sheet_name }}" />

      <button type="submit">Generate Excel</button>
    </form>

    {% if message %}
      <div class="{{ 'ok' if success else 'err' }}">{{ message }}</div>
    {% endif %}

    <div class="meta">Local-only. No data leaves your Mac.</div>
  </div>
</body>
</html>
"""

BROWSER_HTML = """
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Select Path</title>
  <style>
    body{font-family:Georgia,serif;margin:24px;background:#f8f6f2;color:#1d1b16}
    .card{background:white;max-width:980px;margin:auto;padding:20px;border:1px solid #d9d2c4;border-radius:10px}
    a{text-decoration:none;color:#164b7a}
    .row{padding:8px 0;border-top:1px solid #eee5d7}
    .muted{color:#6b6456;font-size:12px}
    .top{display:flex;justify-content:space-between;gap:12px;align-items:center}
    .btn{display:inline-block;padding:8px 12px;background:#164b7a;color:white;border-radius:6px}
  </style>
</head>
<body>
  <div class="card">
    <div class="top">
      <div>
        <h2 style="margin:0 0 8px 0">Select {{ label }}</h2>
        <div class="muted">{{ current_path }}</div>
      </div>
      <div><a class="btn" href="{{ url_for('home') }}">Back</a></div>
    </div>

    {% if allow_dir_select %}
      <p><a class="btn" href="{{ use_current_url }}">Use This Folder</a></p>
    {% else %}
      <p class="muted">Choose an `.xlsx` file below. Folders open for navigation.</p>
    {% endif %}

    {% if parent_url %}
      <div class="row"><a href="{{ parent_url }}">.. Parent Folder</a></div>
    {% endif %}

    {% for item in items %}
      <div class="row">
        {% if item.kind == 'dir' %}
          <a href="{{ item.open_url }}">[Folder] {{ item.name }}</a>
        {% else %}
          <a href="{{ item.select_url }}">[File] {{ item.name }}</a>
        {% endif %}
      </div>
    {% endfor %}
  </div>
</body>
</html>
"""


def default_settings() -> dict:
    return {
        "input_folder": "{APP_DIR}/Statements/Incoming",
        "output_folder": "{APP_DIR}/Statements/Output",
        "template_workbook": "{APP_DIR}/Statements/Docs/Su - HSBC One #833 (PayMe).xlsx",
        "template_sheet": "26",
        "sheet_name_format": "%b-%Y",
    }


def resolve_setting_path(raw: str) -> Path:
    value = (raw or "").strip().replace("\\", "/")
    value = value.replace("{APP_DIR}", str(APP_DIR))
    value = value.replace("{BUNDLE_DIR}", str(APP_DIR))
    return Path(value).expanduser()


def sanitize_sheet_name_for_file(value: str) -> str:
    cleaned = "".join(ch for ch in (value or "").strip() if ord(ch) >= 32)
    cleaned = re.sub(r'[<>:"/\\|?*]', "-", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    return cleaned[:64] or date.today().strftime("%b-%Y")


def sanitize_upload_name(value: str) -> str:
    name = Path((value or "").strip()).name
    cleaned = re.sub(r"[^A-Za-z0-9._ -]", "-", name)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    return cleaned or "statement.pdf"


def load_settings() -> dict:
    if not SETTINGS_PATH.exists():
        settings = default_settings()
        save_settings(settings)
        return settings
    with SETTINGS_PATH.open("r", encoding="utf-8") as f:
        settings = json.load(f)
    defaults = default_settings()
    for key, val in defaults.items():
        settings.setdefault(key, val)
    # Repair older packaged-path defaults when running locally on Mac.
    template_candidate = resolve_setting_path(settings.get("template_workbook", ""))
    local_template = APP_DIR / "Statements" / "Docs" / "Su - HSBC One #833 (PayMe).xlsx"
    if not template_candidate.exists() and local_template.exists():
        settings["template_workbook"] = str(local_template)
    return settings


def save_settings(settings: dict) -> None:
    with SETTINGS_PATH.open("w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2)


def update_single_setting(field: str, value: str) -> None:
    settings = load_settings()
    settings[field] = value
    save_settings(settings)


app = Flask(__name__)


def render_home(message: str = "", success: bool = True):
    settings = load_settings()
    in_dir = resolve_setting_path(settings["input_folder"])
    in_dir.mkdir(parents=True, exist_ok=True)
    pdfs = sorted([p.name for p in in_dir.glob("*.pdf")], reverse=True)
    if not pdfs:
        pdfs = [""]
    return render_template_string(
        HTML,
        settings=settings,
        pdfs=pdfs,
        browse_input_url=url_for("browse_path", field="input_folder", kind="dir"),
        browse_output_url=url_for("browse_path", field="output_folder", kind="dir"),
        browse_template_url=url_for("browse_path", field="template_workbook", kind="file"),
        default_sheet_name=date.today().strftime(settings.get("sheet_name_format", "%b-%Y")),
        message=message,
        success=success,
    )


@app.route("/", methods=["GET"])
def home():
    return render_home()


@app.route("/browse", methods=["GET"])
def browse_path():
    settings = load_settings()
    field = request.args.get("field", "").strip()
    kind = request.args.get("kind", "dir").strip()
    label_map = {
        "input_folder": "Input Folder",
        "output_folder": "Output Folder",
        "template_workbook": "Template Workbook",
    }
    if field not in label_map:
        return redirect(url_for("home"))

    raw_current = request.args.get("path", "").strip() or settings.get(field, "")
    current_path = resolve_setting_path(raw_current)
    if kind == "file":
        if current_path.is_file():
            current_path = current_path.parent
        elif current_path.suffix:
            current_path = current_path.parent
    if not current_path.exists():
        current_path = current_path.parent if current_path.parent.exists() else APP_DIR

    items = []
    try:
        children = sorted(current_path.iterdir(), key=lambda p: (not p.is_dir(), p.name.lower()))
    except Exception:
        children = []

    for child in children:
        if child.name.startswith("."):
            continue
        if child.is_dir():
            items.append(
                {
                    "kind": "dir",
                    "name": child.name,
                    "open_url": url_for("browse_path", field=field, kind=kind, path=str(child)),
                }
            )
        elif kind == "file" and child.suffix.lower() == ".xlsx":
            items.append(
                {
                    "kind": "file",
                    "name": child.name,
                    "select_url": url_for("select_path", field=field, value=str(child)),
                }
            )

    parent_url = None
    if current_path.parent != current_path:
        parent_url = url_for("browse_path", field=field, kind=kind, path=str(current_path.parent))

    use_current_url = url_for("select_path", field=field, value=str(current_path))
    return render_template_string(
        BROWSER_HTML,
        label=label_map[field],
        current_path=str(current_path),
        items=items,
        parent_url=parent_url,
        use_current_url=use_current_url,
        allow_dir_select=(kind == "dir"),
    )


@app.route("/select", methods=["GET"])
def select_path():
    field = request.args.get("field", "").strip()
    value = request.args.get("value", "").strip()
    if field in {"input_folder", "output_folder", "template_workbook"} and value:
        update_single_setting(field, value)
    return redirect(url_for("home"))


@app.route("/run", methods=["POST"])
def run_convert():
    settings = {
        "input_folder": request.form.get("input_folder", ""),
        "output_folder": request.form.get("output_folder", ""),
        "template_workbook": request.form.get("template_workbook", ""),
        "template_sheet": request.form.get("template_sheet", "26"),
        "sheet_name_format": load_settings().get("sheet_name_format", "%b-%Y"),
    }
    save_settings(settings)

    input_folder = resolve_setting_path(settings["input_folder"])
    output_folder = resolve_setting_path(settings["output_folder"])
    template_path = resolve_setting_path(settings["template_workbook"])

    pdf_name = request.form.get("pdf_name", "").strip()
    pdf_path_text = request.form.get("pdf_path", "").strip()
    sheet_name = request.form.get("sheet_name", "").strip() or date.today().strftime("%b-%Y")
    safe_sheet_name = sanitize_sheet_name_for_file(sheet_name)

    uploaded = request.files.get("pdf_upload")
    if pdf_path_text:
        pdf_path = Path(pdf_path_text).expanduser()
    elif uploaded and uploaded.filename:
        input_folder.mkdir(parents=True, exist_ok=True)
        uploaded_name = sanitize_upload_name(uploaded.filename)
        pdf_path = input_folder / uploaded_name
        uploaded.save(pdf_path)
    else:
        if not pdf_name:
            return render_home("No PDF found in input folder.", False)
        pdf_path = input_folder / pdf_name

    if not pdf_path.exists():
        return render_home(f"PDF not found: {pdf_path}", False)

    output_folder.mkdir(parents=True, exist_ok=True)
    output_path = output_folder / f"HSBC_{safe_sheet_name}.xlsx"

    try:
        os.environ["HSBC_DISABLE_RAPIDOCR"] = "1"
        result = convert_statement(
            pdf_path=pdf_path,
            template_path=template_path,
            output_path=output_path,
            template_sheet_name=settings["template_sheet"],
            sheet_name=sheet_name,
            order_mode="strict_pdf",
        )
        if not result.get("has_any_entries"):
            return render_home(
                "No transactions were extracted. This Mac is running without a working OCR engine for scanned PDFs yet. "
                "The app UI is fine, but scanned statements need OCR installed/enabled before conversion will work.",
                False,
            )
        return render_home(
            f"Done. Entries: {result['entries_written']}. Saved: {result['output_path']} (engine: {result.get('parser_engine')})",
            True,
        )
    except Exception as exc:
        return render_home(f"Failed: {exc}", False)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8765, debug=False)
