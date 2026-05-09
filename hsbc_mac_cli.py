#!/usr/bin/env python3
from pathlib import Path
from datetime import date

from hsbc_statement_to_excel import convert_statement


def ask_path(prompt: str, default: Path) -> Path:
    raw = input(f"{prompt} [{default}]: ").strip()
    return Path(raw) if raw else default


def ask_text(prompt: str, default: str) -> str:
    raw = input(f"{prompt} [{default}]: ").strip()
    return raw or default


def main() -> None:
    base = Path(__file__).resolve().parent
    default_pdf = next((base / "Statements" / "Incoming").glob("*.pdf"), None)
    default_pdf = default_pdf or (base / "Statements" / "Incoming" / "statement.pdf")
    default_template = base / "Statements" / "Docs" / "Su - HSBC One #833 (PayMe).xlsx"
    if not default_template.exists():
        alt = base / "Su - HSBC One #833 (PayMe).xlsx"
        if alt.exists():
            default_template = alt

    sheet_name = date.today().strftime("%b-%Y")
    default_output = base / "Statements" / "Output" / f"HSBC_{sheet_name}.xlsx"

    print("\nHSBC OCR Mac Test (CLI fallback)\n")
    pdf_path = ask_path("Statement PDF path", default_pdf)
    template_path = ask_path("Template workbook path", default_template)
    output_path = ask_path("Output workbook path", default_output)
    template_sheet = ask_text("Template sheet", "26")
    sheet_name = ask_text("Output sheet name", sheet_name)

    output_path.parent.mkdir(parents=True, exist_ok=True)

    result = convert_statement(
        pdf_path=pdf_path,
        template_path=template_path,
        output_path=output_path,
        template_sheet_name=template_sheet,
        sheet_name=sheet_name,
        order_mode="strict_pdf",
    )

    print("\nDone")
    print(f"Parser engine: {result.get('parser_engine')}")
    print(f"Entries written: {result.get('entries_written')}")
    print(f"Saved: {result.get('output_path')}")


if __name__ == "__main__":
    main()
