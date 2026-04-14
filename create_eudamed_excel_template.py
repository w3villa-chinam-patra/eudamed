#!/usr/bin/env python3
"""Create a user-fillable Excel template for EUDAMED XML generation."""

from __future__ import annotations

import argparse
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

from eudamed_constants import GUIDE_HEADERS, TEMPLATE_EXAMPLES, TEMPLATE_GUIDANCE_ROWS, TEMPLATE_SHEETS


HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)


SHEETS: dict[str, list[str]] = TEMPLATE_SHEETS
EXAMPLES: dict[str, dict[str, str]] = TEMPLATE_EXAMPLES
GUIDANCE_ROWS = TEMPLATE_GUIDANCE_ROWS


def build_workbook() -> openpyxl.Workbook:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for sheet_name, headers in SHEETS.items():
        ws = wb.create_sheet(sheet_name)
        for col, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = max(18, min(72, len(header) + 2))

            sample_value = EXAMPLES.get(sheet_name, {}).get(header, "")
            ws.cell(row=2, column=col, value=sample_value)

    guide = wb.create_sheet("Guidance")
    for col, header in enumerate(GUIDE_HEADERS, start=1):
        cell = guide.cell(row=1, column=col, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        guide.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 40

    for row, values in enumerate(GUIDANCE_ROWS, start=2):
        for col, value in enumerate(values, start=1):
            guide.cell(row=row, column=col, value=value)

    return wb


def main() -> None:
    parser = argparse.ArgumentParser(description="Create a fillable Excel template for EUDAMED XML generation.")
    parser.add_argument("--output", required=True, help="Output .xlsx path")
    args = parser.parse_args()

    output = Path(args.output)
    if not output.is_absolute():
        output = Path(__file__).resolve().parent / output

    output.parent.mkdir(parents=True, exist_ok=True)
    wb = build_workbook()
    wb.save(output)
    print(f"Template created: {output}")


if __name__ == "__main__":
    main()
