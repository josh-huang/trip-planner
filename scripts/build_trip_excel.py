#!/usr/bin/env python3
"""
Build a formatted, multi-sheet trip itinerary workbook (.xlsx).

Usage:
    python build_trip_excel.py <trip_data.json> [-o OUTPUT_DIR] [--version]

This script owns ONLY the deterministic formatting: sheet layout, header
styling, borders, auto column widths, row heights, freeze panes, the budget
SUM formula, and optional number formatting. Every "what to plan" decision
lives in `trip_data.json`, which the planner writes. Keeping data and
presentation separate means every invocation produces a visually consistent
workbook without re-deriving openpyxl styling.

Output file: `<destination>_Trip_Plan_<days>D.xlsx`.

Input schema (trip_data.json):
{
  "destination": "Osaka / 大阪",
  "days": 5,
  "currency": "¥",              // optional; "" disables the symbol
  "dates": "2026-08-01 ~ 2026-08-05",       // optional
  "travelers": "2 adults",                   // optional
  "budget_style": "Mid-range / comfortable", // optional
  "notes": "Bring a portable charger",       // optional
  "daily_itinerary": [
    {
      "day": 1,
      "slots": [
        {
          "slot": "08:30-09:00",
          "activity": "Leave hotel",
          "duration": "30min",
          "transit": "walk",
          "dining": "Breakfast: Kuromon Market",
          "tips": "Some stalls closed on Monday"
        }
      ]
    }
  ],
  "budget": [
    {"category": "Transport", "item": "Round-trip flight", "estimated": 3000,
     "actual": null, "notes": "tax included"},
    {"category": "Lodging", "item": "Hotel x4 nights", "estimated": 2400,
     "actual": null, "notes": ""}
  ],
  "checklist": [
    {"category": "Documents", "item": "Passport/ID", "status": "", "notes": "check expiry"}
  ]
}

`estimated`/`actual` may be a number (summed) or a string like "3000/person"
(rendered as text, not summed). `status` defaults to "⬜ To pack" when empty.

The optional top-level metadata keys (`dates`, `travelers`, `budget_style`,
`notes`, `currency`) render into an "Overview" sheet when at least one of the
text fields is present.
"""
from __future__ import annotations

import argparse
import json
import math
import os
import re
import sys
from typing import Any, Dict, Iterable, List, Optional, Set

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.worksheet import Worksheet
except ImportError:
    sys.exit("openpyxl is not installed. Run: pip install openpyxl")

__version__ = "1.0.0"


# ---------------------------------------------------------------------------
# Styling constants — the single source of truth for the workbook's look.
# ---------------------------------------------------------------------------
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")
FIELD_FONT = Font(bold=True, color="1F4E78")
TOTAL_FILL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
TOTAL_FONT = Font(bold=True, color="1F4E78")
THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def _display_width(value: Any) -> int:
    """Visual width of a cell value; CJK glyphs count double for column sizing."""
    if value is None:
        return 0
    return sum(2 if ord(ch) > 127 else 1 for ch in str(value))


def _autofit(ws: Worksheet, min_widths: Optional[Dict[int, int]] = None,
             max_width: int = 46) -> None:
    """Set each column width from its longest cell, clamped to a sensible range."""
    min_widths = min_widths or {}
    for col_cells in ws.columns:
        idx = col_cells[0].column
        letter = get_column_letter(idx)
        content = max((_display_width(c.value) for c in col_cells), default=0)
        width = min(max(content + 2, min_widths.get(idx, 8)), max_width)
        ws.column_dimensions[letter].width = width


def _style_header(ws: Worksheet, ncols: int) -> None:
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = THIN_BORDER
    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"


def _style_body(ws: Worksheet, align_map: Dict[int, Alignment]) -> None:
    """Border + alignment for every data row. align_map: {col_index: Alignment}."""
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = THIN_BORDER
            cell.alignment = align_map.get(cell.column, LEFT)


def _set_row_heights(ws: Worksheet, wrap_cols: Set[int], base: int = 15,
                     min_h: int = 22, max_h: int = 130) -> None:
    """Estimate row height so wrapped text is not visually clipped.

    Must run AFTER ``_autofit`` so the wrap-column widths used here are the
    final widths, not openpyxl's default 10.
    """
    for row in ws.iter_rows(min_row=2):
        lines = 1
        for cell in row:
            if cell.column in wrap_cols and cell.value:
                col_w = ws.column_dimensions[get_column_letter(cell.column)].width or 10
                lines = max(lines, math.ceil(_display_width(cell.value) / max(col_w - 2, 6)))
        ws.row_dimensions[row[0].row].height = min(max_h, max(min_h, lines * base))


def _append_row(ws: Worksheet, values: Iterable[Any]) -> None:
    ws.append([("" if v is None else v) for v in values])


def _sanitize_name(name: Any) -> str:
    """Strip characters invalid in a filename and collapse runs of whitespace."""
    name = re.sub(r'[\\/:*?"<>|]', " ", str(name))
    name = re.sub(r"\s+", " ", name).strip()
    return name or "Trip"


def _apply_number_format(cell, currency: str) -> None:
    """Apply a currency number format; harmless for string cells."""
    cell.number_format = f"{currency}#,##0" if currency else "#,##0"


# ---------------------------------------------------------------------------
# Input validation
# ---------------------------------------------------------------------------
_REQUIRED_TOP = ("destination", "days", "daily_itinerary", "budget", "checklist")
_BUDGET_KEYS = ("category", "item", "estimated", "actual", "notes")
_CHECKLIST_KEYS = ("category", "item", "status", "notes")


def _validate_trip(trip: Any) -> List[str]:
    """Return a list of human-readable problems; empty means the input is valid."""
    errors: List[str] = []
    if not isinstance(trip, dict):
        return ["root must be a JSON object"]

    for key in _REQUIRED_TOP:
        if key not in trip:
            errors.append(f"missing required field '{key}'")

    if "destination" in trip and not str(trip["destination"]).strip():
        errors.append("'destination' must be a non-empty string")

    if "days" in trip and not (isinstance(trip["days"], int) and trip["days"] > 0):
        errors.append("'days' must be a positive integer")

    itinerary = trip.get("daily_itinerary")
    if not isinstance(itinerary, list) or not itinerary:
        errors.append("'daily_itinerary' must be a non-empty list")
    else:
        for i, day in enumerate(itinerary, start=1):
            if not isinstance(day, dict) or not isinstance(day.get("day"), int):
                errors.append(f"daily_itinerary[{i}] must be an object with an integer 'day'")
                continue
            slots = day.get("slots")
            if not isinstance(slots, list) or not slots:
                errors.append(f"daily_itinerary[{i}] (day {day.get('day')}) has no 'slots'")
                continue
            for j, slot in enumerate(slots, start=1):
                if not isinstance(slot, dict):
                    errors.append(f"daily_itinerary[{i}].slots[{j}] must be an object")
                elif not str(slot.get("activity", "")).strip():
                    errors.append(
                        f"daily_itinerary[{i}].slots[{j}] (day {day.get('day')}) "
                        "has an empty 'activity'")

    budget = trip.get("budget")
    if not isinstance(budget, list):
        errors.append("'budget' must be a list")
    else:
        for i, item in enumerate(budget, start=1):
            if not isinstance(item, dict):
                errors.append(f"budget[{i}] must be an object")

    checklist = trip.get("checklist")
    if not isinstance(checklist, list):
        errors.append("'checklist' must be a list")
    else:
        for i, item in enumerate(checklist, start=1):
            if not isinstance(item, dict):
                errors.append(f"checklist[{i}] must be an object")

    return errors


# ---------------------------------------------------------------------------
# Sheet builders
# ---------------------------------------------------------------------------
def _meta_present(trip: Dict[str, Any]) -> bool:
    return any(str(trip.get(k, "")).strip() for k in ("dates", "travelers", "budget_style", "notes"))


def _build_overview(wb: Workbook, trip: Dict[str, Any]) -> None:
    ws = wb.create_sheet("Overview", 0)
    _append_row(ws, ["Field", "Value"])
    rows = [
        ("Destination", trip.get("destination", "")),
        ("Duration", f"{trip.get('days', 1)} day(s)"),
        ("Dates", trip.get("dates", "")),
        ("Travelers", trip.get("travelers", "")),
        ("Budget style", trip.get("budget_style", "")),
        ("Notes", trip.get("notes", "")),
    ]
    for label, value in rows:
        if str(value).strip():
            _append_row(ws, [label, value])
    _style_header(ws, 2)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = THIN_BORDER
            cell.alignment = LEFT
        row[0].font = FIELD_FONT
    _autofit(ws, min_widths={1: 16})
    _set_row_heights(ws, wrap_cols={2})


def _build_itinerary(wb: Workbook, trip: Dict[str, Any]) -> None:
    ws = wb.create_sheet("Daily Itinerary")
    headers = ["天数 (Day)", "时间段 (Slot)", "项目/景点 (Activity)", "游玩时长 (Duration)",
               "交通方式 (Transit)", "特色美食/备选 (Dining/Notes)", "预约/避坑提醒 (Tips)"]
    _append_row(ws, headers)
    for day in trip.get("daily_itinerary", []):
        for slot in day.get("slots", []):
            _append_row(ws, [
                day.get("day", ""),
                slot.get("slot", ""),
                slot.get("activity", ""),
                slot.get("duration", ""),
                slot.get("transit", ""),
                slot.get("dining", ""),
                slot.get("tips", ""),
            ])
    _style_header(ws, len(headers))
    _style_body(ws, {1: CENTER, 2: CENTER, 4: CENTER, 5: CENTER})
    # Column widths must be finalised before row heights are estimated.
    _autofit(ws, min_widths={1: 8, 2: 16, 4: 14, 5: 14})
    _set_row_heights(ws, wrap_cols={3, 6, 7})


def _build_budget(wb: Workbook, trip: Dict[str, Any]) -> None:
    ws = wb.create_sheet("Budget")
    headers = ["费用类别 (Category)", "项目明细 (Item)", "预估金额 (Estimated)",
               "实际支出 (Actual)", "备注说明 (Notes)"]
    _append_row(ws, headers)
    for item in trip.get("budget", []):
        _append_row(ws, [
            item.get("category", ""),
            item.get("item", ""),
            item.get("estimated", ""),
            item.get("actual", ""),
            item.get("notes", ""),
        ])

    currency = str(trip.get("currency", "¥"))
    last_data_row = ws.max_row
    total_row = last_data_row + 1
    # With no budget rows, C2:C1 would include the total cell itself.
    estimated_total = f"=SUM(C2:C{last_data_row})" if last_data_row >= 2 else 0
    actual_total = f"=SUM(D2:D{last_data_row})" if last_data_row >= 2 else 0
    _append_row(ws, ["合计 (Total)", "", estimated_total, actual_total, ""])

    _style_header(ws, len(headers))
    _style_body(ws, {1: CENTER, 3: CENTER, 4: CENTER})
    # Number-format the amount columns and the total row's SUM cells.
    for row in ws.iter_rows(min_row=2, min_col=3, max_col=4):
        for cell in row:
            _apply_number_format(cell, currency)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=total_row, column=c)
        cell.fill = TOTAL_FILL
        cell.font = TOTAL_FONT
        cell.border = THIN_BORDER
        cell.alignment = CENTER
    _autofit(ws, min_widths={1: 18, 3: 14, 4: 14})
    _set_row_heights(ws, wrap_cols={2, 5})


def _build_checklist(wb: Workbook, trip: Dict[str, Any]) -> None:
    ws = wb.create_sheet("Checklist")
    headers = ["分类 (Category)", "必备物品 (Item)", "准备状态 (Status)", "备注 (Notes)"]
    _append_row(ws, headers)
    for item in trip.get("checklist", []):
        _append_row(ws, [
            item.get("category", ""),
            item.get("item", ""),
            item.get("status", "") or "⬜ 待准备",
            item.get("notes", ""),
        ])
    _style_header(ws, len(headers))
    _style_body(ws, {1: CENTER, 3: CENTER})
    _autofit(ws, min_widths={1: 14, 3: 14})
    _set_row_heights(ws, wrap_cols={2, 4})


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def _parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build a formatted multi-sheet trip itinerary workbook (.xlsx).",
        epilog="Example: python build_trip_excel.py trip_data.json -o .",
    )
    parser.add_argument("trip_data", help="path to the trip_data.json input file")
    parser.add_argument("-o", "--out-dir", default=".",
                        help="directory to write the .xlsx into (default: current dir)")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> int:
    # Windows console may default to cp1252 and choke on non-ASCII paths in
    # print(); switch to UTF-8 so the final absolute-path line never crashes.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass

    args = _parse_args(argv)

    with open(args.trip_data, encoding="utf-8") as f:
        try:
            trip = json.load(f)
        except json.JSONDecodeError as exc:
            print(f"error: {args.trip_data} is not valid JSON: {exc}", file=sys.stderr)
            return 1

    problems = _validate_trip(trip)
    if problems:
        print("error: invalid trip_data:", file=sys.stderr)
        for p in problems:
            print(f"  - {p}", file=sys.stderr)
        return 1

    os.makedirs(args.out_dir, exist_ok=True)

    wb = Workbook()
    wb.remove(wb.active)  # drop the default empty "Sheet"; each builder creates its own
    if _meta_present(trip):
        _build_overview(wb, trip)
    _build_itinerary(wb, trip)
    _build_budget(wb, trip)
    _build_checklist(wb, trip)

    destination = _sanitize_name(trip.get("destination", "Trip"))
    days = trip.get("days", 1)
    out_name = f"{destination}_Trip_Plan_{days}D.xlsx"
    out_path = os.path.abspath(os.path.join(args.out_dir, out_name))
    wb.save(out_path)
    print(out_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
