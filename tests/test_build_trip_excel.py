"""Tests for the deterministic .xlsx renderer (build_trip_excel.py).

The script is importable as a module; the tests exercise its validation,
formatting helpers, and the three/four sheets it produces, without ever making
"what to plan" decisions (those belong to the model).
"""
import json
import sys
from pathlib import Path

import pytest

SCRIPTS_DIR = Path(__file__).resolve().parent.parent / "scripts"
EXAMPLES_DIR = Path(__file__).resolve().parent.parent / "examples"
sys.path.insert(0, str(SCRIPTS_DIR))

import build_trip_excel as bte  # noqa: E402

SAMPLE = json.loads((EXAMPLES_DIR / "trip_data.json").read_text(encoding="utf-8"))


def _build(trip):
    """Build a workbook from a trip dict without touching the filesystem."""
    wb = bte.Workbook()
    wb.remove(wb.active)
    if bte._meta_present(trip):
        bte._build_overview(wb, trip)
    bte._build_itinerary(wb, trip)
    bte._build_budget(wb, trip)
    bte._build_checklist(wb, trip)
    return wb


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def test_sanitize_name_strips_invalid_chars():
    assert bte._sanitize_name('Osaka/大阪:?*<>"|') == "Osaka 大阪"
    assert bte._sanitize_name("Osaka / 大阪") == "Osaka 大阪"
    assert bte._sanitize_name("   ") == "Trip"


def test_display_width_counts_cjk_double():
    assert bte._display_width("ab") == 2
    assert bte._display_width("大阪") == 4
    assert bte._display_width(None) == 0


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------
def test_validate_accepts_sample():
    assert bte._validate_trip(SAMPLE) == []


def test_validate_requires_top_level_fields():
    errors = bte._validate_trip({})
    missing = [e for e in errors if e.startswith("missing required field")]
    assert len(missing) == 5


def test_validate_rejects_non_positive_days():
    trip = dict(SAMPLE, days=0)
    assert any("positive integer" in e for e in bte._validate_trip(trip))


def test_validate_rejects_empty_itinerary():
    trip = dict(SAMPLE, daily_itinerary=[])
    assert any("daily_itinerary" in e for e in bte._validate_trip(trip))


def test_validate_rejects_slot_without_activity():
    trip = dict(SAMPLE, daily_itinerary=[{"day": 1, "slots": [{"slot": "09:00"}]}])
    assert any("empty 'activity'" in e for e in bte._validate_trip(trip))


# ---------------------------------------------------------------------------
# Workbook shape
# ---------------------------------------------------------------------------
def test_workbook_sheet_order_with_overview():
    wb = _build(SAMPLE)
    assert wb.sheetnames == ["Overview", "Daily Itinerary", "Budget", "Checklist"]


def test_workbook_omits_overview_without_meta():
    trip = {k: v for k, v in SAMPLE.items()
            if k not in ("dates", "travelers", "budget_style", "notes")}
    wb = _build(trip)
    assert "Overview" not in wb.sheetnames


def test_itinerary_header_and_rows():
    wb = _build(SAMPLE)
    ws = wb["Daily Itinerary"]
    assert ws.cell(row=1, column=1).value.startswith("天数")
    assert ws.freeze_panes == "A2"
    # 5 days * 4..5 slots each -> more rows than just the header.
    assert ws.max_row > 5


def test_budget_sum_formula_and_total_row():
    wb = _build(SAMPLE)
    ws = wb["Budget"]
    last_data_row = ws.max_row - 1  # final row is the total
    total = ws.cell(row=ws.max_row, column=3).value
    assert total == f"=SUM(C2:C{last_data_row})"


def test_string_amount_renders_without_breaking():
    trip = dict(SAMPLE)
    trip["budget"] = [{"category": "Transport", "item": "Flight",
                       "estimated": "3000/person", "actual": None, "notes": ""}]
    wb = _build(trip)  # must not raise
    assert wb["Budget"].cell(row=2, column=3).value == "3000/person"


def test_checklist_default_status():
    wb = _build(SAMPLE)
    ws = wb["Checklist"]
    assert all(ws.cell(row=r, column=3).value == "⬜ 待准备"
               for r in range(2, ws.max_row + 1))


# ---------------------------------------------------------------------------
# End-to-end CLI
# ---------------------------------------------------------------------------
def test_main_writes_file(tmp_path):
    out = bte.main([str(EXAMPLES_DIR / "trip_data.json"), "-o", str(tmp_path)])
    assert out == 0
    files = list(tmp_path.glob("*.xlsx"))
    assert len(files) == 1
    assert "Osaka" in files[0].name and "5D" in files[0].name


def test_main_returns_error_on_invalid_input(tmp_path):
    bad = tmp_path / "bad.json"
    bad.write_text('{"days": 5}', encoding="utf-8")
    assert bte.main([str(bad), "-o", str(tmp_path)]) == 1


if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-v"]))
