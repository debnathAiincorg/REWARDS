import json
import os
import sys
from datetime import date, datetime, timedelta
from unittest.mock import patch

import pytest
from openpyxl import Workbook

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
import weekly_report_send_teams as module

CATEGORIES = [
    "Punctuality", "L&D", "Fluency Compliance",
    "Innovation", "Extraordinary Performance",
]


def make_test_workbook(rows):
    """Build an in-memory openpyxl Workbook with the expected sheet structure.

    rows: list of dicts with keys: date (datetime), name (str),
          and optionally each category name (int, default 0).
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Daily Performance Bonus"
    headers = ["Index", "Date", "Name"] + CATEGORIES
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)
    for r, row in enumerate(rows, 2):
        ws.cell(row=r, column=1, value=r - 1)
        ws.cell(row=r, column=2, value=row["date"])
        ws.cell(row=r, column=3, value=row["name"])
        for col, cat in enumerate(CATEGORIES, 4):
            ws.cell(row=r, column=col, value=row.get(cat, 0))
    return wb


# ── _detect_columns ──────────────────────────────────────────────────────────

def test_detect_columns_identifies_all_columns():
    wb = make_test_workbook([])
    ws = wb["Daily Performance Bonus"]
    date_col, name_col, category_cols, point_cols = module._detect_columns(ws)
    assert date_col == 2
    assert name_col == 3
    assert category_cols == {
        "Punctuality": 4, "L&D": 5, "Fluency Compliance": 6,
        "Innovation": 7, "Extraordinary Performance": 8,
    }
    assert sorted(point_cols) == [4, 5, 6, 7, 8]


def test_detect_columns_returns_none_for_missing_columns():
    wb = Workbook()
    ws = wb.active
    ws.cell(row=1, column=1, value="Index")
    date_col, name_col, category_cols, point_cols = module._detect_columns(ws)
    assert date_col is None
    assert name_col is None


# ── get_real_yesterday_data ───────────────────────────────────────────────────

def test_get_real_yesterday_data_returns_only_yesterday(tmp_path):
    yesterday = datetime.combine(date.today() - timedelta(days=1), datetime.min.time())
    today_dt = datetime.combine(date.today(), datetime.min.time())

    wb = make_test_workbook([
        {"date": yesterday, "name": "Alice",
         "Punctuality": 1, "L&D": 1, "Fluency Compliance": 0,
         "Innovation": 0, "Extraordinary Performance": 0},
        {"date": today_dt, "name": "Alice",
         "Punctuality": 1, "L&D": 1, "Fluency Compliance": 1,
         "Innovation": 1, "Extraordinary Performance": 1},
    ])
    test_file = str(tmp_path / "source.xlsx")
    wb.save(test_file)

    with patch.object(module, "TEMP_FILE", test_file):
        result = module.get_real_yesterday_data()

    assert list(result.keys()) == ["Alice"]
    assert result["Alice"] == {
        "Punctuality": 1, "L&D": 1, "Fluency Compliance": 0,
        "Innovation": 0, "Extraordinary Performance": 0,
    }


def test_get_real_yesterday_data_deduplicates_keeps_last(tmp_path):
    yesterday = datetime.combine(date.today() - timedelta(days=1), datetime.min.time())

    wb = make_test_workbook([
        {"date": yesterday, "name": "Bob", "Punctuality": 0, "L&D": 0,
         "Fluency Compliance": 0, "Innovation": 0, "Extraordinary Performance": 0},
        {"date": yesterday, "name": "Bob", "Punctuality": 1, "L&D": 1,
         "Fluency Compliance": 1, "Innovation": 1, "Extraordinary Performance": 1},
    ])
    test_file = str(tmp_path / "source.xlsx")
    wb.save(test_file)

    with patch.object(module, "TEMP_FILE", test_file):
        result = module.get_real_yesterday_data()

    assert result["Bob"]["Punctuality"] == 1


def test_get_real_yesterday_data_empty_when_no_rows(tmp_path):
    wb = make_test_workbook([])
    test_file = str(tmp_path / "source.xlsx")
    wb.save(test_file)

    with patch.object(module, "TEMP_FILE", test_file):
        result = module.get_real_yesterday_data()

    assert result == {}


# ── Snapshot functions ────────────────────────────────────────────────────────

def test_save_and_load_roundtrip(tmp_path):
    data = {"Alice": {"Punctuality": 1, "L&D": 0, "Fluency Compliance": 1,
                      "Innovation": 0, "Extraordinary Performance": 0}}
    snap_file = str(tmp_path / "snap.json")
    with patch.object(module, "SNAPSHOT_FILE", snap_file):
        module.save_snapshot("2026-06-24", data)
        result = module.load_snapshot()
    assert result == {"date": "2026-06-24", "employees": data}


def test_load_snapshot_returns_none_when_file_missing(tmp_path):
    with patch.object(module, "SNAPSHOT_FILE", str(tmp_path / "missing.json")):
        assert module.load_snapshot() is None


def test_save_snapshot_silent_on_bad_path():
    with patch.object(module, "SNAPSHOT_FILE", "/nonexistent_dir/snap.json"):
        module.save_snapshot("2026-06-24", {})  # must not raise


# ── has_yesterday_data_changed ────────────────────────────────────────────────

def test_no_change_when_data_identical(tmp_path):
    data = {"Alice": {"Punctuality": 1, "L&D": 0, "Fluency Compliance": 1,
                      "Innovation": 0, "Extraordinary Performance": 0}}
    snap_file = str(tmp_path / "snap.json")
    with open(snap_file, "w") as f:
        json.dump({"date": "2026-06-24", "employees": data}, f)
    with patch.object(module, "SNAPSHOT_FILE", snap_file):
        assert not module.has_yesterday_data_changed("2026-06-24", data)


def test_change_detected_when_value_differs(tmp_path):
    old = {"Alice": {"Punctuality": 0, "L&D": 0, "Fluency Compliance": 0,
                     "Innovation": 0, "Extraordinary Performance": 0}}
    new = {"Alice": {"Punctuality": 1, "L&D": 0, "Fluency Compliance": 0,
                     "Innovation": 0, "Extraordinary Performance": 0}}
    snap_file = str(tmp_path / "snap.json")
    with open(snap_file, "w") as f:
        json.dump({"date": "2026-06-24", "employees": old}, f)
    with patch.object(module, "SNAPSHOT_FILE", snap_file):
        assert module.has_yesterday_data_changed("2026-06-24", new)


def test_no_change_when_snapshot_date_differs(tmp_path):
    snapshot_data = {"Alice": {"Punctuality": 1, "L&D": 0, "Fluency Compliance": 0,
                               "Innovation": 0, "Extraordinary Performance": 0}}
    fresh_data = {"Alice": {"Punctuality": 0, "L&D": 0, "Fluency Compliance": 0,
                            "Innovation": 0, "Extraordinary Performance": 0}}
    snap_file = str(tmp_path / "snap.json")
    with open(snap_file, "w") as f:
        json.dump({"date": "2026-06-23", "employees": snapshot_data}, f)
    with patch.object(module, "SNAPSHOT_FILE", snap_file):
        # Different date means watcher fired on a new day — not a correction
        assert not module.has_yesterday_data_changed("2026-06-24", fresh_data)


def test_no_change_when_no_snapshot(tmp_path):
    with patch.object(module, "SNAPSHOT_FILE", str(tmp_path / "missing.json")):
        assert not module.has_yesterday_data_changed("2026-06-24", {"Alice": {}})


# ── Dynamic column detection ──────────────────────────────────────────────────

CATEGORIES_PLUS = CATEGORIES + ["Teamwork"]


def make_test_workbook_extra(rows):
    """Like make_test_workbook but with an extra 'Teamwork' column."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Daily Performance Bonus"
    headers = ["Index", "Date", "Name"] + CATEGORIES_PLUS
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)
    for r, row in enumerate(rows, 2):
        ws.cell(row=r, column=1, value=r - 1)
        ws.cell(row=r, column=2, value=row["date"])
        ws.cell(row=r, column=3, value=row["name"])
        for col, cat in enumerate(CATEGORIES_PLUS, 4):
            ws.cell(row=r, column=col, value=row.get(cat, 0))
    return wb


def test_detect_columns_includes_extra_column():
    wb = make_test_workbook_extra([])
    ws = wb["Daily Performance Bonus"]
    date_col, name_col, category_cols, point_cols = module._detect_columns(ws)
    assert "Teamwork" in category_cols
    assert category_cols["Teamwork"] == 9  # 9th column: Index, Date, Name + 5 + Teamwork
    assert sorted(point_cols) == [4, 5, 6, 7, 8, 9]


def test_get_real_yesterday_data_includes_extra_column(tmp_path):
    yesterday = datetime.combine(date.today() - timedelta(days=1), datetime.min.time())
    wb = make_test_workbook_extra([
        {"date": yesterday, "name": "Carol",
         "Punctuality": 1, "L&D": 0, "Fluency Compliance": 1,
         "Innovation": 0, "Extraordinary Performance": 0, "Teamwork": 1},
    ])
    test_file = str(tmp_path / "source.xlsx")
    wb.save(test_file)

    with patch.object(module, "TEMP_FILE", test_file):
        result = module.get_real_yesterday_data()

    assert "Carol" in result
    assert result["Carol"]["Teamwork"] == 1


# ── Notes / free-text column exclusion ────────────────────────────────────────

def make_test_workbook_with_notes(rows):
    """Like make_test_workbook but with an extra free-text 'Notes' column."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Daily Performance Bonus"
    headers = ["Index", "Date", "Name"] + CATEGORIES + ["Notes"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)
    for r, row in enumerate(rows, 2):
        ws.cell(row=r, column=1, value=r - 1)
        ws.cell(row=r, column=2, value=row["date"])
        ws.cell(row=r, column=3, value=row["name"])
        for col, cat in enumerate(CATEGORIES, 4):
            ws.cell(row=r, column=col, value=row.get(cat, 0))
        ws.cell(row=r, column=len(headers), value=row.get("Notes", ""))
    return wb


def test_detect_columns_excludes_notes_column():
    wb = make_test_workbook_with_notes([])
    ws = wb["Daily Performance Bonus"]
    date_col, name_col, category_cols, point_cols = module._detect_columns(ws)
    assert "Notes" not in category_cols
    assert sorted(point_cols) == [4, 5, 6, 7, 8]  # Notes column (9) excluded


def test_detect_columns_excludes_notes_case_insensitive():
    wb = Workbook()
    ws = wb.active
    ws.title = "Daily Performance Bonus"
    headers = ["Index", "Date", "Name"] + CATEGORIES + ["NOTES"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)
    date_col, name_col, category_cols, point_cols = module._detect_columns(ws)
    assert "NOTES" not in category_cols
    assert sorted(point_cols) == [4, 5, 6, 7, 8]


def test_get_cumulative_data_ignores_notes_column_text(tmp_path):
    """A free-text Notes column with a value like 'Washroom work+ Ticket' must
    not crash get_cumulative_data — it should simply be excluded as a point column."""
    yesterday_dt = datetime.combine(date.today() - timedelta(days=1), datetime.min.time())
    wb = make_test_workbook_with_notes([
        {"date": yesterday_dt, "name": "Dave",
         "Punctuality": 1, "L&D": 1, "Fluency Compliance": 1,
         "Innovation": 0, "Extraordinary Performance": 0,
         "Notes": "Washroom work+ Ticket"},
    ])
    test_file = str(tmp_path / "source.xlsx")
    wb.save(test_file)

    with patch.object(module, "TEMP_FILE", test_file):
        week_label, employees, prev_day_breakdown = module.get_cumulative_data()

    assert week_label is not None
    dave = next(e for e in employees if e["name"] == "Dave")
    assert dave["points"] == 3


# ── _safe_int defense-in-depth ────────────────────────────────────────────────

def test_safe_int_returns_zero_for_none_and_empty():
    assert module._safe_int(None) == 0
    assert module._safe_int("") == 0


def test_safe_int_converts_valid_numbers():
    assert module._safe_int(5) == 5
    assert module._safe_int("3") == 3


def test_safe_int_returns_zero_and_warns_on_bad_value(capsys):
    result = module._safe_int("Washroom work+ Ticket", context="row 2, Dave, Punctuality")
    assert result == 0
    captured = capsys.readouterr()
    assert "[WARNING]" in captured.out
    assert "Washroom work+ Ticket" in captured.out
    assert "row 2, Dave, Punctuality" in captured.out


def test_get_cumulative_data_stray_text_in_point_column_treated_as_zero(capsys, tmp_path):
    """If a point column somehow contains stray text (not caught by column
    exclusion), the script must not crash and should treat it as 0 with a warning."""
    yesterday_dt = datetime.combine(date.today() - timedelta(days=1), datetime.min.time())
    wb = make_test_workbook([
        {"date": yesterday_dt, "name": "Erin",
         "Punctuality": 1, "L&D": 1, "Fluency Compliance": 0,
         "Innovation": 0, "Extraordinary Performance": 0},
    ])
    ws = wb["Daily Performance Bonus"]
    # Corrupt a point column cell with free text, simulating misentered data.
    ws.cell(row=2, column=6, value="Washroom work+ Ticket")  # Fluency Compliance column
    test_file = str(tmp_path / "source.xlsx")
    wb.save(test_file)

    with patch.object(module, "TEMP_FILE", test_file):
        week_label, employees, prev_day_breakdown = module.get_cumulative_data()

    erin = next(e for e in employees if e["name"] == "Erin")
    assert erin["points"] == 2  # Punctuality(1) + L&D(1) + Fluency Compliance(treated as 0)

    captured = capsys.readouterr()
    assert "[WARNING]" in captured.out
    assert "Washroom work+ Ticket" in captured.out


def test_get_real_yesterday_data_stray_text_treated_as_zero(capsys, tmp_path):
    yesterday = datetime.combine(date.today() - timedelta(days=1), datetime.min.time())
    wb = make_test_workbook([
        {"date": yesterday, "name": "Frank",
         "Punctuality": 1, "L&D": 0, "Fluency Compliance": 0,
         "Innovation": 0, "Extraordinary Performance": 0},
    ])
    ws = wb["Daily Performance Bonus"]
    ws.cell(row=2, column=5, value="stray text")  # L&D column
    test_file = str(tmp_path / "source.xlsx")
    wb.save(test_file)

    with patch.object(module, "TEMP_FILE", test_file):
        result = module.get_real_yesterday_data()

    assert result["Frank"]["L&D"] == 0
    captured = capsys.readouterr()
    assert "[WARNING]" in captured.out
    assert "stray text" in captured.out


def test_format_prev_day_card_includes_extra_category():
    prev_day_breakdown = {
        "date_label": "Jun 26",
        "employees": [
            {"name": "Alice", "Punctuality": 1, "L&D": 0, "Fluency Compliance": 1,
             "Innovation": 0, "Extraordinary Performance": 0, "Teamwork": 1},
        ]
    }
    card = module.format_prev_day_card(prev_day_breakdown)
    assert card is not None

    # The header ColumnSet is the third element in body (index 2)
    header_columnset = card["body"][2]
    header_texts = [
        col["items"][0]["text"]
        for col in header_columnset["columns"]
    ]
    assert "Teamwork" in header_texts
