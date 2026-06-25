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
