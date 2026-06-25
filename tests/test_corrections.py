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
