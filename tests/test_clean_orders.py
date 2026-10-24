from datetime import date
from pathlib import Path
import sys

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from clean_orders import clean, parse_date, slash_convention, summary, to_number  # noqa: E402
from make_sample import build  # noqa: E402


@pytest.mark.parametrize(
    "raw, expected",
    [
        (1250, 1250.0),
        ("1 250 ₸", 1250.0),
        ("1 250,50 тг", 1250.5),
        ("12 шт", 12.0),
        ("уточнить", None),
        (None, None),
    ],
)
def test_to_number(raw, expected):
    assert to_number(raw) == expected


def test_dot_dates_are_day_first():
    assert parse_date("03.02.2025", None) == (date(2025, 2, 3), None)


def test_slash_date_uses_sheet_convention():
    assert slash_convention(["03/25/2025", "03/04/2025"]) == "mdy"
    assert parse_date("03/04/2025", "mdy") == (date(2025, 3, 4), None)
    assert slash_convention(["25/03/2025"]) == "dmy"
    assert parse_date("03/04/2025", "dmy") == (date(2025, 4, 3), None)


def test_ambiguous_slash_date_is_flagged_not_hidden():
    d, problem = parse_date("03/04/2025", None)
    assert d is not None
    assert problem and "ambiguous" in problem


@pytest.fixture(scope="module")
def result(tmp_path_factory):
    path = tmp_path_factory.mktemp("data") / "messy.xlsx"
    build(path)
    return clean(path)


def test_notes_sheet_is_skipped(result):
    _, _, stats = result
    assert stats["sheets_skipped"] == ["notes"]


def test_dates_fall_in_the_month_of_their_sheet(result):
    data, _, _ = result
    months = {"Jan 2025": 1, "февраль": 2, "Mar-25": 3}
    for sheet, month in months.items():
        assert (data.loc[data["source_sheet"] == sheet, "date"].dt.month == month).all()


def test_no_totals_blanks_or_duplicates_left(result):
    data, _, stats = result
    assert stats["total_rows_removed"] == 3
    assert stats["duplicates_removed"] == 3
    assert data["order_id"].notna().all()
    assert not data.duplicated(subset=["order_id", "date", "sku", "qty"]).any()


def test_bad_amounts_are_flagged(result):
    data, issues, _ = result
    assert (issues["problem"] == "amount is not a number").sum() == 3
    assert issues["problem"].str.contains("≠").sum() == 3
    assert (data["amount_check"] != "ok").sum() == 6


def test_names_are_normalised(result):
    data, _, _ = result
    assert data["sku"].str.fullmatch(r"[A-Z]{3}-\d{3}").all()
    assert data.groupby("sku")["product"].nunique().max() == 1
    assert set(data["warehouse"]) == {"Kokshetau", "Astana", "Petropavl"}


def test_summary_matches_clean_rows(result):
    data, _, _ = result
    s = summary(data)
    ok_total = data.loc[data["amount_check"] == "ok", "amount"].sum()
    assert s["Total"].sum() == pytest.approx(ok_total)
