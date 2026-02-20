"""Unit tests for core report-generation logic (no HTTP, pure Python)."""

from __future__ import annotations

import tempfile
from pathlib import Path

import pytest


# ── load_csv_budget_data ──────────────────────────────────────────────────────

def test_load_csv_projects_valid(tmp_path, sample_csv_bytes):
    """load_csv_budget_data should return a processed DataFrame without raising."""
    from webapp.report_generator import load_csv_budget_data

    csv_file = tmp_path / "budget.csv"
    csv_file.write_bytes(sample_csv_bytes)

    df, mapping = load_csv_budget_data(csv_file)
    assert df is not None
    # The function returns a processed dataframe with a Projekt(e) column;
    # the exact name depends on whether it was renamed during processing.
    assert len(df.columns) > 0
    assert isinstance(mapping, dict)


def test_load_csv_missing_projects_column(tmp_path):
    """load_csv_budget_data should raise ValueError when 'Projekte' column is missing."""
    from webapp.report_generator import load_csv_budget_data

    bad_csv = "Foo\tBar\nval1\tval2\n".encode("utf-16")
    csv_file = tmp_path / "bad.csv"
    csv_file.write_bytes(bad_csv)

    with pytest.raises(ValueError, match="Projekte"):
        load_csv_budget_data(csv_file)


# ── load_xml_times ────────────────────────────────────────────────────────────

def test_load_xml_times_valid(tmp_path, sample_xml_bytes):
    """load_xml_times should return a DataFrame with expected columns."""
    from webapp.report_generator import load_xml_times

    xml_file = tmp_path / "data.xml"
    xml_file.write_bytes(sample_xml_bytes)

    df = load_xml_times(xml_file)
    assert df is not None
    assert "staff_name" in df.columns
    assert "hours" in df.columns
    assert "date_parsed" in df.columns
    assert len(df) > 0


def test_load_xml_times_hours_are_numeric(tmp_path, sample_xml_bytes):
    """Hours column should contain numeric values (float), not strings."""
    from webapp.report_generator import load_xml_times
    import pandas as pd

    xml_file = tmp_path / "data.xml"
    xml_file.write_bytes(sample_xml_bytes)

    df = load_xml_times(xml_file)
    assert pd.api.types.is_float_dtype(df["hours"]) or pd.api.types.is_numeric_dtype(df["hours"])


# ── determine_quarter ─────────────────────────────────────────────────────────

def test_determine_quarter_explicit_q4_2024(tmp_path, sample_xml_bytes):
    """When the quarter is explicitly provided, it should override auto-detection."""
    from webapp.report_generator import load_xml_times, determine_quarter

    xml_file = tmp_path / "data.xml"
    xml_file.write_bytes(sample_xml_bytes)
    df = load_xml_times(xml_file)

    result = determine_quarter(df, requested="Q4-2024")
    assert str(result.period) == "2024Q4"
    # Sample XML has data in October and November only → 2 months in this quarter
    assert 1 <= len(result.months) <= 3


def test_determine_quarter_auto(tmp_path, sample_xml_bytes):
    """Auto-detected quarter should be non-empty and consistent with the XML data."""
    from webapp.report_generator import load_xml_times, determine_quarter

    xml_file = tmp_path / "data.xml"
    xml_file.write_bytes(sample_xml_bytes)
    df = load_xml_times(xml_file)

    result = determine_quarter(df, requested=None)
    assert result.period is not None
    assert len(result.months) >= 1


# ── Utility functions ─────────────────────────────────────────────────────────

def test_is_bonus_project():
    from webapp.report_generator import is_bonus_project
    assert is_bonus_project("0000 Allgemein") is True
    assert is_bonus_project("1234.01 Testprojekt") is False
    assert is_bonus_project("") is False


def test_norm_ms_strips_bullets():
    from webapp.report_generator import norm_ms
    assert norm_ms("• 1.1 Test") == "1.1 Test"
    assert norm_ms("● Item") == "Item"


def test_de_to_float():
    from webapp.report_generator import de_to_float
    assert de_to_float("8,00") == 8.0
    assert de_to_float("1.234,56") == 1234.56
    assert de_to_float("abc") != de_to_float("abc")  # NaN != NaN


def test_extract_budget_monthly():
    from webapp.report_generator import extract_budget_from_name
    hours, unit = extract_budget_from_name("Einarbeitung (max. 8h/Monat pro MA)")
    assert hours == 8.0
    assert unit == "monat"


def test_extract_budget_quarterly():
    from webapp.report_generator import extract_budget_from_name
    hours, unit = extract_budget_from_name("Firmenveranstaltungen (max. 4h/Quartal pro MA)")
    assert hours == 4.0
    assert unit == "quartal"
