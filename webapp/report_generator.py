# -*- coding: utf-8 -*-
"""Core logic for generating quarterly bonus reports.

This module refactors the original CLI script into reusable functions that can
be imported by a web server or other automation. The main entry point is
`generate_quarterly_report`, which expects paths to the CSV and XML sources and
returns the path to the generated Excel workbook.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# ===================== BUDGETS FÜR 0000-PROJEKT =====================
# Diese Budgets gelten für ALLE Mitarbeiter, die diese Meilensteine bearbeiten
# Format: {"Meilensteinname": Stunden_pro_Monat/Quartal}

# Monatliche Budgets (0000-Meilensteine mit "Monat" im Namen)
MONTHLY_BUDGETS: Dict[str, float] = {
    "Einarbeitung neuer Mitarbeiter (max. 8h/Monat pro MA)": 8.0,
    "Angebote-Ausschreibungen-Kalkulationen (max. 8h/Monat pro MA)": 8.0,
    "Erstellung Vorlagen (übergreifend) (max. 8h/Monat pro MA)": 8.0,
}

# Quartalsbudgets (0000-Meilensteine mit "Quartal" im Namen)
QUARTERLY_BUDGETS: Dict[str, float] = {
    "Firmenveranstaltungen (max. 4h/Quartal pro MA)": 4.0,
    "Vorträge, Repräsentation (übergreifend) (max. 4h/Quartal pro MA)": 4.0,
    "Messeauftritt (max. 4h/Quartal pro MA)": 4.0,
}


ProgressCallback = Callable[[int, str], None]


def _noop_progress(_: int, __: str) -> None:
    """Default progress callback that swallows updates."""


# ===================== Hilfsfunktionen =====================
def de_to_float(x):
    if pd.isna(x):
        return np.nan
    s = str(x).strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return np.nan


def norm_ms(text: str) -> str:
    if text is None or (isinstance(text, float) and math.isnan(text)):
        return ""
    s = str(text).replace("\u2022", "").replace("•", "").replace("●", "")
    s = re.sub(r"^[\-\s]+", "", s)
    return s.strip()


def get_milestone_type(milestone_name: str) -> str:
    if milestone_name is None or (isinstance(milestone_name, float) and math.isnan(milestone_name)):
        return "monthly"
    name_lower = str(milestone_name).lower()
    return "quarterly" if "quartal" in name_lower else "monthly"


def extract_budget_from_name(ms_name):
    """Extrahiert Budgetstunden und Einheit (Monat/Quartal) aus dem Meilenstein-Namen."""

    if ms_name is None or (isinstance(ms_name, float) and math.isnan(ms_name)):
        return None, None
    text = str(ms_name)
    m = re.search(r"(?i)(\d+[\.,]?\d*)\s*h\s*(?:/|pro\s+)(monat|quartal)", text)
    if not m:
        return None, None
    try:
        hours = float(m.group(1).replace(',', '.'))
    except Exception:
        return None, None
    unit = m.group(2).lower()
    return hours, unit


def is_bonus_project(name: str) -> bool:
    if name is None or (isinstance(name, float) and math.isnan(name)):
        return False
    s = str(name).strip().lower()
    return s.startswith('0000')


def status_color_hex(p: float) -> str:
    if p < 90:
        return "C6EFCE"  # grün
    elif p <= 100:
        return "FFF2CC"  # gelb
    else:
        return "F8CBAD"  # rot


def load_csv_projects(csv_path: Path) -> pd.DataFrame:
    """CSV laden (Soll/Ist-Basis)."""

    try_encodings = [("utf-16", "\t"), ("utf-8-sig", "\t"), ("cp1252", "\t")]
    df = None
    for enc, delim in try_encodings:
        try:
            df = pd.read_csv(csv_path, delimiter=delim, encoding=enc)
            break
        except Exception:
            continue
    if df is None:
        raise RuntimeError("CSV konnte nicht gelesen werden.")

    df.columns = [c.strip().replace("\u200b", "").replace("\ufeff", "") for c in df.columns]
    df["Projekte"] = df["Projekte"].ffill()

    mask_ms = df["Arbeitspaket"].notna() & (df["Arbeitspaket"].astype(str).str.strip() != "-")
    cols_need = ["Projekte", "Arbeitspaket", "Iststunden", "Sollstunden Budget"]
    ms = df.loc[mask_ms, cols_need].copy()
    ms["Ist"] = ms["Iststunden"].map(de_to_float)
    ms["Soll"] = ms["Sollstunden Budget"].map(de_to_float)
    ms["Meilenstein"] = ms["Arbeitspaket"].map(norm_ms)
    ms = ms[["Projekte", "Meilenstein", "Ist", "Soll"]]

    g = ms.groupby(["Projekte", "Meilenstein"], as_index=False).agg({"Soll": "sum", "Ist": "sum"})
    g["Prozent"] = np.where(
        g["Soll"] > 0,
        (g["Ist"] / g["Soll"]) * 100.0,
        np.where(g["Ist"] > 0, 999.0, 0.0),
    )
    g["proj_norm"] = g["Projekte"].astype(str).str.strip()
    g["ms_norm"] = g["Meilenstein"].map(norm_ms)
    return g


def load_xml_times(xml_path: Path) -> pd.DataFrame:
    """XML laden (Zeiteinträge)."""

    tree = ET.parse(xml_path)
    root = tree.getroot()
    rows = []

    for element in root.iter():
        if element.tag != "row":
            continue
        cells = element.findall("./cell")
        if not cells:
            continue
        entry = {cell.attrib.get("name"): (cell.text or "").strip() for cell in cells}
        if entry and "staff_name" in entry and entry["staff_name"] and "work_package_name" in entry and "date" in entry:
            rows.append(entry)

    if not rows:
        raise ValueError("XML enthält keine Daten.")
    df = pd.DataFrame(rows)

    def extract_date(date_str):
        match = re.search(r'(\d{1,2})\s+(\w{3})\s+(\d{4})', date_str)
        if match:
            day, month_str, year = match.groups()
            months = {'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06',
                      'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'}
            month = months.get(month_str, '01')
            return pd.to_datetime(f"{year}-{month}-{day.zfill(2)}", errors='coerce')
        return pd.NaT

    df["date_parsed"] = df["date"].apply(extract_date)
    df["period"] = df["date_parsed"].dt.to_period("M")
    df["quarter"] = df["date_parsed"].dt.to_period("Q")
    df["proj_norm"] = df["project"].astype(str).str.strip()
    df["ms_norm"] = df["work_package_name"].map(norm_ms)

    def parse_hours(x):
        if pd.isna(x):
            return 0.0
        try:
            return float(str(x).strip())
        except Exception:
            return 0.0

    df['hours'] = df.get('number', 0).apply(parse_hours)
    return df


def list_available_quarters(df_xml: pd.DataFrame) -> Dict[pd.Period, List[pd.Period]]:
    """Gibt verfügbare Quartale und zugehörige Monate zurück."""

    quarters: Dict[pd.Period, List[pd.Period]] = {}
    for period in df_xml["period"].dropna().unique():
        quarter = period.to_timestamp().to_period("Q")
        quarters.setdefault(quarter, [])
        quarters[quarter].append(period)
    for quarter, months in quarters.items():
        months.sort()
    return quarters


def parse_quarter(quarter_str: str) -> pd.Period:
    """Parst Eingaben wie "2025Q3" oder "Q3-2025"."""

    quarter_str = quarter_str.strip().upper()
    match = re.match(r"^Q(\d)[-/]?\s*(\d{4})$", quarter_str)
    if match:
        q, year = match.groups()
        return pd.Period(year=int(year), quarter=int(q), freq="Q")
    match = re.match(r"^(\d{4})[-/\s]*Q(\d)$", quarter_str)
    if match:
        year, q = match.groups()
        return pd.Period(year=int(year), quarter=int(q), freq="Q")
    raise ValueError(f"Ungültiges Quartal: {quarter_str}")


@dataclass
class QuarterSelection:
    period: pd.Period
    months: List[pd.Period]


def determine_quarter(df_xml: pd.DataFrame, requested: Optional[str] = None) -> QuarterSelection:
    """Wählt das Zielquartal basierend auf den XML-Daten."""

    available = list_available_quarters(df_xml)
    if not available:
        raise ValueError("Keine Quartale in den XML-Daten gefunden.")

    if requested:
        target = parse_quarter(requested)
        if target not in available:
            raise ValueError(f"Angefordertes Quartal {requested} nicht in den XML-Daten enthalten.")
        months = sorted(available[target])
        return QuarterSelection(period=target, months=months)

    # Standard: jüngstes Quartal wählen
    target = sorted(available.keys())[-1]
    months = sorted(available[target])
    return QuarterSelection(period=target, months=months)


def build_quarterly_report(
    df_csv: pd.DataFrame,
    df_xml: pd.DataFrame,
    target_quarter: pd.Period,
    months: Iterable[pd.Period],
    out_path: Path,
    progress_cb: ProgressCallback = _noop_progress,
) -> Path:
    """Erstellt Quartals-Excel mit Monats-Tabellen + Quartalsübersicht."""

    df_quarter = df_xml[df_xml["quarter"] == target_quarter].copy()

    wb = Workbook()
    wb.remove(wb.active)
    thin = Side(style="thin", color="DDDDDD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    employees = sorted(df_quarter["staff_name"].unique())
    total_emps = max(len(employees), 1)

    for idx_emp, emp in enumerate(employees, start=1):
        ws = wb.create_sheet(title=emp[:31])
        ws.append([f"{emp} – Quartalsreport {target_quarter}"])
        ws.append([])

        current_row = 3
        total_hours_all_months = 0.0
        total_bonus_hours_quarter = 0.0

        for month in months:
            df_month = df_quarter[(df_quarter["period"] == month) & (df_quarter["staff_name"] == emp)].copy()

            if df_month.empty:
                continue

            month_hours = (
                df_month.groupby(['proj_norm', 'ms_norm'], as_index=False)
                .agg({'hours': 'sum'})
            )

            month_data = month_hours.merge(
                df_csv,
                how="left",
                left_on=["proj_norm", "ms_norm"],
                right_on=["proj_norm", "ms_norm"],
            )

            month_data["Projekte"] = month_data["Projekte"].fillna(month_data["proj_norm"])
            month_data["Meilenstein"] = month_data["Meilenstein"].fillna(month_data["ms_norm"])
            month_data["MeilensteinTyp"] = month_data["Meilenstein"].apply(get_milestone_type)

            month_data["Soll"] = month_data["Soll"].fillna(0.0)
            month_data["Ist"] = month_data["Ist"].fillna(0.0)

            for idx in month_data.index:
                if month_data.loc[idx, "Soll"] == 0.0 and month_data.loc[idx, "Ist"] == 0.0:
                    ms_name = month_data.loc[idx, "Meilenstein"]
                    ms_type = month_data.loc[idx, "MeilensteinTyp"]

                    if ms_type == "monthly" and ms_name in MONTHLY_BUDGETS:
                        month_data.loc[idx, "Soll"] = MONTHLY_BUDGETS[ms_name]
                        month_data.loc[idx, "Ist"] = month_data.loc[idx, "hours"]

            for idx in month_data.index:
                ms_name = month_data.loc[idx, "Meilenstein"]
                ms_type = month_data.loc[idx, "MeilensteinTyp"]
                proj_name = month_data.loc[idx, "Projekte"]
                proj_norm = month_data.loc[idx, "proj_norm"] if "proj_norm" in month_data.columns else ""
                if ms_type == "monthly" and ms_name in MONTHLY_BUDGETS and (is_bonus_project(proj_name) or is_bonus_project(proj_norm)):
                    month_data.loc[idx, "Soll"] = MONTHLY_BUDGETS[ms_name]
                    month_data.loc[idx, "Ist"] = month_data.loc[idx, "hours"]

            def _compute_month_qsoll(row):
                if row.get("MeilensteinTyp") != "quarterly":
                    return 0.0
                hours, unit = extract_budget_from_name(row.get("Meilenstein"))
                if unit == "quartal" and hours is not None:
                    return hours
                name = row.get("Meilenstein")
                if name in QUARTERLY_BUDGETS:
                    return QUARTERLY_BUDGETS[name]
                try:
                    if pd.notna(row.get("Soll")) and float(row.get("Soll")) > 0:
                        return float(row.get("Soll"))
                except Exception:
                    pass
                return 0.0

            month_data["QuartalsSoll"] = month_data.apply(_compute_month_qsoll, axis=1)

            df_to_date = df_quarter[(df_quarter["staff_name"] == emp) & (df_quarter["period"] <= month)]
            cum_hours_map = {
                (r["proj_norm"], r["ms_norm"]): r["hours"]
                for _, r in df_to_date.groupby(["proj_norm", "ms_norm"], as_index=False).agg({"hours": "sum"}).iterrows()
            }

            month_data = month_data.sort_values(["Projekte", "Meilenstein"])

            month_name_de = {
                'January': 'Januar', 'February': 'Februar', 'March': 'März',
                'April': 'April', 'May': 'Mai', 'June': 'Juni',
                'July': 'Juli', 'August': 'August', 'September': 'September',
                'October': 'Oktober', 'November': 'November', 'December': 'Dezember'
            }
            month_str = month.strftime('%B %Y')
            for en, de in month_name_de.items():
                month_str = month_str.replace(en, de)

            ws.append([f"--- {month_str} ---"])
            ws[f"A{current_row}"].font = Font(bold=True, size=12)
            current_row += 1

            ws.append(["Projekt", "Meilenstein", "Typ", "Soll (h)", "Ist (h)", f"{month_str} (h)", "%"])
            for cell in ws[current_row]:
                cell.font = Font(bold=True)
                cell.border = border
            current_row += 1

            bonus_hours_month = 0.0

            for proj, proj_block in month_data.groupby("Projekte", sort=False):
                proj_block = proj_block.reset_index(drop=True)
                block_start = current_row

                for i, (_, row_data) in enumerate(proj_block.iterrows()):
                    ms_type = row_data["MeilensteinTyp"]
                    typ_short = "Q" if ms_type == "quarterly" else "M"

                    if ms_type == "monthly":
                        soll = round(row_data["Soll"], 2)
                        ist = round(row_data["Ist"], 2)
                        prozent = (ist / soll * 100.0) if soll > 0 else 0.0
                        if prozent < 100.0:
                            bonus_hours_month += float(row_data["hours"]) or 0.0
                        ws.append([
                            proj if i == 0 else "",
                            row_data["Meilenstein"],
                            typ_short,
                            soll,
                            ist,
                            round(row_data["hours"], 2),
                            round(prozent, 2)
                        ])
                    else:
                        q_soll = float(row_data.get("QuartalsSoll", 0.0) or 0.0)
                        cum_ist = float(cum_hours_map.get((row_data["proj_norm"], row_data["ms_norm"]), 0.0))
                        prozent = (cum_ist / q_soll * 100.0) if q_soll > 0 else 0.0
                        if prozent < 100.0:
                            bonus_hours_month += float(row_data["hours"]) or 0.0
                        ws.append([
                            proj if i == 0 else "",
                            row_data["Meilenstein"],
                            typ_short,
                            round(q_soll, 2) if q_soll > 0 else "-",
                            round(cum_ist, 2) if cum_ist > 0 else 0.0,
                            round(row_data["hours"], 2),
                            round(prozent, 2) if q_soll > 0 else "-"
                        ])

                    for cell in ws[current_row]:
                        cell.border = border

                    pct_cell = ws.cell(row=current_row, column=7)
                    should_color = False
                    if ms_type == "monthly":
                        if row_data["Soll"] > 0:
                            prozent_col = (row_data["Ist"] / row_data["Soll"] * 100.0)
                            should_color = True
                        else:
                            prozent_col = 0.0
                    else:
                        q_soll_tmp = float(row_data.get("QuartalsSoll", 0.0) or 0.0)
                        if q_soll_tmp > 0:
                            cum_ist_tmp = float(cum_hours_map.get((row_data["proj_norm"], row_data["ms_norm"]), 0.0))
                            prozent_col = (cum_ist_tmp / q_soll_tmp * 100.0)
                            should_color = True
                        else:
                            prozent_col = 0.0
                    if should_color:
                        pct_cell.fill = PatternFill(
                            start_color=status_color_hex(prozent_col),
                            end_color=status_color_hex(prozent_col),
                            fill_type="solid",
                        )
                    current_row += 1

                block_size = len(proj_block)
                if block_size > 1:
                    ws.merge_cells(start_row=block_start, start_column=1,
                                   end_row=block_start + block_size - 1, end_column=1)
                    ws.cell(row=block_start, column=1).alignment = Alignment(vertical="top")

            sum_hours = month_data["hours"].sum()
            total_hours_all_months += sum_hours
            ws.append(["", "Summe", "", "", "", round(sum_hours, 2), ""])
            for cell in ws[current_row]:
                cell.font = Font(bold=True)
                cell.border = border
            current_row += 1

            ws.append(["", "Bonusberechtigte Stunden", "", "", "", round(bonus_hours_month, 2), ""])
            for cell in ws[current_row]:
                cell.font = Font(bold=True)
                cell.border = border
            current_row += 1

            total_bonus_hours_quarter += bonus_hours_month

            ws.append([])
            current_row += 1

        df_emp_quarter = df_quarter[df_quarter["staff_name"] == emp].copy()

        if df_emp_quarter.empty:
            continue

        quarter_hours = (
            df_emp_quarter.groupby(['proj_norm', 'ms_norm'], as_index=False)
            .agg({'hours': 'sum'})
        )

        quarter_data = quarter_hours.merge(
            df_csv,
            how="left",
            left_on=["proj_norm", "ms_norm"],
            right_on=["proj_norm", "ms_norm"],
        )

        quarter_data["Projekte"] = quarter_data["Projekte"].fillna(quarter_data["proj_norm"])
        quarter_data["Meilenstein"] = quarter_data["Meilenstein"].fillna(quarter_data["ms_norm"])
        quarter_data["MeilensteinTyp"] = quarter_data["Meilenstein"].apply(get_milestone_type)

        quarter_quarterly = quarter_data[quarter_data["MeilensteinTyp"] == "quarterly"].copy()

        def _compute_quarter_soll(row):
            hours, unit = extract_budget_from_name(row.get("Meilenstein"))
            if unit == "quartal" and hours is not None:
                return hours
            if row.get("Meilenstein") in QUARTERLY_BUDGETS:
                return QUARTERLY_BUDGETS[row.get("Meilenstein")]
            soll_val = row.get("Soll", np.nan)
            try:
                if pd.notna(soll_val) and float(soll_val) > 0:
                    return float(soll_val)
            except Exception:
                pass
            return 0.0

        quarter_quarterly["QuartalsSoll"] = quarter_quarterly.apply(_compute_quarter_soll, axis=1)

        if not quarter_quarterly.empty:
            ws.append([f"--- Quartalsübersicht {target_quarter} ---"])
            ws[f"A{current_row}"].font = Font(bold=True, size=12)
            current_row += 1

            ws.append(["Projekt", "Meilenstein", "Q-Soll (h)", "Q-Ist (h)", "%"])
            for cell in ws[current_row]:
                cell.font = Font(bold=True)
                cell.border = border
            current_row += 1

            for proj, proj_block in quarter_quarterly.groupby("Projekte", sort=False):
                proj_block = proj_block.reset_index(drop=True)
                block_start = current_row

                for i, (_, row_data) in enumerate(proj_block.iterrows()):
                    ms_name = row_data["Meilenstein"]
                    q_soll = row_data.get("QuartalsSoll", 0.0)
                    q_ist = row_data["hours"]
                    prozent = (q_ist / q_soll * 100.0) if q_soll > 0 else 0.0

                    ws.append([
                        proj if i == 0 else "",
                        ms_name,
                        round(q_soll, 2) if q_soll > 0 else "-",
                        round(q_ist, 2),
                        round(prozent, 2) if q_soll > 0 else "-"
                    ])

                    for cell in ws[current_row]:
                        cell.border = border

                    if q_soll > 0:
                        pct_cell = ws.cell(row=current_row, column=5)
                        pct_cell.fill = PatternFill(
                            start_color=status_color_hex(prozent),
                            end_color=status_color_hex(prozent),
                            fill_type="solid",
                        )
                    current_row += 1

                block_size = len(proj_block)
                if block_size > 1:
                    ws.merge_cells(start_row=block_start, start_column=1,
                                   end_row=block_start + block_size - 1, end_column=1)
                    ws.cell(row=block_start, column=1).alignment = Alignment(vertical="top")

        ws.append([])
        current_row += 1
        ws.append([f"--- Gesamtstunden {target_quarter} ---"])
        ws[f"A{current_row}"].font = Font(bold=True, size=12)
        current_row += 1
        ws.append(["Gesamt eingetragene Stunden:", round(total_hours_all_months, 2)])
        for cell in ws[current_row]:
            cell.font = Font(bold=True)
        current_row += 1

        ws.append(["Bonusberechtigte Stunden (Quartal):", round(total_bonus_hours_quarter, 2)])
        for cell in ws[current_row]:
            cell.font = Font(bold=True)
        current_row += 1

        ws.column_dimensions['A'].width = 40
        ws.column_dimensions['B'].width = 50
        ws.column_dimensions['C'].width = 8
        ws.column_dimensions['D'].width = 12
        ws.column_dimensions['E'].width = 12
        ws.column_dimensions['F'].width = 12
        ws.column_dimensions['G'].width = 8

        progress = int((idx_emp / total_emps) * 80) + 20
        progress_cb(min(progress, 95), f"Verarbeite Mitarbeiter {emp}")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path


def generate_quarterly_report(
    csv_path: Path,
    xml_path: Path,
    output_dir: Path,
    requested_quarter: Optional[str] = None,
    progress_cb: ProgressCallback = _noop_progress,
) -> Path:
    """Hauptfunktion: erzeugt den Bericht und gibt den Pfad zur Excel-Datei zurück."""

    progress_cb(5, "Lade CSV-Daten")
    df_csv = load_csv_projects(csv_path)

    progress_cb(10, "Lade XML-Daten")
    df_xml = load_xml_times(xml_path)

    selection = determine_quarter(df_xml, requested=requested_quarter)
    progress_cb(15, f"Wähle Quartal {selection.period}")

    year = selection.period.year
    quarter_num = selection.period.quarter
    out_path = output_dir / f"Q{quarter_num}-{year}.xlsx"

    result = build_quarterly_report(
        df_csv=df_csv,
        df_xml=df_xml,
        target_quarter=selection.period,
        months=selection.months,
        out_path=out_path,
        progress_cb=progress_cb,
    )

    progress_cb(100, "Fertig")
    return result


__all__ = [
    "generate_quarterly_report",
    "determine_quarter",
    "list_available_quarters",
    "parse_quarter",
    "MONTHLY_BUDGETS",
    "QUARTERLY_BUDGETS",
]
