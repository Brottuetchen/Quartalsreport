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
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation

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


MONTH_NAMES = {
    1: "Januar",
    2: "Februar",
    3: "März",
    4: "April",
    5: "Mai",
    6: "Juni",
    7: "Juli",
    8: "August",
    9: "September",
    10: "Oktober",
    11: "November",
    12: "Dezember",
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


def detect_billing_type(arbeitspaket: str, honorarbereich: str) -> str:
    """
    Erkennt die Abrechnungsart eines Projekts/Meilensteins.
    Returns: "Pauschale" | "Nachweis" | "Unbekannt"
    """
    if pd.isna(honorarbereich) or str(honorarbereich).strip().upper() != "X":
        return "Unbekannt"  # Keine Obermeilenstein-Markierung

    if pd.isna(arbeitspaket):
        return "Unbekannt"

    arbeitspaket_lower = str(arbeitspaket).lower()

    # Pauschale erkennen: (p)
    if "(p)" in arbeitspaket_lower:
        return "Pauschale"

    # Nachweis erkennen: (aN), (a.N.), (a N)
    if any(marker in arbeitspaket_lower for marker in ["(an)", "(a.n.)", "(a n)"]):
        return "Nachweis"

    # Nicht eindeutig erkennbar
    return "Unbekannt"


def load_csv_budget_data(csv_path: Path) -> tuple[pd.DataFrame, dict]:
    """
    Lädt Budget-Informationen aus CSV für Projekt-Budget-Übersicht.
    Returns:
        - DataFrame mit Projekten, Abrechnungsart und Budget-Daten (pro Obermeilenstein)
        - Dictionary mapping (Projekt, work_package_name) -> Obermeilenstein
    """
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

    # Nur Obermeilensteine (X-Markierung) interessieren uns für Budget-Übersicht
    mask_obermeilenstein = (
        df["Honorarbereich"].notna() &
        (df["Honorarbereich"].astype(str).str.strip().str.upper() == "X") &
        df["Arbeitspaket"].notna() &
        (df["Arbeitspaket"].astype(str).str.strip() != "-")
    )

    budget_rows = []
    # Mapping: (Projekt, work_package_name) -> Obermeilenstein
    # This maps sub-items to their parent Obermeilenstein
    work_package_to_obermeilenstein = {}

    current_projekt = None
    current_obermeilenstein = None

    for idx, row in df.iterrows():
        projekt = str(row["Projekte"]).strip() if pd.notna(row["Projekte"]) else current_projekt
        arbeitspaket = str(row["Arbeitspaket"]).strip() if pd.notna(row["Arbeitspaket"]) else ""
        honorarbereich = str(row["Honorarbereich"]).strip() if pd.notna(row["Honorarbereich"]) else ""

        # Check if this is an Obermeilenstein
        is_obermeilenstein = (
            honorarbereich.upper() == "X" and
            arbeitspaket and
            arbeitspaket != "-"
        )

        if is_obermeilenstein:
            current_projekt = projekt
            current_obermeilenstein = arbeitspaket

            # Abrechnungsart erkennen
            billing_type = detect_billing_type(arbeitspaket, row["Honorarbereich"])

            # Budget-Daten extrahieren
            sollhonor = de_to_float(row.get("Sollhonorar", 0))
            verrechnete_honorare = de_to_float(row.get("Verrechnete Honorare", 0))
            istkosten = de_to_float(row.get("Istkosten", 0))
            sollstunden = de_to_float(row.get("Sollstunden Budget", 0))
            iststunden = de_to_float(row.get("Iststunden", 0))
            budget = de_to_float(row.get("Budget", 0))

            # Stundensätze für Positionen sammeln (aus Unterpositionen)
            rate_sv = None
            rate_cad = None
            rate_adm = None

            # Suche nach Unterpositionen mit SV/CAD/ADM
            # Nächste Zeilen nach dem Obermeilenstein durchsuchen
            for sub_idx in range(idx + 1, min(idx + 20, len(df))):
                sub_row = df.iloc[sub_idx]
                sub_arbeitspaket = str(sub_row.get("Arbeitspaket", ""))

                # Prüfe ob es eine Unterposition ist (beginnt mit " und enthält ')
                if not sub_arbeitspaket.startswith('"'):
                    break  # Nächster Obermeilenstein erreicht

                # Extrahiere Position
                if "'   SV" in sub_arbeitspaket or "'   S V" in sub_arbeitspaket:
                    sub_budget = de_to_float(sub_row.get("Budget", 0))
                    sub_sollstunden = de_to_float(sub_row.get("Sollstunden Budget", 0))
                    if sub_sollstunden > 0:
                        rate_sv = sub_budget / sub_sollstunden
                elif "'   CAD" in sub_arbeitspaket or "'   C A D" in sub_arbeitspaket:
                    sub_budget = de_to_float(sub_row.get("Budget", 0))
                    sub_sollstunden = de_to_float(sub_row.get("Sollstunden Budget", 0))
                    if sub_sollstunden > 0:
                        rate_cad = sub_budget / sub_sollstunden
                elif "'   ADM" in sub_arbeitspaket or "'   A D M" in sub_arbeitspaket:
                    sub_budget = de_to_float(sub_row.get("Budget", 0))
                    sub_sollstunden = de_to_float(sub_row.get("Sollstunden Budget", 0))
                    if sub_sollstunden > 0:
                        rate_adm = sub_budget / sub_sollstunden

            # Bei Pauschale: Berechne Stundensatz aus Sollhonor / Sollstunden
            if billing_type == "Pauschale" and sollstunden > 0:
                default_rate = sollhonor / sollstunden
                if rate_sv is None:
                    rate_sv = default_rate
                if rate_cad is None:
                    rate_cad = default_rate
                if rate_adm is None:
                    rate_adm = default_rate

            budget_rows.append({
                "Projekt": projekt,
                "Obermeilenstein": arbeitspaket,
                "Abrechnungsart": billing_type,
                "Gesamtbudget": sollhonor,
                "Abgerechnet": verrechnete_honorare,
                "Istkosten": istkosten,
                "Sollstunden": sollstunden,
                "Iststunden": iststunden,
                "Stundensatz_SV": rate_sv,
                "Stundensatz_CAD": rate_cad,
                "Stundensatz_ADM": rate_adm,
            })
        elif current_obermeilenstein and arbeitspaket and arbeitspaket != "-":
            # This is a sub-item (work package) under the current Obermeilenstein
            # Clean up the work package name (remove leading quotes and special chars)
            clean_arbeitspaket = arbeitspaket.lstrip('"').lstrip("'").strip()

            # Store mapping: (Projekt, work_package) -> Obermeilenstein
            if current_projekt and clean_arbeitspaket:
                # Store with full work package name
                work_package_to_obermeilenstein[(current_projekt, clean_arbeitspaket)] = current_obermeilenstein

                # Also store with just the numeric prefix (e.g., "3100" from "3100-Erarbeiten...")
                if "-" in clean_arbeitspaket:
                    prefix = clean_arbeitspaket.split("-")[0].strip()
                    work_package_to_obermeilenstein[(current_projekt, prefix)] = current_obermeilenstein

    return pd.DataFrame(budget_rows), work_package_to_obermeilenstein


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


def _create_project_budget_sheet(
    wb: Workbook,
    df_budget: pd.DataFrame,
    border: Border,
) -> None:
    """Creates a project budget overview sheet with billing type and hourly rates."""

    ws = wb.create_sheet(title="Projekt-Budget-Übersicht", index=0)

    # Title
    ws.append(["Projekt-Budget-Übersicht"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([])
    ws.append(["Hinweis: Rote Zellen = Manuelle Eingabe erforderlich | Gelbe Zellen = Optional manuell anpassen"])
    ws["A3"].font = Font(italic=True, size=10)
    ws.append([])

    current_row = 5

    # Header row
    headers = [
        "Projekt",
        "Obermeilenstein",
        "Abrechnungsart",
        "Status",
        "Gesamtbudget (€)",
        "Abgerechnet (€)",
        "Verfügbar (€)",
        "Stundensatz SV (€/h)",
        "Stundensatz CAD (€/h)",
        "Stundensatz ADM (€/h)",
        "Bemerkung"
    ]
    ws.append(headers)
    for cell in ws[current_row]:
        cell.font = Font(bold=True)
        cell.border = border
        cell.fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        cell.font = Font(bold=True, color='FFFFFF')
    current_row += 1

    # Data rows
    for _, row_data in df_budget.iterrows():
        projekt = row_data["Projekt"]
        obermeilenstein = row_data["Obermeilenstein"]
        billing_type = row_data["Abrechnungsart"]
        gesamtbudget = row_data["Gesamtbudget"]
        abgerechnet = row_data["Abgerechnet"]
        rate_sv = row_data["Stundensatz_SV"]
        rate_cad = row_data["Stundensatz_CAD"]
        rate_adm = row_data["Stundensatz_ADM"]

        # Status berechnen
        has_billing_type = billing_type != "Unbekannt"
        has_rates = (
            (rate_sv is not None and rate_sv > 0) or
            (rate_cad is not None and rate_cad > 0) or
            (rate_adm is not None and rate_adm > 0)
        )

        if has_billing_type and has_rates:
            status = "✓ Vollständig"
        else:
            status = "⚠ Manual erforderlich"

        # Verfügbar als Formel
        verfuegbar_formula = f"=E{current_row}-F{current_row}"

        # Bemerkung generieren
        bemerkung = []
        if not has_billing_type:
            bemerkung.append("Abrechnungsart wählen")
        if not has_rates:
            bemerkung.append("Stundensätze eingeben")
        bemerkung_text = "; ".join(bemerkung)

        ws.append([
            projekt,
            obermeilenstein,
            billing_type,
            status,
            gesamtbudget if gesamtbudget > 0 else "",
            abgerechnet if abgerechnet > 0 else 0,
            verfuegbar_formula,
            rate_sv if rate_sv is not None else "",
            rate_cad if rate_cad is not None else "",
            rate_adm if rate_adm is not None else "",
            bemerkung_text
        ])

        # Styling für die Zeile
        for col_idx, cell in enumerate(ws[current_row], start=1):
            cell.border = border

            # Abrechnungsart Dropdown (Spalte C)
            if col_idx == 3:
                dv = DataValidation(type="list", formula1='"Pauschale,Nachweis,Unbekannt"', allow_blank=False)
                dv.add(cell)
                ws.add_data_validation(dv)

                # Rot markieren wenn "Unbekannt"
                if billing_type == "Unbekannt":
                    cell.fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
                    cell.font = Font(color='9C0006', bold=True)

            # Status Spalte (Spalte D)
            if col_idx == 4:
                if status.startswith("⚠"):
                    cell.fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
                    cell.font = Font(color='9C5700', bold=True)
                else:
                    cell.fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
                    cell.font = Font(color='006100', bold=True)

            # Budget-Spalten (E, F, G)
            if col_idx in [5, 6, 7]:
                cell.number_format = '#,##0.00'

            # Stundensatz-Spalten (H, I, J) - Gelb markieren wenn leer
            if col_idx in [8, 9, 10]:
                cell.number_format = '#,##0.00'
                if cell.value == "":
                    cell.fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')

        current_row += 1

    # Column widths
    ws.column_dimensions['A'].width = 50
    ws.column_dimensions['B'].width = 50
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 18
    ws.column_dimensions['G'].width = 18
    ws.column_dimensions['H'].width = 22
    ws.column_dimensions['I'].width = 22
    ws.column_dimensions['J'].width = 22
    ws.column_dimensions['K'].width = 40


def _create_cover_sheet(
    wb: Workbook,
    target_quarter: pd.Period,
    months: Iterable[pd.Period],
    employee_summary_data: Dict,
    border: Border,
) -> None:
    """Creates a cover sheet with summary totals across all employees."""

    # Create summary sheet
    ws = wb.create_sheet(title="Übersicht", index=0)

    # Title
    ws.append([f"Quartalsübersicht {target_quarter} - Zusammenfassung aller Mitarbeiter"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([])

    current_row = 3

    # Monthly summary table
    ws.append(["--- Monatliche Summen ---"])
    ws[f"A{current_row}"].font = Font(bold=True, size=12)
    current_row += 1

    # Header row
    ws.append(["Monat", "Gesamtstunden", "Bonusberechtigte Stunden", "Bonusberechtigte Stunden Sonderprojekt"])
    for cell in ws[current_row]:
        cell.font = Font(bold=True)
        cell.border = border
    current_row += 1

    # Get all months from first employee (all employees should have same months)
    if employee_summary_data:
        first_emp = list(employee_summary_data.keys())[0]
        month_labels = list(employee_summary_data[first_emp]['months'].keys())

        # For each month, create a summary row
        for month_label in month_labels:
            # Collect cell references for all employees for this month
            total_hours_refs = []
            bonus_hours_refs = []
            special_bonus_refs = []

            for emp, emp_data in employee_summary_data.items():
                if month_label in emp_data['months']:
                    month_data = emp_data['months'][month_label]
                    total_hours_refs.append(month_data['total_hours_cell'])
                    bonus_hours_refs.append(month_data['bonus_hours_cell'])
                    special_bonus_refs.append(month_data['special_bonus_hours_cell'])

            # Create formulas summing across all employees
            total_hours_formula = f"=SUM({','.join(total_hours_refs)})" if total_hours_refs else "0"
            bonus_hours_formula = f"=SUM({','.join(bonus_hours_refs)})" if bonus_hours_refs else "0"
            special_bonus_formula = f"=SUM({','.join(special_bonus_refs)})" if special_bonus_refs else "0"

            ws.append([month_label, total_hours_formula, bonus_hours_formula, special_bonus_formula])
            for cell in ws[current_row]:
                cell.border = border
                if cell.column > 1:  # Format numeric columns
                    cell.number_format = "0.00"
            current_row += 1

    ws.append([])
    current_row += 1

    # Quarterly summary
    ws.append(["--- Quartalssummen ---"])
    ws[f"A{current_row}"].font = Font(bold=True, size=12)
    current_row += 1

    # Collect quarterly total cell references from all employees
    quarter_total_refs = []
    quarter_bonus_refs = []
    quarter_special_refs = []

    for emp, emp_data in employee_summary_data.items():
        if 'quarter_total_hours_cell' in emp_data:
            quarter_total_refs.append(emp_data['quarter_total_hours_cell'])
        if 'quarter_bonus_hours_cell' in emp_data:
            quarter_bonus_refs.append(emp_data['quarter_bonus_hours_cell'])
        if 'quarter_special_bonus_hours_cell' in emp_data:
            quarter_special_refs.append(emp_data['quarter_special_bonus_hours_cell'])

    # Total hours
    ws.append(["Gesamt eingetragene Stunden:", f"=SUM({','.join(quarter_total_refs)})" if quarter_total_refs else "0"])
    ws[f"B{current_row}"].number_format = "0.00"
    for cell in ws[current_row]:
        cell.font = Font(bold=True)
        cell.border = border
    current_row += 1

    # Bonus hours
    ws.append(["Bonusberechtigte Stunden (Quartal):", f"=SUM({','.join(quarter_bonus_refs)})" if quarter_bonus_refs else "0"])
    ws[f"B{current_row}"].number_format = "0.00"
    for cell in ws[current_row]:
        cell.font = Font(bold=True)
        cell.border = border
    current_row += 1

    # Special bonus hours
    ws.append(["Bonusberechtigte Stunden Sonderprojekt (Quartal):", f"=SUM({','.join(quarter_special_refs)})" if quarter_special_refs else "0"])
    ws[f"B{current_row}"].number_format = "0.00"
    for cell in ws[current_row]:
        cell.font = Font(bold=True)
        cell.border = border
    current_row += 1

    ws.append([])
    current_row += 1

    # Employee list
    ws.append(["--- Mitarbeiter in diesem Quartal ---"])
    ws[f"A{current_row}"].font = Font(bold=True, size=12)
    current_row += 1

    for emp in sorted(employee_summary_data.keys()):
        ws.append([emp])
        ws[f"A{current_row}"].border = border
        current_row += 1

    # Set column widths
    ws.column_dimensions['A'].width = 50
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 25
    ws.column_dimensions['D'].width = 35


def build_quarterly_report(
    df_csv: pd.DataFrame,
    df_budget: pd.DataFrame,
    work_package_to_obermeilenstein: dict,
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

    # Create Projekt-Budget-Übersicht sheet first
    progress_cb(18, "Erstelle Projekt-Budget-Übersicht")
    _create_project_budget_sheet(wb, df_budget, border)

    # Create a lookup dict for budget data keyed by (Projekt, Obermeilenstein)
    # CRITICAL: Budgets are per Obermeilenstein, NOT summed across the project!
    # Each Obermeilenstein has its own separate budget.
    budget_lookup = {}

    # Store each Obermeilenstein's budget separately
    for _, row in df_budget.iterrows():
        projekt = str(row["Projekt"]).strip()
        obermeilenstein = str(row["Obermeilenstein"]).strip()

        if projekt and obermeilenstein and projekt != "-" and obermeilenstein != "-":
            budget_data = {
                "Sollhonorar": row["Gesamtbudget"],
                "Verrechnete_Honorare": row["Abgerechnet"],
                "Istkosten": row["Istkosten"],
            }

            # Store with (Projekt, Obermeilenstein) tuple as key
            budget_lookup[(projekt, obermeilenstein)] = budget_data

            # Also store with (Projekt_code, Obermeilenstein) for flexible matching
            parts = projekt.split(maxsplit=1)
            if parts:
                projekt_code = parts[0].strip()
                budget_lookup[(projekt_code, obermeilenstein)] = budget_data

    employees = sorted(df_quarter["staff_name"].unique())
    total_emps = max(len(employees), 1)

    # Build a map of which employees work on which project/milestone combinations
    # Format: {(proj_norm, ms_norm): [list of employee names]}
    project_milestone_employees = {}
    for _, row in df_quarter.iterrows():
        key = (row["proj_norm"], row["ms_norm"])
        emp_name = row["staff_name"]
        if key not in project_milestone_employees:
            project_milestone_employees[key] = set()
        project_milestone_employees[key].add(emp_name)
    # Convert sets to sorted lists
    for key in project_milestone_employees:
        project_milestone_employees[key] = sorted(list(project_milestone_employees[key]))

    # Dictionary to store cell references for summary sheet
    employee_summary_data = {}

    # Dictionary to track row assignments for "Von anderen" formula generation
    # Format: {employee: {(proj_norm, ms_norm, month): row_number}}
    row_assignments = {}

    # Dictionary to track month sections for summing "Von anderen" cells
    # Format: {employee: {month: (start_row, end_row, assigned_from_others_cell_row)}}
    month_sections = {}

    for idx_emp, emp in enumerate(employees, start=1):
        row_assignments[emp] = {}
        month_sections[emp] = {}
        ws = wb.create_sheet(title=emp[:31])
        sheet_name = emp[:31]
        monthly_bonus_total_cells: List[str] = []
        monthly_special_bonus_total_cells: List[str] = []
        transfer_entries: List[Tuple[str, str, str, str]] = []

        # Store monthly data for summary sheet
        if emp not in employee_summary_data:
            employee_summary_data[emp] = {
                'sheet_name': sheet_name,
                'months': {}
            }

        ws.append([f"{emp} - Quartalsreport {target_quarter}"])

        # Position dropdown for employee
        ws.append(["Position:", "SV"])  # Default: SV
        position_cell = ws.cell(row=2, column=2)
        position_dv = DataValidation(type="list", formula1='"SV,CAD,ADM,Pauschale,-"', allow_blank=False)
        position_dv.add(position_cell)
        ws.add_data_validation(position_dv)
        ws.cell(row=2, column=1).font = Font(bold=True)

        ws.append([])

        current_row = 4
        total_hours_all_months = 0.0
        total_bonus_hours_quarter = 0.0
        total_bonus_special_hours_quarter = 0.0

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
                        # Monthly budget (NOT cumulative - each month has its own budget)
                        month_data.loc[idx, "Soll"] = MONTHLY_BUDGETS[ms_name]
                        month_data.loc[idx, "Ist"] = month_data.loc[idx, "hours"]

            for idx in month_data.index:
                ms_name = month_data.loc[idx, "Meilenstein"]
                ms_type = month_data.loc[idx, "MeilensteinTyp"]
                proj_name = month_data.loc[idx, "Projekte"]
                proj_norm = month_data.loc[idx, "proj_norm"] if "proj_norm" in month_data.columns else ""
                if ms_type == "monthly" and ms_name in MONTHLY_BUDGETS and (is_bonus_project(proj_name) or is_bonus_project(proj_norm)):
                    # Monthly budget (NOT cumulative - each month has its own budget)
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

            # Cumulative XML hours up to current month (for quarterly milestones)
            df_to_date = df_quarter[(df_quarter["staff_name"] == emp) & (df_quarter["period"] <= month)]
            cum_hours_map = {
                (r["proj_norm"], r["ms_norm"]): r["hours"]
                for _, r in df_to_date.groupby(["proj_norm", "ms_norm"], as_index=False).agg({"hours": "sum"}).iterrows()
            }

            # XML hours for months AFTER the current month - FOR ALL EMPLOYEES (for backward calculation)
            # This ensures all employees see the same IST value for the same project/milestone
            df_after_month_all_employees = df_quarter[df_quarter["period"] > month]
            future_hours_all_employees_map = {
                (r["proj_norm"], r["ms_norm"]): r["hours"]
                for _, r in df_after_month_all_employees.groupby(["proj_norm", "ms_norm"], as_index=False).agg({"hours": "sum"}).iterrows()
            }

            month_data = month_data.sort_values(["Projekte", "Meilenstein"])

            month_name = MONTH_NAMES.get(int(month.month), month.strftime('%B'))
            month_str = f"{month_name} {month.year}"

            ws.append([f"--- {month_str} ---"])
            ws[f"A{current_row}"].font = Font(bold=True, size=12)
            current_row += 1

            ws.append(["Projekt", "Meilenstein", "Typ", "Soll (h)", "Ist (h)", f"{month_str} (h)", "%", "Bonus-Anpassung (h)", "Differenz (h)", "Zuordnen an", "Von anderen (h)", "Stundensatz (€/h)", "Umsatz (€)", "Budget Gesamt (€)", "Budget Ist (€)", "Budget Erwirtschaftet (€)"])
            for cell in ws[current_row]:
                cell.font = Font(bold=True)
                cell.border = border
            current_row += 1

            # Track start of month data section
            month_data_start_row = current_row

            bonus_hours_month = 0.0
            bonus_hours_month_special = 0.0
            adjustment_cells_regular: List[str] = []
            adjustment_cells_special: List[str] = []

            for proj, proj_block in month_data.groupby("Projekte", sort=False):
                proj_block = proj_block.reset_index(drop=True)
                block_start = current_row

                for i, (_, row_data) in enumerate(proj_block.iterrows()):
                    ms_type = row_data["MeilensteinTyp"]
                    hours_value = float(row_data.get("hours") or 0.0)
                    is_special_project = is_bonus_project(proj) or is_bonus_project(row_data.get("proj_norm", ""))
                    if is_special_project:
                        typ_short = "Q" if ms_type == "quarterly" else "M"
                    else:
                        typ_short = "G"
                    bonus_candidate = False
                    should_color = False
                    color_percentage = 0.0

                    if ms_type == "monthly":
                        soll_value = float(row_data.get("Soll") or 0.0)
                        csv_ist_total = float(row_data.get("Ist") or 0.0)
                        ms_name = row_data["Meilenstein"]

                        # For 0000-projects with MONTHLY_BUDGETS: IST = current month hours (no backward calc)
                        if is_special_project and ms_name in MONTHLY_BUDGETS:
                            ist_display = hours_value
                        else:
                            # For normal projects: backward calculation from CSV IST
                            # Using ALL EMPLOYEES' future hours to ensure consistent IST across all employees
                            # Last month: IST = CSV_IST
                            # Previous months: IST = CSV_IST - XML_hours_of_future_months_ALL_EMPLOYEES
                            key = (row_data["proj_norm"], row_data["ms_norm"])
                            xml_future_hours_all = float(future_hours_all_employees_map.get(key, 0.0))
                            ist_display = csv_ist_total - xml_future_hours_all

                        if soll_value > 0:
                            pct_value = (ist_display / soll_value) * 100.0 if soll_value else 0.0
                            should_color = True
                            color_percentage = pct_value
                            if pct_value <= 100.0:
                                bonus_candidate = True
                        else:
                            pct_value = 0.0
                            bonus_candidate = True
                        ws.append([
                            proj if i == 0 else "",
                            row_data["Meilenstein"],
                            typ_short,
                            round(soll_value, 2),
                            round(ist_display, 2),
                            round(hours_value, 2),
                            round(pct_value, 2),
                            None,  # Bonus-Anpassung (H)
                            None,  # Differenz (I) - will be filled with formula
                            None,  # Zuordnen an (J) - Dropdown will be added later
                            None,  # Von anderen (K) - will be filled with formula
                            None,  # Stundensatz (L) - Formula will be added later
                            None,  # Umsatz (M) - Formula will be added later
                            None,  # Budget Gesamt (N) - Formula will be added later
                            None,  # Budget Ist (O) - From CSV
                            None,  # Budget Erwirtschaftet (P) - Formula will be added later
                        ])
                    else:
                        q_soll = float(row_data.get("QuartalsSoll", 0.0) or 0.0)
                        cum_ist = float(cum_hours_map.get((row_data["proj_norm"], row_data["ms_norm"]), 0.0))
                        prozent = (cum_ist / q_soll * 100.0) if q_soll > 0 else 0.0
                        if prozent <= 100.0:
                            bonus_candidate = True
                        should_color = q_soll > 0
                        color_percentage = prozent
                        ws.append([
                            proj if i == 0 else "",
                            row_data["Meilenstein"],
                            typ_short,
                            round(q_soll, 2) if q_soll > 0 else "-",
                            round(cum_ist, 2) if cum_ist > 0 else 0.0,
                            round(hours_value, 2),
                            round(prozent, 2) if q_soll > 0 else "-",
                            None,  # Bonus-Anpassung (H)
                            None,  # Differenz (I) - will be filled with formula
                            None,  # Zuordnen an (J) - Dropdown will be added later
                            None,  # Von anderen (K) - will be filled with formula
                            None,  # Stundensatz (L) - Formula will be added later
                            None,  # Umsatz (M) - Formula will be added later
                            None,  # Budget Gesamt (N) - Formula will be added later
                            None,  # Budget Ist (O) - From CSV
                            None,  # Budget Erwirtschaftet (P) - Formula will be added later
                        ])

                    for cell in ws[current_row]:
                        cell.border = border

                    # Bonus-Anpassung cell (column H)
                    adj_cell = ws.cell(row=current_row, column=8)
                    if is_special_project:
                        adjustment_cells_special.append(adj_cell.coordinate)
                    else:
                        adjustment_cells_regular.append(adj_cell.coordinate)
                    adj_cell.number_format = "0.00"

                    # Differenz cell (column I) - Formula: F - H (only show if H != 0)
                    diff_cell = ws.cell(row=current_row, column=9)
                    diff_cell.value = f"=IF(H{current_row}=0,\"\",F{current_row}-H{current_row})"
                    diff_cell.number_format = "0.00"

                    # Zuordnen an cell (column J) - Dropdown with other employees on same project/milestone
                    assign_cell = ws.cell(row=current_row, column=10)
                    key = (row_data["proj_norm"], row_data["ms_norm"])
                    other_employees = [e for e in project_milestone_employees.get(key, []) if e != emp]
                    if other_employees:
                        # Create dropdown with other employees
                        employee_list = ",".join(other_employees)
                        dv = DataValidation(type="list", formula1=f'"{employee_list}"', allow_blank=True)
                        dv.add(assign_cell)
                        ws.add_data_validation(dv)

                    # Add conditional formatting to turn cell red when negative
                    red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
                    red_font = Font(color='9C0006')
                    ws.conditional_formatting.add(
                        diff_cell.coordinate,
                        CellIsRule(operator='lessThan', formula=['0'], stopIfTrue=True, fill=red_fill, font=red_font)
                    )

                    # Von anderen cell (column K) - Formula to sum hours assigned by other employees
                    # This will be filled in a second pass after all sheets are created
                    from_others_cell = ws.cell(row=current_row, column=11)
                    from_others_cell.number_format = "0.00"
                    # Mark this cell with row info for later formula injection
                    from_others_cell.value = 0  # Placeholder, will be replaced with formula

                    # Stundensatz cell (column L) - VLOOKUP formula to Projekt-Budget-Übersicht
                    # Uses Position from B2 (employee position dropdown)
                    rate_cell = ws.cell(row=current_row, column=12)
                    rate_cell.number_format = '#,##0.00'
                    # Formula depends on Position dropdown in B2: IF B2="-" then "", ELSE VLOOKUP
                    # Returns the corresponding rate (SV=H, CAD=I, ADM=J, Pauschale=H as fallback)
                    # Uses TRIM to remove extra spaces and make matching more robust
                    rate_formula = (
                        f'=IF($B$2="-","",IF($B$2="Pauschale",'
                        f'IFERROR(INDEX(\'Projekt-Budget-Übersicht\'!$H:$H,MATCH(TRIM(A{current_row}),\'Projekt-Budget-Übersicht\'!$A:$A,0)),0),'
                        f'IF($B$2="SV",'
                        f'IFERROR(INDEX(\'Projekt-Budget-Übersicht\'!$H:$H,MATCH(TRIM(A{current_row}),\'Projekt-Budget-Übersicht\'!$A:$A,0)),0),'
                        f'IF($B$2="CAD",'
                        f'IFERROR(INDEX(\'Projekt-Budget-Übersicht\'!$I:$I,MATCH(TRIM(A{current_row}),\'Projekt-Budget-Übersicht\'!$A:$A,0)),0),'
                        f'IF($B$2="ADM",'
                        f'IFERROR(INDEX(\'Projekt-Budget-Übersicht\'!$J:$J,MATCH(TRIM(A{current_row}),\'Projekt-Budget-Übersicht\'!$A:$A,0)),0),'
                        f'0)))))'
                    )
                    rate_cell.value = rate_formula

                    # Umsatz cell (column M) - Formula: Stundensatz * (Monat + Differenz + Von anderen)
                    revenue_cell = ws.cell(row=current_row, column=13)
                    revenue_cell.number_format = '#,##0.00'
                    revenue_formula = f'=IF($B$2="-","",L{current_row}*(F{current_row}+I{current_row}+K{current_row}))'
                    revenue_cell.value = revenue_formula

                    # Budget cells (columns N, O, P) - Get from budget_lookup
                    # CRITICAL: Budget is per Obermeilenstein, so we need to map work package -> Obermeilenstein
                    budget_total_cell = ws.cell(row=current_row, column=14)
                    budget_total_cell.number_format = '#,##0.00'
                    budget_ist_cell = ws.cell(row=current_row, column=15)
                    budget_ist_cell.number_format = '#,##0.00'
                    budget_earned_cell = ws.cell(row=current_row, column=16)
                    budget_earned_cell.number_format = '#,##0.00'

                    projekt_name = row_data["proj_norm"]
                    ms_name = row_data["ms_norm"]

                    # Find which Obermeilenstein this work package belongs to
                    obermeilenstein = work_package_to_obermeilenstein.get((projekt_name, ms_name))

                    # If not found with full project name, try with project code
                    if not obermeilenstein and " " in projekt_name:
                        projekt_code = projekt_name.split(maxsplit=1)[0].strip()
                        obermeilenstein = work_package_to_obermeilenstein.get((projekt_code, ms_name))

                    # Look up budget using (Projekt, Obermeilenstein) key
                    budget_key = (projekt_name, obermeilenstein) if obermeilenstein else None
                    budget_data = budget_lookup.get(budget_key) if budget_key else None

                    if budget_data:
                        # Budget Gesamt (Sollhonorar)
                        budget_total = budget_data.get("Sollhonorar", 0)
                        budget_total_cell.value = budget_total if budget_total > 0 else ""

                        # Budget Ist (Istkosten)
                        istkosten = budget_data.get("Istkosten", 0)
                        budget_ist_cell.value = istkosten if istkosten > 0 else ""

                        # Budget Erwirtschaftet (Verrechnete Honorare)
                        verrechnete = budget_data.get("Verrechnete_Honorare", 0)
                        budget_earned_cell.value = verrechnete if verrechnete > 0 else ""
                    else:
                        budget_total_cell.value = ""
                        budget_ist_cell.value = ""
                        budget_earned_cell.value = ""

                    # Track row for this project/milestone/month combination
                    track_key = (row_data["proj_norm"], row_data["ms_norm"], month)
                    row_assignments[emp][track_key] = current_row

                    if should_color:
                        pct_cell = ws.cell(row=current_row, column=7)
                        pct_cell.fill = PatternFill(
                            start_color=status_color_hex(color_percentage),
                            end_color=status_color_hex(color_percentage),
                            fill_type="solid",
                        )

                    if bonus_candidate:
                        if is_special_project:
                            bonus_hours_month_special += hours_value
                        else:
                            bonus_hours_month += hours_value

                    current_row += 1

                block_size = len(proj_block)
                if block_size > 1:
                    ws.merge_cells(start_row=block_start, start_column=1,
                                   end_row=block_start + block_size - 1, end_column=1)
                    ws.cell(row=block_start, column=1).alignment = Alignment(vertical="top")

            # Track end of month data section (before summary rows)
            month_data_end_row = current_row - 1

            sum_hours = month_data["hours"].sum()
            total_hours_all_months += sum_hours
            ws.append(["", "Summe", "", "", "", round(sum_hours, 2), "", "", "", "", ""])
            sum_row_idx = current_row
            for cell in ws[current_row]:
                cell.font = Font(bold=True)
                cell.border = border
            sum_total_cell = ws.cell(row=sum_row_idx, column=6)
            sum_total_cell.number_format = "0.00"
            sum_total_cell.value = round(sum_hours, 2)
            current_row += 1

            ws.append(["", "Bonusberechtigte Stunden", "", "", "", 0, "", "", "", "", "", "", "", "", "", ""])
            bonus_row_idx = current_row
            for cell in ws[current_row]:
                cell.font = Font(bold=True)
                cell.border = border
            bonus_base_cell = ws.cell(row=bonus_row_idx, column=7)
            bonus_base_cell.number_format = "0.00"
            bonus_base_cell.value = round(bonus_hours_month, 2)
            bonus_total_cell = ws.cell(row=bonus_row_idx, column=6)
            bonus_total_cell.number_format = "0.00"
            if adjustment_cells_regular:
                adj_formula = ",".join(adjustment_cells_regular)
                adj_sum_formula = f"SUM({adj_formula})"
                bonus_adj_cell = ws.cell(row=bonus_row_idx, column=8)
                bonus_adj_cell.value = f"={adj_sum_formula}"
                bonus_adj_cell.number_format = "0.00"
                bonus_total_cell.value = f"={bonus_base_cell.coordinate}+{bonus_adj_cell.coordinate}"
            else:
                bonus_total_cell.value = round(bonus_hours_month, 2)
                bonus_adj_cell = ws.cell(row=bonus_row_idx, column=8)
                bonus_adj_cell.value = 0
                bonus_adj_cell.number_format = "0.00"
            monthly_bonus_total_cells.append(bonus_total_cell.coordinate)
            total_bonus_hours_quarter += bonus_hours_month
            current_row += 1

            ws.append(["", "Bonusberechtigte Stunden Sonderprojekt", "", "", "", 0, "", "", "", "", "", "", "", "", "", ""])
            special_row_idx = current_row
            for cell in ws[current_row]:
                cell.font = Font(bold=True)
                cell.border = border
            special_base_cell = ws.cell(row=special_row_idx, column=7)
            special_base_cell.number_format = "0.00"
            special_base_cell.value = round(bonus_hours_month_special, 2)
            special_total_cell = ws.cell(row=special_row_idx, column=6)
            special_total_cell.number_format = "0.00"
            if adjustment_cells_special:
                adj_formula_special = ",".join(adjustment_cells_special)
                adj_sum_formula_special = f"SUM({adj_formula_special})"
                special_adj_cell = ws.cell(row=special_row_idx, column=8)
                special_adj_cell.value = f"={adj_sum_formula_special}"
                special_adj_cell.number_format = "0.00"
                special_total_cell.value = f"={special_base_cell.coordinate}+{special_adj_cell.coordinate}"
            else:
                special_total_cell.value = round(bonus_hours_month_special, 2)
                special_adj_cell = ws.cell(row=special_row_idx, column=8)
                special_adj_cell.value = 0
                special_adj_cell.number_format = "0.00"
            monthly_special_bonus_total_cells.append(special_total_cell.coordinate)
            total_bonus_special_hours_quarter += bonus_hours_month_special
            current_row += 1

            # Zugeordnete Stunden von anderen MA - will be calculated with formula
            ws.append(["", "Zugeordnete Stunden von anderen MA", "", "", "", 0, "", "", "", "", "", "", "", "", "", ""])
            assigned_from_others_row_idx = current_row
            for cell in ws[current_row]:
                cell.font = Font(bold=True)
                cell.border = border
            assigned_from_others_cell = ws.cell(row=assigned_from_others_row_idx, column=6)
            assigned_from_others_cell.number_format = "0.00"
            # Formula will sum all "Von anderen (K)" cells in this month's section
            # This will be filled after all sheets are created
            assigned_from_others_cell.value = 0  # Placeholder

            # Track month section info
            month_sections[emp][month] = (month_data_start_row, month_data_end_row, assigned_from_others_row_idx)

            current_row += 1

            # Gesamt Bonus Stunden = Bonusberechtigte + Sonderprojekt + Zugeordnete
            ws.append(["", "Gesamt Bonus Stunden", "", "", "", 0, "", "", "", "", ""])
            total_bonus_row_idx = current_row
            for cell in ws[current_row]:
                cell.font = Font(bold=True)
                cell.border = border
                cell.fill = PatternFill(start_color='D9EAD3', end_color='D9EAD3', fill_type='solid')
            total_bonus_cell = ws.cell(row=total_bonus_row_idx, column=6)
            total_bonus_cell.number_format = "0.00"
            total_bonus_cell.value = f"={bonus_total_cell.coordinate}+{special_total_cell.coordinate}+{assigned_from_others_cell.coordinate}"
            current_row += 1

            transfer_entries.append((month_str, sum_total_cell.coordinate, bonus_total_cell.coordinate, special_total_cell.coordinate, assigned_from_others_cell.coordinate, total_bonus_cell.coordinate))

            # Store cell references for summary sheet
            employee_summary_data[emp]['months'][month_str] = {
                'total_hours_cell': f"'{sheet_name}'!{sum_total_cell.coordinate}",
                'bonus_hours_cell': f"'{sheet_name}'!{bonus_total_cell.coordinate}",
                'special_bonus_hours_cell': f"'{sheet_name}'!{special_total_cell.coordinate}"
            }

            ws.append([])
            current_row += 1

        if transfer_entries:
            ws.append(["--- Übertragshilfe ---"])
            ws[f"A{current_row}"].font = Font(bold=True, size=12)
            current_row += 1

            ws.append(["Monat", "Mitarbeiter", "Prod. Stunden", "Bonusberechtigte Stunden", "Bonusberechtigte Stunden Sonderprojekt", "Zugeordnet von anderen", "Gesamt Bonus"])
            for cell in ws[current_row]:
                cell.font = Font(bold=True)
                cell.border = border
            current_row += 1

            for month_label, total_cell, bonus_cell, special_cell, assigned_cell, total_bonus_cell in transfer_entries:
                ws.append([month_label, emp, f"={total_cell}", f"={bonus_cell}", f"={special_cell}", f"={assigned_cell}", f"={total_bonus_cell}"])
                for cell in ws[current_row]:
                    cell.border = border
                current_row += 1

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

        ws.append(["Bonusberechtigte Stunden (Quartal):", 0])
        quarter_bonus_row = current_row
        quarter_bonus_cell = ws.cell(row=quarter_bonus_row, column=2)
        if monthly_bonus_total_cells:
            quarter_bonus_cell.value = f"=SUM({','.join(monthly_bonus_total_cells)})"
        else:
            quarter_bonus_cell.value = round(total_bonus_hours_quarter, 2)
        quarter_bonus_cell.number_format = "0.00"
        for cell in ws[current_row]:
            cell.font = Font(bold=True)
        current_row += 1

        ws.append(["Bonusberechtigte Stunden Sonderprojekt (Quartal):", 0])
        quarter_special_row = current_row
        quarter_special_cell = ws.cell(row=quarter_special_row, column=2)
        if monthly_special_bonus_total_cells:
            quarter_special_cell.value = f"=SUM({','.join(monthly_special_bonus_total_cells)})"
        else:
            quarter_special_cell.value = round(total_bonus_special_hours_quarter, 2)
        quarter_special_cell.number_format = "0.00"
        for cell in ws[current_row]:
            cell.font = Font(bold=True)
        current_row += 1

        # Store quarterly summary cell references
        employee_summary_data[emp]['quarter_total_hours_cell'] = f"'{sheet_name}'!B{quarter_bonus_row - 2}"
        employee_summary_data[emp]['quarter_bonus_hours_cell'] = f"'{sheet_name}'!{quarter_bonus_cell.coordinate}"
        employee_summary_data[emp]['quarter_special_bonus_hours_cell'] = f"'{sheet_name}'!{quarter_special_cell.coordinate}"

        ws.column_dimensions['A'].width = 40
        ws.column_dimensions['B'].width = 50
        ws.column_dimensions['C'].width = 8
        ws.column_dimensions['D'].width = 12
        ws.column_dimensions['E'].width = 12
        ws.column_dimensions['F'].width = 12
        ws.column_dimensions['G'].width = 8
        ws.column_dimensions['H'].width = 16
        ws.column_dimensions['I'].width = 12  # Differenz
        ws.column_dimensions['J'].width = 25  # Zuordnen an (Dropdown)
        ws.column_dimensions['K'].width = 15  # Von anderen
        ws.column_dimensions['L'].width = 18  # Stundensatz (€/h)
        ws.column_dimensions['M'].width = 15  # Umsatz (€)
        ws.column_dimensions['N'].width = 18  # Budget Gesamt (€)
        ws.column_dimensions['O'].width = 15  # Budget Ist (€)
        ws.column_dimensions['P'].width = 22  # Budget Erwirtschaftet (€)

        progress = int((idx_emp / total_emps) * 80) + 20
        progress_cb(min(progress, 95), f"Verarbeite Mitarbeiter {emp}")

    # Second pass: Fill "Von anderen" formulas
    progress_cb(90, "Erstelle Zuordnungs-Formeln")
    for emp in employees:
        ws = wb[emp[:31]]
        for track_key, row_num in row_assignments[emp].items():
            proj_norm, ms_norm, month = track_key
            # Build formula to sum hours from other employees who assigned to this employee
            formula_parts = []
            for other_emp in employees:
                if other_emp == emp:
                    continue
                other_key = (proj_norm, ms_norm, month)
                if other_key in row_assignments[other_emp]:
                    other_row = row_assignments[other_emp][other_key]
                    other_sheet = other_emp[:31]
                    # Add SUMIF formula part: IF assign cell (J) = current emp, then add diff cell (I)
                    formula_parts.append(f"IF('{other_sheet}'!J{other_row}=\"{emp}\",'{other_sheet}'!I{other_row},0)")

            # Set the formula in "Von anderen" cell (column K)
            from_others_cell = ws.cell(row=row_num, column=11)
            if formula_parts:
                formula = "=" + "+".join(formula_parts)
                from_others_cell.value = formula
            else:
                from_others_cell.value = 0

        # Fill "Zugeordnete Stunden von anderen MA" sum formulas
        for month, (start_row, end_row, assigned_row) in month_sections[emp].items():
            assigned_cell = ws.cell(row=assigned_row, column=6)
            # Sum all "Von anderen (K)" cells in this month section
            if start_row <= end_row:
                assigned_cell.value = f"=SUM(K{start_row}:K{end_row})"
            else:
                assigned_cell.value = 0

    # Create summary cover sheet
    progress_cb(96, "Erstelle Deckblatt")
    _create_cover_sheet(wb, target_quarter, months, employee_summary_data, border)

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

    progress_cb(8, "Lade Budget-Daten")
    df_budget, work_package_to_obermeilenstein = load_csv_budget_data(csv_path)

    progress_cb(10, "Lade XML-Daten")
    df_xml = load_xml_times(xml_path)

    selection = determine_quarter(df_xml, requested=requested_quarter)
    progress_cb(15, f"Wähle Quartal {selection.period}")

    year = selection.period.year
    quarter_num = selection.period.quarter
    out_path = output_dir / f"Q{quarter_num}-{year}.xlsx"

    result = build_quarterly_report(
        df_csv=df_csv,
        df_budget=df_budget,
        work_package_to_obermeilenstein=work_package_to_obermeilenstein,
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



