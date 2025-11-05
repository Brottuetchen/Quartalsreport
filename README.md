# Quartalsreport Generator

FastAPI-based web service and automation toolkit for assembling quarterly bonus reports from a workload CSV (Soll/Ist) and an XML export of recorded hours. The service produces an Excel workbook per quarter and per employee that includes colour-coded status tables, bonus-hour adjustments, and helper tables for transferring totals into the corporate template.

---

## Features
- Upload CSV (Soll/Ist budgets) and XML (time tracking) files via the web UI or REST API.
- Automatically determines the most recent quarter or accepts an explicit quarter selection (e.g. `Q3-2025`).
- Builds an `.xlsx` workbook with per-employee sheets, bonus adjustments, special-project handling, and a transfer helper table.
- Docker image for easy deployment; optional HTTP Basic Auth via environment variables.
- Portable PowerShell script for running the legacy CLI on Windows without administrator rights.

---

## Repository Layout

```
Quartalsreport/
├─ data/                  # Local job storage (created automatically)
├─ webapp/
│  ├─ report_generator.py # Core report-building logic
│  ├─ server.py           # FastAPI server + job queue
│  ├─ templates/          # Jinja2 HTML templates
│  └─ static/             # CSS/JS assets for the web UI
├─ Dockerfile             # Container image definition
├─ requirements.txt       # Python package dependencies
├─ run_portable.ps1       # Windows helper to run the legacy CLI without install
└─ run_portable.cmd       # Convenience wrapper for PowerShell script
```

---

## Prerequisites

- CSV export with columns `Projekte`, `Arbeitspaket`, `Iststunden`, `Sollstunden Budget` (tab-delimited UTF-16/UTF-8).
- XML export with time-tracking data (one record per employee/project/milestone).
- Optional: Python 3.11+ for local development.
- Optional: Docker Engine 20.10+ for containerised runs.

---

## Docker Usage

### 1. Build the image

```powershell
docker build -t quartalsreport .
```

### 2. Run the container

```powershell
docker run --rm ^
  -p 9999:9999 ^
  -v "$(pwd)/data:/app/data" ^
  quartalsreport
```

This exposes the web UI at <http://localhost:9999>.  
The `/app/data` volume persists generated jobs (`data/jobs/<job-id>/`).

### Optional: HTTP Basic Auth

Set environment variables before running:

```powershell
docker run --rm ^
  -e BASIC_AUTH_USERNAME=youruser ^
  -e BASIC_AUTH_PASSWORD=yourpass ^
  -p 9999:9999 ^
  -v "$(pwd)/data:/app/data" ^
  quartalsreport
```

---

## Local Development (without Docker)

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
uvicorn webapp.server:app --host 0.0.0.0 --port 9999 --reload
```

Then open <http://localhost:9999>.

---

## REST API Overview

| Method | Endpoint                 | Description                                                 |
|--------|--------------------------|-------------------------------------------------------------|
| POST   | `/api/jobs`              | Create a job (`csv_file`, `xml_file`, optional `quarter`)   |
| GET    | `/api/jobs/{job_id}`     | Check status/progress                                       |
| GET    | `/api/jobs/{job_id}/download` | Download finished Excel (if `status == finished`)     |
| DELETE | `/api/jobs/{job_id}`     | Cancel/remove a queued or finished job                      |
| GET    | `/healthz`               | Simple health probe                                          |

Example job creation with explicit quarter:

```powershell
curl -X POST http://localhost:9999/api/jobs ^
  -F "csv_file=@C:\path\report.csv" ^
  -F "xml_file=@C:\path\zeiten.xml" ^
  -F "quarter=Q3-2025"
```

Accepted quarter formats: `Q3-2025`, `2025Q3`, `Q3/2025`, `2025-Q3`.

---

## Generated Workbook Structure

For each employee present in the selected quarter:

1. **Monthly tables** with Soll/Ist, recorded hours, colour-coded percentages, and a `Bonus-Anpassung (h)` column for manual adjustments.
2. **Monthly summary rows** (`Summe`, `Bonusberechtigte Stunden`, `Bonusberechtigte Stunden Sonderprojekt`) where totals update automatically when adjustments are entered.
3. **Quarterly overview** summarising cumulative progress on quarterly milestones.
4. **Transfer helper** table (`Monat`, `Mitarbeiter`, `Prod. Stunden`, `Bonusberechtigte Stunden`, `Bonusberechtigte Stunden Sonderprojekt`) for copy/paste into the corporate Excel template.

All generated workbooks are stored under `data/jobs/<job-id>/Q{quarter}-{year}.xlsx`.

---

## Legacy Portable Runner (Windows)

`run_portable.ps1` downloads a portable Python runtime, installs dependencies locally, and launches the legacy CLI (`Monatsbericht_Bonus_Quartal.py`). Usage:

```powershell
.\run_portable.ps1 -CsvPath "C:\path\report.csv" -XmlPath "C:\path\zeiten.xml" -OutputDir "C:\out"
```

> The portable script executes interactively and mirrors the logic within `webapp/report_generator.py`. Prefer the web service for multi-user scenarios.

---

## Troubleshooting

- **HTTP 422 on `/`**: Occurs when visiting the API without the required form data; use the web UI or POST `/api/jobs`.
- **Job stuck in "queued"**: Check Docker/container logs; ensure only one worker is running and the XML contains the requested quarter.
- **Empty Excel output**: Verify that CSV/ XML share normalised project and milestone names (`proj_norm`, `ms_norm`); the generator matches on these fields.
- **Permission errors on `data/jobs`**: Confirm the mounted volume (Docker) or local filesystem allows write access for the running user.
- **Basic Auth prompts unexpectedly**: Remove `BASIC_AUTH_*` env vars or supply the correct credentials.

---

## Contributing Guidelines

1. Create a virtual environment and install dependencies (`pip install -r requirements.txt`).
2. Run `python -m compileall webapp/report_generator.py` or generate a sample workbook to validate changes.
3. Ensure Docker builds cleanly (`docker build .`).
4. Open a pull request with a concise summary of changes and test evidence.

---

## License

Internal project. Copyright © 2025.

