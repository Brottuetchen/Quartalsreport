# Projekt-Struktur

## Übersicht der wichtigen Dateien

### Haupt-Anwendung (`webapp/`)

```
webapp/
├── api/                           # REST API-Endpoints
│   ├── __init__.py
│   └── reports.py                 # Flexible Report API (NEU)
│
├── models/                        # Datenmodelle (NEU)
│   ├── __init__.py
│   └── report_config.py          # ReportConfig, TimeBlock, ReportType, TimeGrouping
│
├── services/                      # Business Logic (NEU)
│   ├── __init__.py
│   └── flexible_report_generator.py  # FlexibleReportGenerator
│
├── static/                        # Frontend-Assets
│   ├── app.js                    # Original JavaScript (alt)
│   ├── app-flexible.js           # Erweitert mit flexiblen Reports (NEU, aktiv)
│   └── styles.css                # CSS (erweitert)
│
├── templates/                     # HTML-Templates
│   └── index.html                # Web UI (erweitert mit Tabs)
│
├── __init__.py
├── server.py                      # FastAPI Server (erweitert)
└── report_generator.py            # Core Report Logic (bestehend)
```

### Beispieldaten (`BSP/`)

```
BSP/
├── gesamt.csv                    # Budget-Daten (Sollhonorar, Istkosten, etc.)
├── pauschale.csv                 # Alternative CSV-Datei
├── Nachweis.csv                  # Alternative CSV-Datei
├── alt.csv                       # Alte Version
├── LUD.xml                       # XML-Zeiterfassung
├── Q3-2025.xlsm                  # Generierter Standard-Report
├── Q3-2025-eintrag.xlsx          # Excel-Eingabedatei
└── *.xlsm                        # Weitere generierte Reports
```

### Dokumentation

```
.
├── README.md                     # Haupt-Dokumentation
├── BENUTZERANLEITUNG.md         # Benutzerhandbuch
├── FLEXIBLE_REPORTS_README.md   # Flexible Reports Feature-Doku (NEU)
└── PROJECT_STRUCTURE.md         # Diese Datei (NEU)
```

### Konfiguration & Maintenance

```
.
├── .gitignore                    # Git-Ausschlüsse (erweitert)
├── cleanup.py                    # Cleanup-Script (NEU)
├── Dockerfile                    # Docker-Image
├── docker-compose.yml            # Docker-Compose
└── requirements.txt              # Python-Dependencies
```

### Utility-Scripts (`scripts/`)

```
scripts/
├── check_encoding.py            # Encoding-Prüfung
└── fix_encoding.py              # Encoding-Korrektur
```

## Was wurde NEU hinzugefügt?

### 1. Flexible Report-Architektur

**Dateien:**
- `webapp/models/report_config.py` - Datenmodelle für flexible Konfiguration
- `webapp/services/flexible_report_generator.py` - Generator mit Zeitgruppierung
- `webapp/api/reports.py` - REST API für flexible Reports

**Features:**
- Benutzerdefinierte Zeiträume (z.B. 15.08-15.09)
- Verschiedene Zeitgliederungen (Monat, Woche, Periode, Keine)
- Filter für Projekte und Mitarbeiter
- Konfigurierbare Report-Komponenten

### 2. Erweitertes Web UI

**Dateien:**
- `webapp/templates/index.html` - Tab-Interface (Standard + Flexibel)
- `webapp/static/app-flexible.js` - JavaScript für beide Report-Typen
- `webapp/static/styles.css` - Erweitert mit Tab-Styles

**Features:**
- Tab-Navigation zwischen Standard- und Flexible Reports
- Formular mit erweiterten Optionen
- Filter-Checkboxen mit dynamischen Eingabefelder n
- Datums-Picker für Von/Bis

### 3. Maintenance-Tools

**Dateien:**
- `cleanup.py` - Automatisches Aufräumen
- `.gitignore` - Erweitert für Test-/Debug-Dateien

**Features:**
- Entfernt Test-Dateien (`test_*.py`, `check_*.py`, etc.)
- Löscht Python-Cache (`__pycache__`)
- Bereinigt temporäre Excel-Dateien

## Datei-Verantwortlichkeiten

### Core Report Generation
- `webapp/report_generator.py` - Basis-Logik (unverändert, kompatibel)
- `webapp/services/flexible_report_generator.py` - Erweiterte Logik (NEU)

### API Layer
- `webapp/server.py` - FastAPI Server, Job-Queue (erweitert)
- `webapp/api/reports.py` - Flexible Report Endpoints (NEU)

### Data Models
- `webapp/models/report_config.py` - Konfiguration, TimeBlocks (NEU)

### Frontend
- `webapp/templates/index.html` - Web UI (erweitert)
- `webapp/static/app-flexible.js` - JavaScript (NEU, ersetzt app.js)
- `webapp/static/styles.css` - CSS (erweitert)

## Was wurde gelöscht/bereinigt?

**Test-Dateien (via cleanup.py):**
- `test_*.py` - Alle Test-Scripts
- `check_*.py` - Debug-Scripts
- `validate_*.py` - Validierungs-Scripts
- `analyze_*.py` - Analyse-Scripts

**Temporäre Dateien:**
- `BSP/*_fixed.xlsm` - Test-Reports
- `BSP/*_vorlauf.xlsm` - Alte Versionen
- `BSP/Test_*.xlsm` - Test-Outputs
- `excel_check_output.txt` - Debug-Output

**Python-Cache:**
- `__pycache__/` - In allen Verzeichnissen

## Wichtige Konstanten & Konfiguration

**In `webapp/report_generator.py`:**
```python
MONTHLY_BUDGETS = {
    "Einarbeitung neuer Mitarbeiter": 8.0,
    "Angebote-Ausschreibungen-Kalkulationen": 8.0,
    # ...
}

QUARTERLY_BUDGETS = {
    "Firmenveranstaltungen": 4.0,
    # ...
}

MONTH_NAMES = {
    1: "Januar", 2: "Februar", ...
}
```

**In `webapp/models/report_config.py`:**
```python
class ReportType(Enum):
    QUARTERLY = "quarterly"
    CUSTOM_PERIOD = "custom_period"
    # ...

class TimeGrouping(Enum):
    BY_MONTH = "monthly"
    BY_PERIOD = "period"
    BY_WEEK = "weekly"
    NONE = "none"
```

## Rückwärtskompatibilität

✅ **Alles bleibt kompatibel:**
- Standard-Quartalsreports funktionieren wie bisher
- Bestehende API-Endpoints unverändert
- Alte Excel-Formeln bleiben korrekt
- Bestehender Code läuft weiter

## Für neue Entwickler

### Wo anfangen?
1. **Dokumentation lesen**: `README.md`, `FLEXIBLE_REPORTS_README.md`
2. **Server starten**: `uvicorn webapp.server:app --reload`
3. **UI öffnen**: http://localhost:8000
4. **Code ansehen**: Start bei `webapp/server.py`

### Typischer Workflow:
1. Änderung in Models (`webapp/models/`)
2. Business Logic in Services (`webapp/services/`)
3. API-Endpoint in API (`webapp/api/`)
4. Frontend in Templates/Static (`webapp/templates/`, `webapp/static/`)

### Testing:
```python
from webapp.services import FlexibleReportGenerator
from webapp.models import ReportConfig, ReportType, TimeGrouping
# ... siehe FLEXIBLE_REPORTS_README.md
```

## Deployment

### Development:
```bash
uvicorn webapp.server:app --reload --port 8000
```

### Production:
```bash
uvicorn webapp.server:app --host 0.0.0.0 --port 8000 --workers 4
```

### Docker:
```bash
docker-compose up -d
```
