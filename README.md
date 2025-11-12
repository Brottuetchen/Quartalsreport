# Quartalsreport Generator

Webdienst auf Basis von FastAPI zur Erstellung quartalsweiser Bonusberichte aus einer Soll/Ist‑CSV und einer XML‑Zeiterfassung. Das Tool erzeugt je Quartal und Mitarbeiter eine Excel‑Arbeitsmappe mit farbcodierten Tabellen, Bonus-Anpassungen, Sonderprojekt-Logik und einer Übertragshilfe für die Konzernvorlage.

---

## Inhalt

- [Funktionen](#funktionen)
- [Benutzeranleitung](BENUTZERANLEITUNG.md) 📖
- [Projektstruktur](#projektstruktur)
- [Voraussetzungen](#voraussetzungen)
- [Docker-Nutzung](#docker-nutzung)
- [Lokale Entwicklung ohne Docker](#lokale-entwicklung-ohne-docker)
- [REST-API](#rest-api)
- [Aufbau der generierten Excel-Dateien](#aufbau-der-generierten-excel-dateien)
- [Portabler Windows-Runner](#portabler-windows-runner)
- [Troubleshooting](#troubleshooting)
- [Beitragen](#beitragen)
- [Lizenz](#lizenz)

---

## Funktionen

- CSV (Soll/Ist) und XML (Zeiterfassung) per Weboberfläche oder REST hochladen.
- Automatische Quartalsauswahl oder explizite Vorgabe (z. B. `Q3-2025`).
- Erstellung einer `.xlsx`-Mappe mit Monatsübersichten, Bonus-Anpassungsfeldern und separater Sonderprojekt-Summe.
- **Deckblatt mit Gesamtübersicht**: Automatisch generiertes Übersichtsblatt mit dynamischen Summen aller Mitarbeiter.
- Generierte Werte per Übertragshilfe einfach in die Firmenvorlage kopieren.
- Bereitstellung per Docker-Container, optional mit HTTP Basic Auth.
- Windows-Skripte für den portablen Offline-Einsatz.

---

## Projektstruktur

```
Quartalsreport/
├─ data/                  # Jobdaten (wird automatisch erzeugt)
├─ webapp/
│  ├─ report_generator.py # Kernlogik zur Excel-Ausgabe
│  ├─ server.py           # FastAPI-Server + Job-Queue
│  ├─ templates/          # Jinja2-Templates für das Web-UI
│  └─ static/             # CSS/JS-Assets
├─ Dockerfile             # Containerdefinition
├─ requirements.txt       # Python-Abhängigkeiten
├─ run_portable.ps1       # Portabler Windows-Runner (PowerShell)
└─ run_portable.cmd       # CMD-Wrapper für das PowerShell-Skript
```

---

## Voraussetzungen

- CSV-Export mit Spalten `Projekte`, `Arbeitspaket`, `Iststunden`, `Sollstunden Budget` (meist tab-getrennt, UTF-16 oder UTF-8).
- XML-Export der Zeiterfassung mit Mitarbeiter-, Projekt- und Meilensteininformationen.
- Optional: Python 3.11+ für lokale Entwicklung.
- Optional: Docker Engine 20.10+ für Containerbetrieb.

---

## Docker-Nutzung

### 1. Image bauen

```powershell
docker build -t quartalsreport .
```

### 2. Container starten

```powershell
docker run --rm ^
  -p 9999:9999 ^
  -v "$(pwd)/data:/app/data" ^
  quartalsreport
```

- Weboberfläche: <http://localhost:9999>
- Datenverzeichnis: `data/jobs/<job-id>/`

### Optional: HTTP Basic Auth

```powershell
docker run --rm ^
  -e BASIC_AUTH_USERNAME=benutzer ^
  -e BASIC_AUTH_PASSWORD=geheim ^
  -p 9999:9999 ^
  -v "$(pwd)/data:/app/data" ^
  quartalsreport
```

---

## Lokale Entwicklung ohne Docker

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
uvicorn webapp.server:app --host 0.0.0.0 --port 9999 --reload
```

Danach <http://localhost:9999> im Browser öffnen.

---

## REST-API

| Methode | Endpoint                           | Beschreibung                                                |
|---------|-----------------------------------|-------------------------------------------------------------|
| POST    | `/api/jobs`                       | Job mit `csv_file`, `xml_file`, optional `quarter` anlegen |
| GET     | `/api/jobs/{job_id}`              | Status und Fortschritt abrufen                             |
| GET     | `/api/jobs/{job_id}/download`     | Fertige Excel herunterladen (Status `finished`)            |
| DELETE  | `/api/jobs/{job_id}`              | Job löschen (falls nicht in Bearbeitung)                   |
| GET     | `/healthz`                        | Gesundheitscheck                                           |

Beispiel (PowerShell, `curl.exe`):

```powershell
curl.exe -X POST http://localhost:9999/api/jobs ^
  -F "csv_file=@C:\Pfad\report.csv" ^
  -F "xml_file=@C:\Pfad\zeiten.xml" ^
  -F "quarter=Q3-2025"
```

Akzeptierte Quartals-Formate: `Q3-2025`, `2025Q3`, `Q3/2025`, `2025-Q3`.

---

## Aufbau der generierten Excel-Dateien

Die generierte Excel-Datei enthält:

1. **Übersichtsblatt (Deckblatt)**: Zeigt monatliche und quartalsweise Summen über alle Mitarbeiter hinweg. Die Werte werden dynamisch über Formeln aus den Mitarbeiterblättern berechnet und aktualisieren sich automatisch bei Änderungen.

Für jeden Mitarbeiter des gewählten Quartals:

2. **Monatsbereiche** mit Soll/Ist, gebuchten Stunden, Farbkennzeichnung und der Spalte `Bonus-Anpassung (h)` für manuelle Korrekturen.
3. **Monatssummen** (Gesamtstunden, Bonusstunden, Bonusstunden Sonderprojekt) mit automatischer Aktualisierung bei Anpassungen.
4. **Quartalsübersicht** für Meilensteine mit Quartalssoll.
5. **Übertragshilfe**: Tabelle `Monat`, `Mitarbeiter`, `Prod. Stunden`, `Bonusberechtigte Stunden`, `Bonusberechtigte Stunden Sonderprojekt`.

Dateien liegen nach Fertigstellung unter `data/jobs/<job-id>/Q{Quartal}-{Jahr}.xlsx`.

---

## Portabler Windows-Runner

Das Skript `run_portable.ps1` lädt eine portable Python-Version, richtet Abhängigkeiten ein und startet den Legacy-Generator `Monatsbericht_Bonus_Quartal.py`.

```powershell
.\run_portable.ps1 -CsvPath "C:\Pfad\report.csv" `
                   -XmlPath "C:\Pfad\zeiten.xml" `
                   -OutputDir "C:\Ausgabe"
```

Hinweis: Der portable Runner arbeitet interaktiv. Für Mehrbenutzerbetrieb wird die Webanwendung empfohlen.

---

## Troubleshooting

- **HTTP 422 auf `/`**: Direktaufruf ohne Formulardaten. Weboberfläche nutzen oder POST `/api/jobs`.
- **Job bleibt auf „queued“**: Logs prüfen. Sicherstellen, dass genau ein Worker läuft und das Quartal in der XML vorhanden ist.
- **Excel enthält kaum Daten**: CSV- und XML-Projekte müssen nach Normalisierung (`proj_norm`, `ms_norm`) übereinstimmen.
- **Schreibrechte**: Docker-Volume oder lokales Dateisystem auf Schreibrechte prüfen (`data/jobs`).
- **Basic Auth unerwartet aktiv**: Umgebungsvariablen `BASIC_AUTH_USERNAME/PASSWORD` entfernen oder korrekte Daten nutzen.

---

## Beitragen

1. Virtuelle Umgebung erstellen, Abhängigkeiten installieren: `pip install -r requirements.txt`.
2. Tests bzw. `python -m compileall webapp/report_generator.py` ausführen oder eine Beispiel-Excel erzeugen.
3. Sicherstellen, dass der Docker-Build erfolgreich ist.
4. Pull Request mit kurzer Änderungsbeschreibung und Testergebnissen eröffnen.

---

## Lizenz

Internes Projekt. © 2025.

