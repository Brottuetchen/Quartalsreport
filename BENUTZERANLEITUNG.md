# Benutzeranleitung - Quartalsreport Generator

## Überblick

Das Tool ist als Webdienst verfügbar. Öffnen Sie Ihren Browser und navigieren Sie zur URL des Quartalsreport Generators (z.B. http://localhost:9999 oder die von Ihrem Administrator bereitgestellte Adresse).

---

## Schritt-für-Schritt-Anleitung

### 1. Dateien hochladen

**Upload-Formular:**

- Klicken Sie auf **CSV (Soll/Ist)** und wählen Sie Ihre Soll-Ist-CSV-Datei aus
- Klicken Sie auf **XML (Zeiteinträge)** und wählen Sie Ihre Zeiterfassungs-XML-Datei aus
- **(Optional)** Geben Sie das gewünschte Quartal an, z.B. `2025Q3` oder `Q3-2025`
  - Wenn leer gelassen, wählt das Tool automatisch das aktuellste Quartal aus der XML (neueste Buchungsperiode)

### 2. Report erzeugen

- Klicken Sie auf den Button **Report erzeugen**
- Der Upload-Fortschritt wird angezeigt
- Die Verarbeitung beginnt automatisch

### 3. Verarbeitung verfolgen

Während der Verarbeitung sehen Sie:

- **Fortschrittsbalken** – Zeigt den aktuellen Bearbeitungsstand
- **Statusmeldung** – Informiert über den aktuellen Verarbeitungsschritt
- **Warteschlange** – Position in der Warteschlange (falls mehrere Jobs aktiv)

### 4. Ergebnis herunterladen

Nach erfolgreicher Verarbeitung:

- Der Button **Ergebnis herunterladen** erscheint
- Klicken Sie darauf, um die Excel-Datei herunterzuladen
- Die Datei heißt z.B. `Q3-2025.xlsx`

### 5. PDFs exportieren (optional)

Nach dem Download der Excel-Datei:

- Klicken Sie auf **Als PDFs exportieren**
- Das Tool erstellt für jedes Arbeitsblatt ein separates PDF
- Die PDFs werden als Liste angezeigt und können einzeln heruntergeladen werden
- Benennung: `{Arbeitsblattname}_{Dateiname}.pdf`

### 6. Neuen Report starten

- Klicken Sie auf **Neuen Report starten**, um weitere Reports zu erzeugen

---

## Was enthält die generierte Excel-Datei?

### Übersichtsblatt (Deckblatt)

Das erste Arbeitsblatt "Übersicht" zeigt:

- **Monatliche Summen** über alle Mitarbeiter
  - Gesamtstunden
  - Bonusberechtigte Stunden
  - Bonusberechtigte Stunden Sonderprojekt
- **Quartalssummen** über alle Mitarbeiter
- Liste aller Mitarbeiter im Quartal

**Wichtig:** Alle Werte sind dynamisch und aktualisieren sich automatisch bei Änderungen in den Mitarbeiterblättern!

### Mitarbeiterblätter

Für jeden Mitarbeiter gibt es ein eigenes Arbeitsblatt mit:

1. **Monatliche Übersichten**
   - Projekt und Meilenstein
   - Soll/Ist-Stunden
   - Gebuchte Stunden pro Monat
   - Prozentuale Auslastung (farbcodiert)
   - Spalte **Bonus-Anpassung (h)** für manuelle Korrekturen

2. **Monatssummen**
   - Gesamtstunden
   - Bonusberechtigte Stunden
   - Bonusberechtigte Stunden Sonderprojekt

3. **Quartalsübersicht**
   - Quartals-Meilensteine mit Quartalssoll
   - Kumulative Ist-Stunden

4. **Übertragshilfe**
   - Tabelle zum einfachen Kopieren in die Firmenvorlage
   - Enthält: Monat, Mitarbeiter, Prod. Stunden, Bonusstunden

---

## Farbcodierung

Die Prozentangaben in den Tabellen sind farbcodiert:

- 🟢 **Grün** (< 90%): Projekt liegt unter Budget
- 🟡 **Gelb** (90-100%): Projekt nahe am Budget
- 🔴 **Rot** (> 100%): Budget überschritten

---

## Bonus-Anpassungen

In der Spalte **Bonus-Anpassung (h)** können Sie manuelle Korrekturen vornehmen:

- Positive Werte erhöhen die Bonusstunden
- Negative Werte verringern die Bonusstunden
- Die Summen aktualisieren sich automatisch

---

## Häufige Fragen (FAQ)

### Welche Dateiformate werden unterstützt?

- **CSV:** Tab-getrennt, UTF-8 oder UTF-16
- **XML:** Zeiterfassungs-Export mit Mitarbeiter-, Projekt- und Meilensteininformationen

### Wie gebe ich ein bestimmtes Quartal an?

Folgende Formate werden akzeptiert:
- `Q3-2025`
- `2025Q3`
- `Q3/2025`
- `2025-Q3`

### Was passiert, wenn kein Quartal angegeben wird?

Das Tool wählt automatisch das neueste Quartal aus der XML-Datei.

### Kann ich mehrere Reports gleichzeitig erzeugen?

Ja, das System verfügt über eine Warteschlange. Mehrere Jobs werden nacheinander abgearbeitet.

### Wie lange dauert die Verarbeitung?

Die Verarbeitung dauert je nach Datenmenge zwischen einigen Sekunden und wenigen Minuten.

### Was benötige ich für den PDF-Export?

Im Docker-Container ist LibreOffice bereits enthalten. Bei lokaler Installation muss LibreOffice separat installiert werden.

### Werden meine Daten gespeichert?

Jobs und generierte Dateien werden automatisch nach 7 Tagen gelöscht.

---

## Fehlerbehebung

### Upload schlägt fehl

- Prüfen Sie, ob die Dateien das richtige Format haben
- Stellen Sie sicher, dass die CSV tab-getrennt ist
- Überprüfen Sie die XML-Struktur

### "Job bleibt auf 'queued'"

- Prüfen Sie, ob das angegebene Quartal in der XML vorhanden ist
- Überprüfen Sie die Server-Logs

### Excel enthält kaum Daten

- CSV- und XML-Projekte müssen übereinstimmen
- Prüfen Sie die Projektnamen und Meilensteine in beiden Dateien

### PDF-Export schlägt fehl

- Bei lokalem Betrieb: Installieren Sie LibreOffice
- Download: https://www.libreoffice.org/download/

---

## Technische Hinweise

### REST-API

Für die Automatisierung steht eine REST-API zur Verfügung:

```bash
# Job erstellen
curl -X POST http://localhost:9999/api/jobs \
  -F "csv_file=@report.csv" \
  -F "xml_file=@zeiten.xml" \
  -F "quarter=Q3-2025"

# Status abfragen
curl http://localhost:9999/api/jobs/{job_id}

# Excel herunterladen
curl -O http://localhost:9999/api/jobs/{job_id}/download

# PDFs exportieren
curl -X POST http://localhost:9999/api/jobs/{job_id}/export-pdf

# PDF herunterladen
curl -O http://localhost:9999/api/jobs/{job_id}/pdf/{filename}
```

Weitere Informationen finden Sie in der [README.md](README.md).

---

## Support

Bei Fragen oder Problemen wenden Sie sich bitte an Ihren IT-Administrator.

---

© 2025 - Internes Projekt
