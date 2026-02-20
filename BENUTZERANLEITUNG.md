# Benutzeranleitung: Quartalsreport Generator

## Inhaltsverzeichnis

1. [Überblick](#überblick)
2. [Benötigte Dateien](#benötigte-dateien)
3. [Standard Quartalsreport](#standard-quartalsreport)
4. [Flexibler Report](#flexibler-report)
5. [Admin-Bereich](#admin-bereich)
6. [Berechnungslogik](#berechnungslogik)
7. [Ausgabedatei verstehen](#ausgabedatei-verstehen)
8. [Häufige Probleme](#häufige-probleme)

---

## Überblick

Der **Quartalsreport Generator** erstellt automatisiert Excel-Berichte für die quartalsweise Bonusabrechnung. Das Tool vergleicht Soll- und Ist-Stunden aus der Projektverwaltung mit den tatsächlich gebuchten Zeiten und berechnet daraus bonusberechtigte Stunden.

Das Interface hat drei Bereiche:

| Tab | Zweck |
|-----|-------|
| **Standard Quartalsreport** | Vollständiger Quartalsbericht mit Bonus-Berechnung für alle Mitarbeiter |
| **Flexibler Report** | Benutzerdefinierte Berichte für beliebige Zeiträume, Mitarbeiter oder Projekte |
| **Admin** | Zentrale Budget-CSV verwalten, System-Updates einspielen |

---

## Benötigte Dateien

### 1. Soll-Ist CSV-Datei

Export aus der Projektverwaltung mit Budget-Informationen über alle Projekte.

**Erforderliche Spalten:**
- `Projekte` – Projektname oder Projektnummer
- `Honorarbereich` – Markierung der Obermeilensteine (`X`)
- `Arbeitspaket` – Meilenstein-/Arbeitspaketname
- `Sollstunden Budget` – Geplante Stunden (deutsches Zahlenformat, z.B. `1.234,56`)

**Dateiformat:**
- Endung: `.csv`
- Trennzeichen: Tab
- Kodierung: UTF-16 (Standard-Export) oder UTF-8

**Hinweis:** Diese Datei kann einmalig zentral vom Administrator hinterlegt werden – dann muss sie beim Standard-Report nicht mehr bei jedem Upload angegeben werden. Beim Flexiblen Report muss sie immer hochgeladen werden.

---

### 2. Zeiterfassungs-XML-Datei

XML-Export der gebuchten Zeiteinträge aus dem Zeiterfassungssystem.

**Erforderliche Felder in der XML:**
- `staff_name` – Mitarbeitername
- `project` – Projektname/-nummer
- `work_package_name` – Meilenstein/Arbeitspaket
- `date` – Buchungsdatum
- `number` – Gebuchte Stunden

**Dateiformat:**
- Endung: `.xml`
- Muss Zeiteinträge des gewünschten Zeitraums enthalten

**Mehrere Mitarbeiter / Zeiträume:** Die XML-Datei kann Einträge mehrerer Mitarbeiter und mehrerer Monate enthalten. Das Tool filtert automatisch das relevante Quartal heraus.

---

## Standard Quartalsreport

Der Standard-Report erstellt den klassischen Quartalsbericht mit vollständiger Bonus-Berechnung für alle Mitarbeiter in der XML-Datei.

### Schritt-für-Schritt

#### 1. Tab „Standard Quartalsreport" auswählen

Der Tab ist standardmäßig aktiv. Falls nicht, einfach oben auf **„Standard Quartalsreport"** klicken.

#### 2. Dateien auswählen

| Feld | Pflicht | Beschreibung |
|------|---------|--------------|
| **CSV (Soll/Ist)** | Optional* | Budget-CSV-Datei. Wird nicht benötigt, wenn der Administrator bereits eine zentrale CSV hinterlegt hat. |
| **XML (Zeiteinträge)** | Ja | Zeiterfassungs-XML. Mehrfachauswahl möglich – alle gewählten XML-Dateien werden zusammengeführt. |
| **Quartal** | Optional | Format: `Q3-2025` oder `2025Q3`. Wenn leer, wählt das Tool automatisch das Quartal mit den meisten/neuesten Buchungen aus der XML. |

*Wenn keine zentrale CSV hinterlegt ist und keine hochgeladen wird, erscheint eine Fehlermeldung.

#### 3. Report erzeugen

Klicken Sie auf **„Report erzeugen"**. Der Fortschrittsbereich öffnet sich automatisch:

- **Statusmeldung** – zeigt den aktuellen Verarbeitungsschritt
- **Fortschrittsbalken** – prozentualer Fortschritt
- **Warteschlange** – Position, falls andere Reports gleichzeitig verarbeitet werden

#### 4. Ergebnis herunterladen

Nach erfolgreicher Verarbeitung erscheint der Button **„Ergebnis herunterladen"**. Die Datei wird als `<XMLDateiname>_Q3-2025.xlsm` gespeichert.

#### 5. Weiterer Report

Klicken Sie auf **„Neuen Report starten"** um weitere Reports zu erstellen.

---

### Quartal automatisch vs. manuell

| Situation | Empfehlung |
|-----------|-----------|
| XML enthält genau ein Quartal | Feld leer lassen – automatische Erkennung |
| XML enthält mehrere Quartale, aktuellstes gewünscht | Feld leer lassen |
| XML enthält mehrere Quartale, älteres gewünscht | Quartal explizit eingeben, z.B. `Q2-2025` |
| Jahresende: XML enthält Q4 und evtl. Q1 des Folgejahres | Quartal explizit eingeben |

---

## Flexibler Report

Der flexible Report ermöglicht Berichte für beliebige Zeiträume, einzelne Mitarbeiter oder Projekte – auch über Quartalsgrenzen hinaus.

> **Hinweis:** Beim Flexiblen Report müssen CSV und XML immer hochgeladen werden – eine zentrale CSV wird hier nicht verwendet.

---

### Übersicht der Einstellungen

```
┌─────────────────────────────────────────────────┐
│  Dateien      CSV + XML (beide Pflicht)          │
│  Report-Typ   Welche Art von Bericht?            │
│  Zeitraum     Von/Bis (Datum)                    │
│  Gliederung   Wie wird der Zeitraum aufgeteilt?  │
│  Filter       Nur bestimmte Projekte/Mitarbeiter │
│  Optionen     Was soll der Report enthalten?     │
└─────────────────────────────────────────────────┘
```

---

### Report-Typen

#### Quartalsreport (`quarterly`)

Entspricht dem Standard-Report, aber über den flexiblen Weg. Der Zeitraum wird automatisch auf das vollständige Quartal gerundet, das das „Von"-Datum enthält.

**Typische Verwendung:**
- Wenn zusätzliche Filter (Mitarbeiter, Projekte) benötigt werden
- Wenn die Gliederung angepasst werden soll

**Empfohlene Gliederung:** Nach Monaten getrennt

---

#### Benutzerdefinierter Zeitraum (`custom_period`)

Beliebiger Zeitraum von Datum A bis Datum B – auch monats- oder quartalsübergreifend.

**Typische Verwendung:**
- Halbjahresberichte (z.B. 01.01.–30.06.)
- Projektphasen, die nicht an Monatsgrenzen ausgerichtet sind
- Berichte für z.B. 15.08.–15.09.

**Beispiel-Konfigurationen:**

| Zeitraum | Von | Bis | Empfohlene Gliederung |
|----------|-----|-----|----------------------|
| Q4 2025 komplett | 01.10.2025 | 31.12.2025 | Nach Monaten |
| Halbjahr 2025 | 01.01.2025 | 30.06.2025 | Nach Monaten |
| Einzelne Phase | 15.08.2025 | 15.09.2025 | Ein Block |
| Gesamtjahr | 01.01.2025 | 31.12.2025 | Nach Monaten |

---

#### Monatsreport (`monthly`)

Bericht für genau einen Monat. Das Datumsintervall wird automatisch auf den vollständigen Monat des „Von"-Datums gerundet.

**Typische Verwendung:**
- Einzelner Monatsauszug für Rückfragen
- Zwischenstand während des laufenden Quartals

---

#### Jahresreport (`yearly`)

Bericht für das vollständige Kalenderjahr des „Von"-Datums.

**Typische Verwendung:**
- Jahresüberblick für Führungskräfte
- Gesamtauswertung Projektportfolio

**Empfohlene Gliederung:** Nach Monaten getrennt (sonst wird die Gesamtmenge kaum lesbar)

---

#### Projekt-Zusammenfassung (`project`)

Alle Buchungen auf ein oder mehrere Projekte, über den gewählten Zeitraum.

**Typische Verwendung:**
- Wie viele Stunden wurden insgesamt auf Projekt X gebucht?
- Mitarbeiterübersicht für ein bestimmtes Projekt

**Filter:** Immer mit dem Projektfilter kombinieren (sonst wird alles ausgegeben).

---

#### Mitarbeiter-Zusammenfassung (`employee`)

Alle Buchungen eines oder mehrerer Mitarbeiter über den gewählten Zeitraum.

**Typische Verwendung:**
- Individuelle Stundenauswertung
- Vergleich zwischen Mitarbeitern

**Filter:** Immer mit dem Mitarbeiterfilter kombinieren.

---

### Zeitliche Gliederung

Die Gliederung bestimmt, wie der gewählte Zeitraum in der Excel-Ausgabe aufgeteilt wird.

#### Nach Monaten getrennt (`monthly`)

Der Zeitraum wird in Monatsblöcke aufgeteilt. Jeder Monat erhält einen eigenen Abschnitt im Mitarbeiterblatt.

**Ergebnis:** Jan | Feb | Mär | ... (je ein Block)

**Empfehlung für:** Quartals-, Halbjahres- und Jahresberichte. Identisch zum Standard-Report.

**Hinweis:** Wenn der Zeitraum genau ein Quartal umfasst und der Typ „Quartalsreport" gewählt ist, wird automatisch der Standard-Quartalsreport-Modus verwendet (mit Quartalssummary und Übertragshilfe).

---

#### Ein zusammenhängender Block (`period`)

Der gesamte Zeitraum wird als ein einzelner Block behandelt – unabhängig davon, ob er mehrere Monate umfasst.

**Ergebnis:** 15.08–15.09 (ein Block)

**Empfehlung für:**
- Kurze, nicht monatsausgerichtete Zeiträume
- Wenn nur die Gesamtsumme interessiert, nicht die Monatsaufteilung

---

#### Nach Wochen (`weekly`)

Der Zeitraum wird in Wochenblöcke (Mo–So) aufgeteilt.

**Ergebnis:** KW 33 | KW 34 | KW 35 | ...

**Empfehlung für:**
- Sehr detaillierte Auswertungen
- Kurze Zeiträume (wenige Wochen)

---

#### Keine Gliederung (`none`)

Keine zeitliche Unterteilung – nur eine Gesamtsumme aller Buchungen im Zeitraum.

**Ergebnis:** Gesamtsumme über den gesamten Zeitraum

**Empfehlung für:**
- Schnelle Übersicht ohne Details
- Wenn nur die Endsumme benötigt wird

---

### Kombinationsmatrix – was passt zusammen?

| Ziel | Report-Typ | Zeitraum | Gliederung | Filter |
|------|-----------|----------|-----------|--------|
| Normaler Quartalsbericht | Quartalsreport | erstes Datum im Quartal | Nach Monaten | – |
| Halbjahr mit Monatsspalten | Benutzerdefiniert | 01.01.–30.06. | Nach Monaten | – |
| Nur eine Projektphase | Benutzerdefiniert | Phasendaten | Ein Block | Projekt |
| Ein Mitarbeiter, ein Quartal | Quartalsreport | – | Nach Monaten | Mitarbeiter |
| Gesamtstunden Projekt X | Projekt-Zusammenfassung | Jahresanfang–Jahresende | Keine | Projekt X |
| Vergleich zwei MA über Q3 | Mitarbeiter-Zusammenfassung | Q3-Zeitraum | Nach Monaten | MA A, MA B |
| Jahresübersicht | Jahresreport | irgendein Datum im Jahr | Nach Monaten | – |
| Wochendetails einer Phase | Benutzerdefiniert | 4 Wochen | Nach Wochen | Projekt oder MA |

---

### Filter

Filter schränken ein, welche Daten in den Report einfließen.

#### Nur bestimmte Projekte

Aktivieren Sie die Checkbox **„Nur bestimmte Projekte"** und tragen Sie die Projektkürzel/Codes kommagetrennt ein.

```
Beispiel: 0283.05, 0299.02, 7716.01
```

Die Eingabe muss dem Projektcode entsprechen, wie er in der XML-Zeiterfassung verwendet wird (nicht zwingend der vollständige Name).

**Tipp:** Wenn Sie sich unsicher sind, welcher Code verwendet wird, schauen Sie in die Zeiterfassungs-XML: Das `project`-Feld enthält den genauen Wert.

---

#### Nur bestimmte Mitarbeiter

Aktivieren Sie die Checkbox **„Nur bestimmte Mitarbeiter"** und tragen Sie die Namen kommagetrennt ein.

```
Beispiel: C. Trapp SV, A. Kokott, M. Mustermann
```

Der Name muss exakt dem `staff_name`-Feld in der XML entsprechen (Groß-/Kleinschreibung beachten).

---

### Report-Optionen (Checkboxen)

| Option | Standard | Beschreibung |
|--------|---------|--------------|
| **Bonus-Berechnung einschließen** | ✓ An | Berechnet bonusberechtigte Stunden (Soll/Ist-Vergleich). Deaktivieren für reine Stundenauswertungen ohne Budgetbezug. |
| **Budget-Übersicht einschließen** | ✓ An | Fügt ein Blatt „Projekt-Budget-Übersicht" mit allen Budgets aus der CSV hinzu. |
| **Zusammenfassungsblatt einschließen** | ✓ An | Erstellt ein Deckblatt mit Gesamtübersicht aller Mitarbeiter. |
| **Sonderprojekte ausschließen** | ✗ Aus | Filtert 0000-Projekte komplett aus dem Bericht heraus. Nützlich für reine Projektarbeit-Berichte. |

---

### Typische Workflows – Flexibler Report

#### Workflow A: Quartalsbericht für einen Mitarbeiter

1. Report-Typ: **Quartalsreport**
2. Zeitraum: **01.10.2025 – 31.12.2025**
3. Gliederung: **Nach Monaten getrennt**
4. Filter: **Nur bestimmte Mitarbeiter → Name eintragen**
5. Optionen: alle Standard

---

#### Workflow B: Projektauswertung für ein Halbjahr

1. Report-Typ: **Benutzerdefinierter Zeitraum**
2. Zeitraum: **01.01.2025 – 30.06.2025**
3. Gliederung: **Nach Monaten getrennt**
4. Filter: **Nur bestimmte Projekte → Projekt-Code eintragen**
5. Optionen: Bonus-Berechnung **aus**, Budget-Übersicht **an**

---

#### Workflow C: Schnell-Check über Quartalsgrenzen

Manchmal liegen Projektphasen über zwei Quartale (z.B. 15.11.2025–15.02.2026).

1. Report-Typ: **Benutzerdefinierter Zeitraum**
2. Zeitraum: **15.11.2025 – 15.02.2026**
3. Gliederung: **Nach Monaten getrennt** (empfohlen) oder **Ein Block**
4. Filter: nach Bedarf
5. Optionen: alle Standard

Das Tool verarbeitet in diesem Fall Daten quartalsübergreifend korrekt.

---

#### Workflow D: Jahresübersicht ohne Bonus

1. Report-Typ: **Jahresreport**
2. Zeitraum: **01.01.2025** (Enddatum wird automatisch auf 31.12.2025 gesetzt)
3. Gliederung: **Nach Monaten getrennt**
4. Optionen: Bonus-Berechnung **aus**, Zusammenfassungsblatt **an**

---

## Admin-Bereich

Der Admin-Bereich ist passwortgeschützt und nur für Administratoren zugänglich.

### Anmelden

1. Tab **„Admin"** klicken
2. Benutzername und Passwort eingeben
3. **„Anmelden"** klicken

Die Anmeldedaten werden für die aktuelle Browser-Sitzung gespeichert. Beim Schließen des Browsers oder manuellen Abmelden werden sie gelöscht.

> Die Zugangsdaten werden vom Administrator in der `.env`-Datei auf dem Server festgelegt.

---

### Zentrale Budget-CSV

Die zentrale Budget-CSV ermöglicht es, die Soll-Ist-Daten einmalig zu hinterlegen. Standard-Reports können dann ohne erneuten CSV-Upload gestartet werden.

#### Aktuelle CSV anzeigen

Nach dem Anmelden zeigt der Bereich **„Zentrale Budget-CSV"** automatisch:
- Dateiname der aktuell hinterlegten CSV
- Dateigröße
- Datum der letzten Aktualisierung

Wenn noch keine CSV hinterlegt ist, erscheint die Meldung *„Noch keine Budget-CSV hinterlegt."*

#### CSV aktualisieren

1. Klicken Sie auf **„Datei auswählen"** im Abschnitt „Neue CSV hochladen"
2. Wählen Sie die neue CSV-Datei aus
3. Klicken Sie auf **„CSV aktualisieren"**

Nach dem Upload wird die Dateiinfo automatisch aktualisiert.

**Wann aktualisieren?** Immer wenn in der Projektverwaltung neue Projekte oder Meilensteine angelegt wurden oder sich Budgets geändert haben.

---

### System-Update (OTA)

Ermöglicht das Einspielen von Software-Updates, ohne Docker neu starten zu müssen.

#### Update-Datei erstellen

Im Repository-Verzeichnis auf dem Entwicklungsrechner:

```bash
git archive HEAD --format=zip > update.zip
```

#### Update einspielen

1. Klicken Sie auf **„Datei auswählen"** im Abschnitt „System-Update"
2. Wählen Sie die `update.zip` aus
3. Klicken Sie auf **„Update einspielen"**
4. Nach erfolgreichem Upload lädt die Seite automatisch neu (~4 Sekunden)

**Hinweis:** Das System erkennt Änderungen an den Webdateien automatisch durch den eingebauten Reload-Mechanismus. Kein Neustart des Docker-Containers notwendig.

---

## Berechnungslogik

### Grundprinzip

**Bonusberechtigt sind alle Stunden, bei denen das Budget nicht zu 100% ausgeschöpft wurde.**

Sobald das Soll eines Eintrags überschritten wird, ist der gesamte Eintrag nicht bonusberechtigt – eine anteilige Anerkennung findet nicht statt. Manuelle Korrekturen sind über die Spalte „Bonus-Anpassung" möglich.

---

### Meilenstein-Typen

Die Spalte **Typ** in der Excel-Ausgabe unterscheidet drei Fälle:

#### Typ G – Reguläre Projekte

- Alle Projekte, deren Nummer **nicht** mit `0000` beginnt
- Soll/Ist-Werte stammen vollständig aus der CSV-Datei
- Bonusberechtigt: solange `Ist / Soll < 100%`

#### Typ M – Sonderprojekt (Monatsbudget)

- Nur für **0000-Sonderprojekte** mit festem **Monatsbudget** (z.B. 8h/Monat)
- Vergleich: Ist-Stunden des Monats vs. Monatsbudget
- Bonusberechtigt: wenn `Ist-Monat / Monatsbudget < 100%`

**0000-Meilensteine mit Monatsbudget:**

| Meilenstein | Budget |
|------------|--------|
| Einarbeitung neuer Mitarbeiter (max. 8h/Monat pro MA) | 8 h/Monat |
| Angebote-Ausschreibungen-Kalkulationen (max. 8h/Monat pro MA) | 8 h/Monat |
| Erstellung Vorlagen (übergreifend) (max. 8h/Monat pro MA) | 8 h/Monat |

#### Typ Q – Sonderprojekt (Quartalsbudget)

- Nur für **0000-Sonderprojekte** mit festem **Quartalsbudget** (z.B. 4h/Quartal)
- Vergleich: **kumulierte** Ist-Stunden bis einschließlich des aktuellen Monats vs. Quartalsbudget
- Bonusberechtigt: wenn `Kumuliert-Ist / Quartalsbudget < 100%`

**0000-Meilensteine mit Quartalsbudget:**

| Meilenstein | Budget |
|------------|--------|
| Firmenveranstaltungen (max. 4h/Quartal pro MA) | 4 h/Quartal |
| Vorträge, Repräsentation (übergreifend) (max. 4h/Quartal pro MA) | 4 h/Quartal |
| Messeauftritt (max. 4h/Quartal pro MA) | 4 h/Quartal |

---

### Farbkennzeichnung (Prozentspalte)

| Bereich | Farbe | Bedeutung |
|---------|-------|-----------|
| < 90% | Grün | Budget deutlich unterschritten |
| 90% – 100% | Gelb | Budget nahezu erreicht |
| > 100% | Rot | Budget überschritten, kein Bonus |

---

### Berechnungsbeispiele

#### Beispiel 1: Reguläres Projekt – im Budget

```
Projekt: 5678.01 Brandschutzkonzept
Typ: G
Soll: 100 h
Ist (laut Budget): 85 h
Gebuchte Stunden Januar: 40 h
Prozent: 85%

→ Bonusberechtigt: 40 h (regulär)
```

#### Beispiel 2: Quartals-Sonderprojekt

```
Projekt: 0000
Meilenstein: Firmenveranstaltungen (max. 4h/Quartal pro MA)
Typ: Q
Quartals-Soll: 4 h

Oktober:   1,5 h gebuchte Stunden → kumuliert 1,5 h → 37,5% → Bonus 1,5 h
November:  1,5 h gebuchte Stunden → kumuliert 3,0 h → 75,0% → Bonus 1,5 h
Dezember:  1,0 h gebuchte Stunden → kumuliert 4,0 h → 100,0% → Bonus 1,0 h

→ Gesamt bonusberechtigt: 4 h
```

#### Beispiel 3: Budget überschritten

```
Projekt: 4321.02 Gutachten
Typ: G
Soll: 50 h
Ist (laut Budget): 55 h
Prozent: 110%

→ NICHT bonusberechtigt (Budget überschritten)
→ Manuelle Anpassung möglich über Spalte "Bonus-Anpassung"
```

---

### Bonus-Anpassungen

Die Spalte **„Bonus-Anpassung (h)"** in der Excel-Datei ermöglicht manuelle Korrekturen.

**Wann nötig?**
- Fehlbuchungen wurden nachträglich identifiziert
- Bonus trotz 100%-Auslastung nach Rücksprache mit Projektleitung
- Anteilige Anpassungen bei besonderen Umständen

**Wie?**
- Positive Werte (+) erhöhen die Bonusstunden
- Negative Werte (-) verringern die Bonusstunden
- Formel: `Bonus gesamt = Bonus Basis + Summe(Anpassungen)`

Anpassungen werden automatisch in die Monatssumme und ins Deckblatt übernommen.

**Wichtig:** Anpassungen für reguläre Projekte (Typ G) und Sonderprojekte (Typ M/Q) werden getrennt summiert und separat ausgewiesen.

---

## Ausgabedatei verstehen

### Struktur der Excel-Datei

Die generierte Excel-Datei enthält (je nach Report-Konfiguration):

| Blatt | Inhalt |
|-------|--------|
| **Deckblatt** | Gesamtübersicht aller Mitarbeiter mit Monats- und Quartalssummen |
| **Projekt-Budget-Übersicht** | Alle Projekte aus der CSV mit Budgets und Stundensätzen |
| **[Mitarbeitername]** | Ein Blatt pro Mitarbeiter mit monatlichen Detaildaten |

---

### Deckblatt (Übersichtsblatt)

Das erste Blatt zeigt eine Gesamtübersicht:

- Monatliche Summen über **alle** Mitarbeiter (Gesamtstunden, Bonusstunden regulär, Bonusstunden Sonderprojekt)
- Quartalssummen
- Liste aller Mitarbeiter im Quartal

Alle Werte sind **dynamische Excel-Formeln** – Änderungen in den Mitarbeiterblättern (z.B. Bonus-Anpassungen) werden automatisch übernommen.

---

### Mitarbeiterblätter

Pro Mitarbeiter ein Blatt. Aufbau pro Zeitblock (Monat/Woche/Block):

#### Spalten

| Spalte | Bedeutung |
|--------|-----------|
| **Projekt** | Projektname oder -nummer |
| **Meilenstein** | Arbeitspaket/Meilensteinname |
| **Typ** | `G` = regulär, `M` = 0000 Monatsbudget, `Q` = 0000 Quartalsbudget |
| **Soll (h)** | Budget-Sollstunden |
| **Ist (h)** | Verbrauchte Ist-Stunden laut Budget (bei Q: kumuliert über Quartal) |
| **[Monat] (h)** | Tatsächlich gebuchte Stunden im Block |
| **%** | Budget-Auslastung (farblich markiert) |
| **Bonus-Anpassung (h)** | Editierbar – für manuelle Korrekturen |

#### Summenzeilenbereich pro Block

- **Summe** – Gesamtstunden des Blocks
- **Bonusberechtigte Stunden** – reguläre Bonusstunden (Basis / Anpassung / Gesamt)
- **Bonusberechtigte Stunden Sonderprojekt** – 0000-Bonusstunden (Basis / Anpassung / Gesamt)

---

### Übertragshilfe (Standard-Quartalsbericht)

Die letzte Tabelle auf jedem Mitarbeiterblatt erleichtert das Übertragen in die Konzernvorlage:

| Spalte | Bedeutung |
|--------|-----------|
| **Monat** | Monat (z.B. Oktober 2025) |
| **Mitarbeiter** | Mitarbeitername |
| **Prod. Stunden** | Gesamtstunden des Monats |
| **Bonusberechigte Stunden** | Reguläre Bonusstunden inkl. Anpassungen |
| **Bonusberechtigte Stunden Sonderprojekt** | 0000-Bonusstunden inkl. Anpassungen |

**Verwendung:** Zeile markieren → in Konzernvorlage einfügen.

---

## Häufige Probleme

### „Job ist fehlgeschlagen" / „Interner Fehler"

**Ursachen:**
- CSV hat nicht die erwarteten Spalten (`Projekte`, `Arbeitspaket`, etc.)
- XML enthält keine Daten für das gewählte Quartal
- Ungültiges Datumsformat

**Lösung:**
1. Prüfen Sie, ob die CSV-Spalten korrekt sind (besonders `Projekte` und `Arbeitspaket`)
2. Lassen Sie das Quartal automatisch erkennen (Feld leer lassen)
3. Prüfen Sie, ob die XML Zeiteinträge im richtigen Zeitraum enthält

---

### „Noch keine Budget-CSV hinterlegt" (Standard-Report)

**Ursache:** Keine zentrale CSV im Admin-Bereich hinterlegt und keine CSV beim Upload angegeben.

**Lösung:** Entweder beim Standard-Report eine CSV hochladen, oder im Admin-Bereich eine zentrale CSV hinterlegen.

---

### „Mitarbeiter fehlt in der Ausgabe"

**Ursache:** Der Mitarbeiter hat im gewählten Zeitraum keine Buchungen in der XML.

**Lösung:** Prüfen Sie, ob der Mitarbeiter in der XML enthalten ist. Beim Flexiblen Report: Zeitraum prüfen.

---

### Bonusstunden stimmen nicht

**Ursache:** Typ-Spalte (G/M/Q) wird falsch zugeordnet, oder Budgetwerte in der CSV sind inkorrekt.

**Lösung:**
1. Prüfen Sie, ob Quartalsmeilensteine das Wort „Quartal" im Namen enthalten
2. Prüfen Sie die Soll-Werte in der CSV
3. Nutzen Sie die Spalte „Bonus-Anpassung" für manuelle Korrekturen

---

### CSV und XML passen nicht zusammen (kaum Daten)

**Ursache:** Projektnamen/-codes in CSV und XML weichen voneinander ab (z.B. `1234.01` vs. `1234.01 Projektname`).

**Lösung:**
- Projektnamen in CSV und XML vergleichen
- Das Tool verwendet den ersten Token (Code vor dem Leerzeichen) als Fallback-Schlüssel – kurze Codes funktionieren meistens besser

---

### Prozentspalte ist nicht farbig

**Ursache:** Kein Soll-Budget definiert (Soll = 0). Ohne Budget kann keine Auslastung berechnet werden.

**Erklärung:** Bei Typ G mit Soll = 0 greift keine Farbmarkierung. Bei Typ M/Q werden die fest hinterlegten Budgets verwendet.

---

### Admin: „Ungültige Zugangsdaten"

**Ursache:** Falsches Passwort oder falscher Benutzername.

**Hinweis:** Nach 5 Fehlversuchen innerhalb einer Minute wird der Zugang für 60 Sekunden gesperrt (Sicherheitsmechanismus). Danach normal weiterprobieren.

---

### Admin: „Zu viele Anfragen"

**Ursache:** Sicherheitsmechanismus ausgelöst (5 Fehlversuche/Minute).

**Lösung:** 60 Sekunden warten, dann erneut versuchen.

---

## Kontakt und Support

Bei technischen Problemen wenden Sie sich an Ihren Administrator.

---

*Version: 2.0 · Februar 2026*
