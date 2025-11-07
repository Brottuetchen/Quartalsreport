# Benutzeranleitung: Quartalsreport Generator

## Inhaltsverzeichnis

1. [Übersicht](#übersicht)
2. [Benötigte Dateien](#benötigte-dateien)
3. [Verwendung des Tools](#verwendung-des-tools)
4. [Berechnungslogik](#berechnungslogik)
5. [Bonus-Anpassungen](#bonus-anpassungen)
6. [Limits und Budgets](#limits-und-budgets)
7. [Ausgabedatei verstehen](#ausgabedatei-verstehen)
8. [Häufige Probleme](#häufige-probleme)

---

## Übersicht

Der Quartalsreport Generator erstellt automatisiert Excel-Berichte für die quartalsweise Bonusabrechnung. Das Tool vergleicht Soll- und Ist-Stunden aus der Projektverwaltung mit den tatsächlich gebuchten Zeiten und berechnet daraus bonusberechtigte Stunden.

### Was macht das Tool?

- Vergleicht Soll/Ist-Budgets mit gebuchten Stunden
- Identifiziert bonusberechtigte Zeiten
- Unterscheidet zwischen regulären und Sonderprojekten
- Erzeugt Übersichtsblätter für alle Mitarbeiter
- Ermöglicht manuelle Bonus-Anpassungen
- Erstellt Übertragshilfen für die Konzernvorlage

---

## Benötigte Dateien

### 1. Zentrale Soll-Ist CSV-Datei

**Beschreibung:** Exportdatei aus der Projektverwaltung mit Budget-Informationen über das gesamte Unternehmen.

**Erforderliche Spalten:**
- `Projekte` – Projektname oder Projektnummer
- `Arbeitspaket` – Meilenstein-/Arbeitspaketname
- `Sollstunden Budget` – Geplante Stunden (Format: deutsche Zahlen mit Komma)
- `Iststunden` – Verbrauchte Stunden laut Budget (Format: deutsche Zahlen mit Komma)

**Format:**
- Dateiformat: `.csv`
- Trennzeichen: Tab (`\t`) oder Semikolon (`;`)
- Kodierung: UTF-16 oder UTF-8
- Zahlenformat: Deutsch (z.B. `1.234,56`)

**Hinweis:** Wie diese Datei aus der Projektverwaltung exportiert wird, wird separat behandelt.

### 2. Zeiterfassungs-XML-Datei

**Beschreibung:** XML-Export der gebuchten Zeiteinträge für die gewünschte Mitarbeitergruppe im betrachteten Quartal.

**Erforderliche Informationen:**
- Mitarbeitername
- Projekt/Projektnummer
- Meilenstein/Arbeitspaket
- Gebuchte Stunden
- Buchungsdatum/-periode

**Format:**
- Dateiformat: `.xml`
- Muss Zeiteinträge des gewünschten Quartals enthalten

**Hinweis:** Wie diese Datei aus der Zeiterfassung exportiert wird, wird separat behandelt.

---

## Verwendung des Tools

### Zugriff auf das Tool

Das Tool ist als Webdienst verfügbar. Öffnen Sie Ihren Browser und navigieren Sie zur URL des Quartalsreport Generators (z.B. `http://localhost:9999` oder die von Ihrem Administrator bereitgestellte Adresse).

### Schritt-für-Schritt-Anleitung

#### 1. Dateien hochladen

![Upload-Formular](docs/screenshot-upload.png)

1. Klicken Sie auf **CSV (Soll/Ist)** und wählen Sie Ihre Soll-Ist-CSV-Datei aus
2. Klicken Sie auf **XML (Zeiteinträge)** und wählen Sie Ihre Zeiterfassungs-XML-Datei aus
3. (Optional) Geben Sie das gewünschte Quartal an, z.B. `2025Q3` oder `Q3-2025`
   - Wenn leer gelassen, wählt das Tool automatisch das Quartal mit den meisten Einträgen

#### 2. Report erzeugen

1. Klicken Sie auf den Button **Report erzeugen**
2. Der Upload-Fortschritt wird angezeigt
3. Die Verarbeitung beginnt automatisch

#### 3. Verarbeitung verfolgen

Während der Verarbeitung sehen Sie:
- **Fortschrittsbalken** – Zeigt den aktuellen Bearbeitungsstand
- **Statusmeldung** – Informiert über den aktuellen Verarbeitungsschritt
- **Warteschlange** – Position in der Warteschlange (falls mehrere Jobs aktiv)

#### 4. Ergebnis herunterladen

Nach erfolgreicher Verarbeitung:
1. Der Button **Ergebnis herunterladen** erscheint
2. Klicken Sie darauf, um die Excel-Datei herunterzuladen
3. Die Datei heißt z.B. `Q3-2025.xlsx`

#### 5. Neuen Report starten

- Klicken Sie auf **Neuen Report starten**, um weitere Reports zu erzeugen

---

## Berechnungslogik

### Grundprinzip

Das Tool berechnet bonusberechtigte Stunden nach folgendem Prinzip:

**Bonusberechtigt sind alle Stunden, bei denen das Budget nicht zu 100% ausgeschöpft wurde.**

### Meilenstein-Typen

Das Tool unterscheidet zwei Arten von Meilensteinen:

#### 1. Monatsmeilensteine (Typ: M)

- Budget gilt **pro Monat**
- Vergleich: Ist-Stunden des Monats vs. Soll-Budget des Monats
- Bonusberechtigt: Wenn `(Ist / Soll) < 100%`

**Beispiel:**
```
Projekt: 1234 - Kundenprojekt
Meilenstein: Entwicklung
Soll (h): 80
Ist (h): 60
Januar (h): 60
% = 75% → BONUSBERECHTIGT (60 Stunden)
```

#### 2. Quartalsmeilensteine (Typ: Q)

- Budget gilt **für das gesamte Quartal**
- Vergleich: Kumulierte Ist-Stunden bis einschließlich Monat vs. Quartals-Budget
- Bonusberechtigt: Wenn `(Kumuliert Ist / Quartals-Soll) < 100%`
- Erkennungsmerkmal: Das Wort "Quartal" im Meilensteinnamen

**Beispiel:**
```
Projekt: 0000 - Intern
Meilenstein: Firmenveranstaltungen (max. 4h/Quartal pro MA)
Quartals-Soll (h): 4
Kumuliert Ist (h): 2 (über Jan+Feb+März)
März (h): 1
% = 50% → BONUSBERECHTIGT (1 Stunde im März)
```

### Sonderprojekte (0000-Projekte)

Projekte, die mit `0000` beginnen, werden als **Sonderprojekte** behandelt:

- Gelten als bonusberechtigt (sofern Budget nicht ausgeschöpft)
- Werden separat ausgewiesen als "Bonusberechtigte Stunden Sonderprojekt"
- Haben oft feste monatliche oder quartalsweise Budgets

**Typische 0000-Meilensteine:**

| Meilenstein | Typ | Budget |
|------------|-----|--------|
| Einarbeitung neuer Mitarbeiter | Monat | 8h/Monat |
| Angebote-Ausschreibungen-Kalkulationen | Monat | 8h/Monat |
| Erstellung Vorlagen (übergreifend) | Monat | 8h/Monat |
| Firmenveranstaltungen | Quartal | 4h/Quartal |
| Vorträge, Repräsentation (übergreifend) | Quartal | 4h/Quartal |
| Messeauftritt | Quartal | 4h/Quartal |

### Berechnungsbeispiele

#### Beispiel 1: Monatsmeilenstein, nicht ausgeschöpft

```
Projekt: 5678
Meilenstein: Testing
Typ: M (monatlich)
Soll: 100h
Ist: 85h
Gebuchte Stunden im Januar: 85h
Prozent: 85%

→ Bonusberechtigt: 85 Stunden (regulär)
```

#### Beispiel 2: Quartalsmeilenstein, teilweise ausgeschöpft

```
Projekt: 0000
Meilenstein: Messeauftritt (max. 4h/Quartal pro MA)
Typ: Q (quartalsweise)
Quartals-Soll: 4h
Januar gebuchte Stunden: 2h
Februar gebuchte Stunden: 1h
März gebuchte Stunden: 0.5h
Kumuliert bis März: 3.5h
Prozent: 87.5%

→ Januar: Bonusberechtigt 2h (Sonderprojekt)
→ Februar: Bonusberechtigt 1h (Sonderprojekt)
→ März: Bonusberechtigt 0.5h (Sonderprojekt)
```

#### Beispiel 3: Budget zu 100% oder mehr ausgeschöpft

```
Projekt: 4321
Meilenstein: Dokumentation
Typ: M (monatlich)
Soll: 50h
Ist: 55h
Gebuchte Stunden im Februar: 55h
Prozent: 110%

→ NICHT bonusberechtigt (Budget überschritten)
```

### Farbkennzeichnung

Die Prozentspalte wird zur besseren Übersicht farblich markiert:

| Prozentbereich | Farbe | Bedeutung |
|---------------|-------|-----------|
| < 90% | 🟢 Grün | Budget deutlich unterschritten |
| 90% - 100% | 🟡 Gelb | Budget nahezu erreicht |
| > 100% | 🔴 Rot | Budget überschritten |

---

## Bonus-Anpassungen

### Zweck der Bonus-Anpassung

Die Spalte **Bonus-Anpassung (h)** ermöglicht **manuelle Korrekturen** der automatisch berechneten bonusberechtigten Stunden.

### Wann werden Anpassungen benötigt?

- Nachträgliche Korrekturen aufgrund von Fehlbuchungen
- Manuelle Bonusgewährung trotz 100% Budget-Auslastung
- Abzüge bei besonderen Umständen
- Korrekturen nach Rücksprache mit Projektleitung

### Wie funktionieren Anpassungen?

1. **Positive Werte** (+) erhöhen die bonusberechtigten Stunden
2. **Negative Werte** (-) verringern die bonusberechtigten Stunden
3. Anpassungen werden **automatisch summiert** und zur Basis-Bonusberechnung addiert

**Formel:**
```
Bonusberechtigte Stunden (Gesamt) = Bonusberechtigte Stunden (Basis) + Summe(Bonus-Anpassungen)
```

### Beispiel für Anpassungen

**Ausgangssituation:**

| Meilenstein | Typ | Gebuchte Stunden | Bonus (Basis) | Bonus-Anpassung | Bonus (Gesamt) |
|------------|-----|------------------|---------------|-----------------|----------------|
| Entwicklung | M | 75h | 75h | 0 | 75h |
| Testing | M | 50h | 50h | 0 | 50h |
| **Summe** | | **125h** | **125h** | **0** | **125h** |

**Nach Anpassung:**

Sie tragen in der Spalte "Bonus-Anpassung" folgende Werte ein:
- Entwicklung: `-10` (Fehlbuchung wurde identifiziert)
- Testing: `+5` (Nachträgliche Bonusgewährung nach Rücksprache)

| Meilenstein | Typ | Gebuchte Stunden | Bonus (Basis) | Bonus-Anpassung | Bonus (Gesamt) |
|------------|-----|------------------|---------------|-----------------|----------------|
| Entwicklung | M | 75h | 75h | -10 | 65h |
| Testing | M | 50h | 50h | +5 | 55h |
| **Summe** | | **125h** | **125h** | **-5** | **120h** |

**Wichtig:**
- Anpassungen werden **automatisch** in die Monatssumme übernommen
- Die Quartals-Gesamtsumme wird ebenfalls automatisch aktualisiert
- Die Übertragshilfe berücksichtigt die angepassten Werte

### Getrennte Anpassungen: Regulär vs. Sonderprojekt

- Anpassungen für **reguläre Projekte** beeinflussen "Bonusberechtigte Stunden"
- Anpassungen für **0000-Sonderprojekte** beeinflussen "Bonusberechtigte Stunden Sonderprojekt"
- Die Trennung erfolgt automatisch anhand der Projektnummer

---

## Limits und Budgets

### Monatliche Budgets (0000-Projekte)

Folgende Meilensteine haben **feste monatliche Budgets** pro Mitarbeiter:

| Meilenstein | Budget pro Monat |
|------------|------------------|
| Einarbeitung neuer Mitarbeiter (max. 8h/Monat pro MA) | 8 Stunden |
| Angebote-Ausschreibungen-Kalkulationen (max. 8h/Monat pro MA) | 8 Stunden |
| Erstellung Vorlagen (übergreifend) (max. 8h/Monat pro MA) | 8 Stunden |

**Verhalten:**
- Das Tool setzt automatisch `Soll = 8h` und `Ist = Gebuchte Stunden`
- Wenn mehr als 8h gebucht werden, sind nur die ersten 8h bonusberechtigt

### Quartalsbudgets (0000-Projekte)

Folgende Meilensteine haben **feste Quartalsbudgets** pro Mitarbeiter:

| Meilenstein | Budget pro Quartal |
|------------|-------------------|
| Firmenveranstaltungen (max. 4h/Quartal pro MA) | 4 Stunden |
| Vorträge, Repräsentation (übergreifend) (max. 4h/Quartal pro MA) | 4 Stunden |
| Messeauftritt (max. 4h/Quartal pro MA) | 4 Stunden |

**Verhalten:**
- Budgetprüfung erfolgt **kumuliert** über das gesamte Quartal
- Erst wenn 4h im Quartal erreicht sind, wird der Meilenstein zu 100% ausgelastet
- Überstunden sind nicht bonusberechtigt

### Budgets aus CSV überschreiben

Wenn in der CSV-Datei für 0000-Projekte Soll/Ist-Werte vorhanden sind, werden diese **nicht** überschrieben – das Tool respektiert die CSV-Werte.

Nur wenn **Soll = 0** und **Ist = 0**, greift das Tool auf die fest definierten Budgets zurück.

---

## Ausgabedatei verstehen

### Struktur der Excel-Datei

Die generierte Excel-Datei (z.B. `Q3-2025.xlsx`) enthält:

1. **Übersichtsblatt** (Deckblatt)
2. **Pro Mitarbeiter ein separates Arbeitsblatt**

### 1. Übersichtsblatt (Deckblatt)

Das erste Blatt zeigt eine **Gesamtübersicht** über alle Mitarbeiter:

**Inhalt:**

- **Monatliche Summen** über alle Mitarbeiter:
  - Gesamtstunden
  - Bonusberechtigte Stunden
  - Bonusberechtigte Stunden Sonderprojekt
- **Quartalssummen** über alle Mitarbeiter
- **Liste aller Mitarbeiter** im Quartal

**Besonderheit:** Alle Werte sind **dynamische Excel-Formeln**, die sich automatisch aktualisieren, wenn Änderungen in den Mitarbeiterblättern vorgenommen werden.

### 2. Mitarbeiterblätter

Für jeden Mitarbeiter im Quartal wird ein separates Arbeitsblatt erstellt.

**Aufbau pro Monat:**

#### Tabellenkopf

| Spalte | Bedeutung |
|--------|-----------|
| **Projekt** | Projektname oder -nummer |
| **Meilenstein** | Arbeitspaket/Meilensteinname |
| **Typ** | `M` = Monatsmeilenstein, `Q` = Quartalsmeilenstein |
| **Soll (h)** | Budget-Sollstunden (Monat/Quartal) |
| **Ist (h)** | Verbrauchte Ist-Stunden laut Budget (Monat) oder kumuliert (Quartal) |
| **[Monat] (h)** | Tatsächlich gebuchte Stunden im jeweiligen Monat |
| **%** | Prozentsatz der Budget-Auslastung (farblich markiert) |
| **Bonus-Anpassung (h)** | Feld für manuelle Korrekturen |

#### Summenwerte (pro Monat)

- **Summe** – Gesamtstunden des Monats
- **Bonusberechtigte Stunden** – Automatisch berechnete bonusberechtigte Stunden (regulär)
  - Spalte 7: Basis-Wert
  - Spalte 8: Summe der Anpassungen
  - Spalte 6: **Gesamt-Wert** (= Basis + Anpassungen)
- **Bonusberechtigte Stunden Sonderprojekt** – Automatisch berechnete bonusberechtigte Stunden (0000-Projekte)
  - Spalte 7: Basis-Wert
  - Spalte 8: Summe der Anpassungen
  - Spalte 6: **Gesamt-Wert** (= Basis + Anpassungen)

#### Quartalszusammenfassung

Am Ende jedes Mitarbeiterblattes:

- **Quartalsmeilensteine mit Quartalssoll** – Übersicht aller Q-Meilensteine
- **Gesamtstunden (Quartal)** – Summe aller gebuchten Stunden
- **Bonusberechtigte Stunden (Quartal)** – Quartalssumme regulärer Bonusstunden
- **Bonusberechtigte Stunden Sonderprojekt (Quartal)** – Quartalssumme 0000-Bonusstunden

#### Übertragshilfe

Die letzte Tabelle "**Übertragshilfe**" erleichtert das Übertragen in die Konzernvorlage:

| Spalte | Bedeutung |
|--------|-----------|
| **Monat** | Monat (z.B. Januar 2025) |
| **Mitarbeiter** | Mitarbeitername |
| **Prod. Stunden** | Produktive Stunden (Gesamtstunden des Monats) |
| **Bonusberechtigte Stunden** | Bonusstunden regulär (inklusive Anpassungen) |
| **Bonusberechtigte Stunden Sonderprojekt** | Bonusstunden 0000-Projekte (inklusive Anpassungen) |

**Verwendung:**
- Markieren Sie die Zeile für den gewünschten Monat
- Kopieren Sie die Werte in Ihre Konzernvorlage

---

## Häufige Probleme

### Problem: "Job ist fehlgeschlagen"

**Ursachen:**
- CSV-Datei hat nicht die erwarteten Spalten
- XML-Datei ist fehlerhaft oder leer
- Keine Daten für das gewählte Quartal vorhanden

**Lösung:**
1. Überprüfen Sie, ob die CSV-Datei die Spalten `Projekte`, `Arbeitspaket`, `Sollstunden Budget`, `Iststunden` enthält
2. Überprüfen Sie, ob die XML-Datei Zeiteinträge für das gewünschte Quartal enthält
3. Versuchen Sie, das Quartal automatisch wählen zu lassen (Feld leer lassen)

### Problem: "Mitarbeiter fehlt in der Ausgabe"

**Ursache:**
- Der Mitarbeiter hat im betrachteten Quartal keine Zeiteinträge in der XML-Datei

**Lösung:**
- Überprüfen Sie, ob der Mitarbeiter in der XML-Datei enthalten ist
- Stellen Sie sicher, dass die XML alle gewünschten Mitarbeiter enthält

### Problem: "Bonusberechtigte Stunden stimmen nicht"

**Ursache:**
- Meilenstein-Typ wird falsch erkannt
- Budget-Werte aus CSV sind inkorrekt

**Lösung:**
1. Überprüfen Sie, ob Quartalsmeilensteine das Wort "Quartal" im Namen enthalten
2. Prüfen Sie die Soll/Ist-Werte in der CSV-Datei
3. Nutzen Sie die Spalte "Bonus-Anpassung" für manuelle Korrekturen

### Problem: "Excel-Datei enthält kaum Daten"

**Ursache:**
- CSV- und XML-Projekte/Meilensteine stimmen nicht überein
- Normalisierung der Projektnamen schlägt fehl

**Lösung:**
- Stellen Sie sicher, dass Projektnamen/-nummern in CSV und XML übereinstimmen
- Überprüfen Sie, ob Meilensteinnamen konsistent sind

### Problem: "Prozentspalte ist nicht farbig"

**Ursache:**
- Kein Soll-Budget definiert oder Soll = 0

**Erklärung:**
- Farbmarkierung erfolgt nur, wenn ein Budget (Soll > 0) definiert ist
- Ohne Budget kann keine Prozent-Auslastung berechnet werden

---

## Kontakt und Support

Bei technischen Problemen oder Fragen zur Nutzung wenden Sie sich bitte an Ihren IT-Administrator oder die verantwortliche Fachabteilung.

---

**Version:** 1.0
**Datum:** Januar 2025
**Tool-Version:** Siehe README.md
