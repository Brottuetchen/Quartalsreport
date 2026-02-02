# Benutzeranleitung: Quartalsreport Generator

## Inhaltsverzeichnis

1. [Ãœbersicht](#Ã¼bersicht)
2. [BenÃ¶tigte Dateien](#benÃ¶tigte-dateien)
3. [Verwendung des Tools](#verwendung-des-tools)
4. [Berechnungslogik](#berechnungslogik)
5. [Bonus-Anpassungen](#bonus-anpassungen)
6. [Limits und Budgets](#limits-und-budgets)
7. [Ausgabedatei verstehen](#ausgabedatei-verstehen)
8. [HÃ¤ufige Probleme](#hÃ¤ufige-probleme)

---

## Ãœbersicht

Der Quartalsreport Generator erstellt automatisiert Excel-Berichte fÃ¼r die quartalsweise Bonusabrechnung. Das Tool vergleicht Soll- und Ist-Stunden aus der Projektverwaltung mit den tatsÃ¤chlich gebuchten Zeiten und berechnet daraus bonusberechtigte Stunden.

### Was macht das Tool?

- Vergleicht Soll/Ist-Budgets mit gebuchten Stunden
- Identifiziert bonusberechtigte Zeiten
- Unterscheidet zwischen regulÃ¤ren und Sonderprojekten
- Erzeugt ÃœbersichtsblÃ¤tter fÃ¼r alle Mitarbeiter
- ErmÃ¶glicht manuelle Bonus-Anpassungen
- Erstellt Ãœbertragshilfen fÃ¼r die Konzernvorlage

---

## BenÃ¶tigte Dateien

### 1. Zentrale Soll-Ist CSV-Datei

**Beschreibung:** Exportdatei aus der Projektverwaltung mit Budget-Informationen Ã¼ber das gesamte Unternehmen.

**Erforderliche Spalten:**
- `Projekte` â€“ Projektname oder Projektnummer
- `Arbeitspaket` â€“ Meilenstein-/Arbeitspaketname
- `Sollstunden Budget` â€“ Geplante Stunden (Format: deutsche Zahlen mit Komma)
- `Iststunden` â€“ Verbrauchte Stunden laut Budget (Format: deutsche Zahlen mit Komma)

**Format:**
- Dateiformat: `.csv`
- Trennzeichen: Tab (`\t`) oder Semikolon (`;`)
- Kodierung: UTF-16 oder UTF-8
- Zahlenformat: Deutsch (z.B. `1.234,56`)

**Hinweis:** Wie diese Datei aus der Projektverwaltung exportiert wird, wird separat behandelt.

### 2. Zeiterfassungs-XML-Datei

**Beschreibung:** XML-Export der gebuchten ZeiteintrÃ¤ge fÃ¼r die gewÃ¼nschte Mitarbeitergruppe im betrachteten Quartal.

**Erforderliche Informationen:**
- Mitarbeitername
- Projekt/Projektnummer
- Meilenstein/Arbeitspaket
- Gebuchte Stunden
- Buchungsdatum/-periode

**Format:**
- Dateiformat: `.xml`
- Muss ZeiteintrÃ¤ge des gewÃ¼nschten Quartals enthalten

**Hinweis:** Wie diese Datei aus der Zeiterfassung exportiert wird, wird separat behandelt.

---

## Verwendung des Tools

### Zugriff auf das Tool

Das Tool ist als Webdienst verfÃ¼gbar. Ã–ffnen Sie Ihren Browser und navigieren Sie zur URL des Quartalsreport Generators (z.B. `http://localhost:9999` oder die von Ihrem Administrator bereitgestellte Adresse).

### Schritt-fÃ¼r-Schritt-Anleitung

#### 1. Dateien hochladen

![Upload-Formular](docs/screenshot-upload.png)

1. Klicken Sie auf **CSV (Soll/Ist)** und wÃ¤hlen Sie Ihre Soll-Ist-CSV-Datei aus
2. Klicken Sie auf **XML (ZeiteintrÃ¤ge)** und wÃ¤hlen Sie Ihre Zeiterfassungs-XML-Datei aus
3. (Optional) Geben Sie das gewÃ¼nschte Quartal an, z.B. `2025Q3` oder `Q3-2025`
   - Wenn leer gelassen, wÃ¤hlt das Tool automatisch das aktuellste Quartal aus der XML (neueste Buchungsperiode)

#### 2. Report erzeugen

1. Klicken Sie auf den Button **Report erzeugen**
2. Der Upload-Fortschritt wird angezeigt
3. Die Verarbeitung beginnt automatisch

#### 3. Verarbeitung verfolgen

WÃ¤hrend der Verarbeitung sehen Sie:
- **Fortschrittsbalken** â€“ Zeigt den aktuellen Bearbeitungsstand
- **Statusmeldung** â€“ Informiert Ã¼ber den aktuellen Verarbeitungsschritt
- **Warteschlange** â€“ Position in der Warteschlange (falls mehrere Jobs aktiv)

#### 4. Ergebnis herunterladen

Nach erfolgreicher Verarbeitung:
1. Der Button **Ergebnis herunterladen** erscheint
2. Klicken Sie darauf, um die Excel-Datei herunterzuladen
3. Die Datei heiÃŸt z.B. `Q3-2025.xlsm`

#### 5. Neuen Report starten

- Klicken Sie auf **Neuen Report starten**, um weitere Reports zu erzeugen

---

## Berechnungslogik

### Grundprinzip

Das Tool berechnet bonusberechtigte Stunden nach folgendem Prinzip:

**Bonusberechtigt sind alle Stunden, bei denen das Budget nicht zu 100% ausgeschÃ¶pft wurde.**

Sobald das Soll eines Eintrags Ã¼berschritten wird, ist der gesamte Eintrag nicht bonusberechtigt â€“ eine anteilige Anerkennung findet nicht statt, auch wenn ein Teil der Stunden noch im Budget liegt. Dies muss dann manuell bearbeitet werden.

### Berechnung

### Meilenstein-Typen

Die Spalte **Typ** in der Excel-Ausgabe unterscheidet drei FÃ¤lle:

#### Typ G â€“ RegulÃ¤re Projekte

- Alle Projekte, deren Nummer **nicht** mit `0000` beginnt, werden als `G` gekennzeichnet.
- Die Soll-/Ist-Werte stammen vollstÃ¤ndig aus der CSV-Datei; es gibt keine festen Standard-Budgets.
- Bonusberechtigt: Solange `(Ist / Soll) < 100%`. Sobald das Soll Ã¼berschritten ist, fÃ¤llt der komplette Eintrag aus der Bonusberechtigung â€“ es gibt keine anteilige Anerkennung.
- Die interne Monats-/Quartalslogik bleibt zwar bestehen, wird aber nicht separat in der Typ-Spalte ausgewiesen.

#### Typ M â€“ Sonderprojekt (Monat)

- `M` wird ausschlieÃŸlich fÃ¼r **0000-Sonderprojekte** genutzt, die Ã¼ber ein festes **Monatsbudget** verfÃ¼gen (z.B. 8h/Monat pro Mitarbeiter).
- Vergleich: Ist-Stunden des Monats vs. Monatsbudget.
- Bonusberechtigt: Wenn `(Ist / Monatsbudget) < 100%`.
- Das Tool ergÃ¤nzt fehlende Soll-/Ist-Werte automatisch anhand der hinterlegten Budgets (siehe *Limits und Budgets*).

#### Typ Q â€“ Sonderprojekt (Quartal)

- `Q` wird ebenfalls nur fÃ¼r **0000-Sonderprojekte** verwendet, die quartalsweise Budgets haben (z.B. 4h/Quartal).
- Vergleich: Kumulierte Ist-Stunden bis einschlieÃŸlich des aktuellen Monats vs. Quartalsbudget.
- Bonusberechtigt: Wenn `(Kumuliert Ist / Quartalsbudget) < 100%`.
- Erkennungsmerkmal: Das Wort â€žQuartalâ€œ im Meilensteinnamen oder ein hinterlegtes Quartalsbudget.

### Sonderprojekte (0000-Projekte)

Projekte, die mit `0000` beginnen, werden als **Sonderprojekte** behandelt:

- Gelten als bonusberechtigt (sofern Budget nicht ausgeschÃ¶pft)
- Werden separat ausgewiesen als "Bonusberechtigte Stunden Sonderprojekt"
- Haben oft feste monatliche oder quartalsweise Budgets

**Typische 0000-Meilensteine:**

| Meilenstein | Typ | Budget |
|------------|-----|--------|
| Einarbeitung neuer Mitarbeiter | Monat | 8h/Monat |
| Angebote-Ausschreibungen-Kalkulationen | Monat | 8h/Monat |
| Erstellung Vorlagen (Ã¼bergreifend) | Monat | 8h/Monat |
| Firmenveranstaltungen | Quartal | 4h/Quartal |
| VortrÃ¤ge, ReprÃ¤sentation (Ã¼bergreifend) | Quartal | 4h/Quartal |
| Messeauftritt | Quartal | 4h/Quartal |

### Berechnungsbeispiele

#### Beispiel 1: Monatsmeilenstein, nicht ausgeschÃ¶pft

```
Projekt: 5678
Meilenstein: Testing
Typ: G (regulär, Monatslogik)
Soll: 100h
Ist: 85h
Gebuchte Stunden im Januar: 85h
Prozent: 85%

â†’ Bonusberechtigt: 85 Stunden (regulÃ¤r)
```

#### Beispiel 2: Quartalsmeilenstein, teilweise ausgeschÃ¶pft

```
Projekt: 0000
Meilenstein: Messeauftritt (max. 4h/Quartal pro MA)
Typ: Q (Sonderprojekt, quartalsweise)
Quartals-Soll: 4h
Januar gebuchte Stunden: 2h
Februar gebuchte Stunden: 1h
MÃ¤rz gebuchte Stunden: 0.5h
Kumuliert bis MÃ¤rz: 3.5h
Prozent: 87.5%

â†’ Januar: Bonusberechtigt 2h (Sonderprojekt)
â†’ Februar: Bonusberechtigt 1h (Sonderprojekt)
â†’ MÃ¤rz: Bonusberechtigt 0.5h (Sonderprojekt)
```

#### Beispiel 3: Budget zu 100% oder mehr ausgeschÃ¶pft

```
Projekt: 4321
Meilenstein: Dokumentation
Typ: G (regulär, Monatslogik)
Soll: 50h
Ist: 55h
Gebuchte Stunden im Februar: 55h
Prozent: 110%

â†’ NICHT bonusberechtigt (Budget Ã¼berschritten)
```

### Farbkennzeichnung

Die Prozentspalte wird zur besseren Ãœbersicht farblich markiert:

| Prozentbereich | Farbe | Bedeutung |
|---------------|-------|-----------|
| < 90% | ðŸŸ¢ GrÃ¼n | Budget deutlich unterschritten |
| 90% - 100% | ðŸŸ¡ Gelb | Budget nahezu erreicht |
| > 100% | ðŸ”´ Rot | Budget Ã¼berschritten |

---

## Bonus-Anpassungen

### Zweck der Bonus-Anpassung

Die Spalte **Bonus-Anpassung (h)** ermÃ¶glicht **manuelle Korrekturen** der automatisch berechneten bonusberechtigten Stunden.

### Wann werden Anpassungen benÃ¶tigt?

- NachtrÃ¤gliche Korrekturen aufgrund von Fehlbuchungen
- Manuelle BonusgewÃ¤hrung trotz 100% Budget-Auslastung
- AbzÃ¼ge bei besonderen UmstÃ¤nden
- Korrekturen nach RÃ¼cksprache mit Projektleitung

### Wie funktionieren Anpassungen?

1. **Positive Werte** (+) erhÃ¶hen die bonusberechtigten Stunden
2. **Negative Werte** (-) verringern die bonusberechtigten Stunden
3. Anpassungen werden **automatisch summiert** und zur Basis-Bonusberechnung addiert

**Formel:**
```
Bonusberechtigte Stunden (Gesamt) = Bonusberechtigte Stunden (Basis) + Summe(Bonus-Anpassungen)
```

### Beispiel fÃ¼r Anpassungen

**Ausgangssituation:**

| Meilenstein | Typ | Gebuchte Stunden | Bonus (Basis) | Bonus-Anpassung | Bonus (Gesamt) |
|------------|-----|------------------|---------------|-----------------|----------------|
| Entwicklung | G | 75h | 75h | 0 | 75h |
| Testing | G | 50h | 50h | 0 | 50h |
| **Summe** | | **125h** | **125h** | **0** | **125h** |

**Nach Anpassung:**

Sie tragen in der Spalte "Bonus-Anpassung" folgende Werte ein:
- Entwicklung: `-10` (Fehlbuchung wurde identifiziert)
- Testing: `+5` (NachtrÃ¤gliche BonusgewÃ¤hrung nach RÃ¼cksprache)

| Meilenstein | Typ | Gebuchte Stunden | Bonus (Basis) | Bonus-Anpassung | Bonus (Gesamt) |
|------------|-----|------------------|---------------|-----------------|----------------|
| Entwicklung | G | 75h | 75h | -10 | 65h |
| Testing | G | 50h | 50h | +5 | 55h |
| **Summe** | | **125h** | **125h** | **-5** | **120h** |

**Wichtig:**
- Anpassungen werden **automatisch** in die Monatssumme Ã¼bernommen
- Die Quartals-Gesamtsumme wird ebenfalls automatisch aktualisiert
- Die Ãœbertragshilfe berÃ¼cksichtigt die angepassten Werte

### Getrennte Anpassungen: RegulÃ¤r vs. Sonderprojekt

- Anpassungen fÃ¼r **regulÃ¤re Projekte** beeinflussen "Bonusberechtigte Stunden"
- Anpassungen fÃ¼r **0000-Sonderprojekte** beeinflussen "Bonusberechtigte Stunden Sonderprojekt"
- Die Trennung erfolgt automatisch anhand der Projektnummer

---

## Limits und Budgets

### Monatliche Budgets (0000-Projekte)

Folgende Meilensteine haben **feste monatliche Budgets** pro Mitarbeiter:

| Meilenstein | Budget pro Monat |
|------------|------------------|
| Einarbeitung neuer Mitarbeiter (max. 8h/Monat pro MA) | 8 Stunden |
| Angebote-Ausschreibungen-Kalkulationen (max. 8h/Monat pro MA) | 8 Stunden |
| Erstellung Vorlagen (Ã¼bergreifend) (max. 8h/Monat pro MA) | 8 Stunden |

**Verhalten:**
- Das Tool setzt automatisch `Soll = 8h` und `Ist = Gebuchte Stunden`
- Wenn mehr als 8h gebucht werden, sind nur die ersten 8h bonusberechtigt

### Quartalsbudgets (0000-Projekte)

Folgende Meilensteine haben **feste Quartalsbudgets** pro Mitarbeiter:

| Meilenstein | Budget pro Quartal |
|------------|-------------------|
| Firmenveranstaltungen (max. 4h/Quartal pro MA) | 4 Stunden |
| VortrÃ¤ge, ReprÃ¤sentation (Ã¼bergreifend) (max. 4h/Quartal pro MA) | 4 Stunden |
| Messeauftritt (max. 4h/Quartal pro MA) | 4 Stunden |

**Verhalten:**
- BudgetprÃ¼fung erfolgt **kumuliert** Ã¼ber das gesamte Quartal
- Erst wenn 4h im Quartal erreicht sind, wird der Meilenstein zu 100% ausgelastet
- Ãœberstunden sind nicht bonusberechtigt

### Budgets aus CSV Ã¼berschreiben

Wenn in der CSV-Datei fÃ¼r 0000-Projekte Soll/Ist-Werte vorhanden sind, werden diese **nicht** Ã¼berschrieben â€“ das Tool respektiert die CSV-Werte.

Nur wenn **Soll = 0** und **Ist = 0**, greift das Tool auf die fest definierten Budgets zurÃ¼ck.

---

## Ausgabedatei verstehen

### Struktur der Excel-Datei

Die generierte Excel-Datei (z.B. `Q3-2025.xlsm`) enthÃ¤lt:

1. **Ãœbersichtsblatt** (Deckblatt)
2. **Pro Mitarbeiter ein separates Arbeitsblatt**

### 1. Ãœbersichtsblatt (Deckblatt)

Das erste Blatt zeigt eine **GesamtÃ¼bersicht** Ã¼ber alle Mitarbeiter:

**Inhalt:**

- **Monatliche Summen** Ã¼ber alle Mitarbeiter:
  - Gesamtstunden
  - Bonusberechtigte Stunden
  - Bonusberechtigte Stunden Sonderprojekt
- **Quartalssummen** Ã¼ber alle Mitarbeiter
- **Liste aller Mitarbeiter** im Quartal

**Besonderheit:** Alle Werte sind **dynamische Excel-Formeln**, die sich automatisch aktualisieren, wenn Ã„nderungen in den MitarbeiterblÃ¤ttern vorgenommen werden.

### 2. MitarbeiterblÃ¤tter

FÃ¼r jeden Mitarbeiter im Quartal wird ein separates Arbeitsblatt erstellt.

**Aufbau pro Monat:**

#### Tabellenkopf

| Spalte | Bedeutung |
|--------|-----------|
| **Projekt** | Projektname oder -nummer |
| **Meilenstein** | Arbeitspaket/Meilensteinname |
| **Typ** | `G` = reguläres Projekt, `M` = 0000-Monatsbudget, `Q` = 0000-Quartalsbudget |
| **Soll (h)** | Budget-Sollstunden (Monat/Quartal) |
| **Ist (h)** | Verbrauchte Ist-Stunden laut Budget (Monat) oder kumuliert (Quartal) |
| **[Monat] (h)** | TatsÃ¤chlich gebuchte Stunden im jeweiligen Monat |
| **%** | Prozentsatz der Budget-Auslastung (farblich markiert) |
| **Bonus-Anpassung (h)** | Feld fÃ¼r manuelle Korrekturen |

#### Summenwerte (pro Monat)

- **Summe** â€“ Gesamtstunden des Monats
- **Bonusberechtigte Stunden** â€“ Automatisch berechnete bonusberechtigte Stunden (regulÃ¤r)
  - Spalte 7: Basis-Wert
  - Spalte 8: Summe der Anpassungen
  - Spalte 6: **Gesamt-Wert** (= Basis + Anpassungen)
- **Bonusberechtigte Stunden Sonderprojekt** â€“ Automatisch berechnete bonusberechtigte Stunden (0000-Projekte)
  - Spalte 7: Basis-Wert
  - Spalte 8: Summe der Anpassungen
  - Spalte 6: **Gesamt-Wert** (= Basis + Anpassungen)

#### Quartalszusammenfassung

Am Ende jedes Mitarbeiterblattes:

- **Quartalsmeilensteine mit Quartalssoll** â€“ Ãœbersicht aller Q-Meilensteine
- **Gesamtstunden (Quartal)** â€“ Summe aller gebuchten Stunden
- **Bonusberechtigte Stunden (Quartal)** â€“ Quartalssumme regulÃ¤rer Bonusstunden
- **Bonusberechtigte Stunden Sonderprojekt (Quartal)** â€“ Quartalssumme 0000-Bonusstunden

#### Ãœbertragshilfe

Die letzte Tabelle "**Ãœbertragshilfe**" erleichtert das Ãœbertragen in die Konzernvorlage:

| Spalte | Bedeutung |
|--------|-----------|
| **Monat** | Monat (z.B. Januar 2025) |
| **Mitarbeiter** | Mitarbeitername |
| **Prod. Stunden** | Produktive Stunden (Gesamtstunden des Monats) |
| **Bonusberechtigte Stunden** | Bonusstunden regulÃ¤r (inklusive Anpassungen) |
| **Bonusberechtigte Stunden Sonderprojekt** | Bonusstunden 0000-Projekte (inklusive Anpassungen) |

**Verwendung:**
- Markieren Sie die Zeile fÃ¼r den gewÃ¼nschten Monat
- Kopieren Sie die Werte in Ihre Konzernvorlage

---

## HÃ¤ufige Probleme

### Problem: "Job ist fehlgeschlagen"

**Ursachen:**
- CSV-Datei hat nicht die erwarteten Spalten
- XML-Datei ist fehlerhaft oder leer
- Keine Daten fÃ¼r das gewÃ¤hlte Quartal vorhanden

**LÃ¶sung:**
1. ÃœberprÃ¼fen Sie, ob die CSV-Datei die Spalten `Projekte`, `Arbeitspaket`, `Sollstunden Budget`, `Iststunden` enthÃ¤lt
2. ÃœberprÃ¼fen Sie, ob die XML-Datei ZeiteintrÃ¤ge fÃ¼r das gewÃ¼nschte Quartal enthÃ¤lt
3. Versuchen Sie, das Quartal automatisch wÃ¤hlen zu lassen (Feld leer lassen)

### Problem: "Mitarbeiter fehlt in der Ausgabe"

**Ursache:**
- Der Mitarbeiter hat im betrachteten Quartal keine ZeiteintrÃ¤ge in der XML-Datei

**LÃ¶sung:**
- ÃœberprÃ¼fen Sie, ob der Mitarbeiter in der XML-Datei enthalten ist
- Stellen Sie sicher, dass die XML alle gewÃ¼nschten Mitarbeiter enthÃ¤lt

### Problem: "Bonusberechtigte Stunden stimmen nicht"

**Ursache:**
- Typ-Spalte (G/M/Q) wird falsch zugeordnet
- Budget-Werte aus CSV sind inkorrekt

**LÃ¶sung:**
1. ÃœberprÃ¼fen Sie, ob Quartalsmeilensteine das Wort "Quartal" im Namen enthalten
2. PrÃ¼fen Sie die Soll/Ist-Werte in der CSV-Datei
3. Nutzen Sie die Spalte "Bonus-Anpassung" fÃ¼r manuelle Korrekturen

### Problem: "Excel-Datei enthÃ¤lt kaum Daten"

**Ursache:**
- CSV- und XML-Projekte/Meilensteine stimmen nicht Ã¼berein
- Normalisierung der Projektnamen schlÃ¤gt fehl

**LÃ¶sung:**
- Stellen Sie sicher, dass Projektnamen/-nummern in CSV und XML Ã¼bereinstimmen
- ÃœberprÃ¼fen Sie, ob Meilensteinnamen konsistent sind

### Problem: "Prozentspalte ist nicht farbig"

**Ursache:**
- Kein Soll-Budget definiert oder Soll = 0

**ErklÃ¤rung:**
- Farbmarkierung erfolgt nur, wenn ein Budget (Soll > 0) definiert ist
- Ohne Budget kann keine Prozent-Auslastung berechnet werden

---

## Kontakt und Support

Bei technischen Problemen oder Fragen zur Nutzung wenden Sie sich bitte an Ihren IT-Administrator oder die verantwortliche Fachabteilung.

---

**Version:** 1.0
**Datum:** Januar 2025
**Tool-Version:** Siehe README.md


