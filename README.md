# GANTT Generator – DROEGE GROUP

## Vollständige Dokumentation | v21.3

---

## Inhaltsverzeichnis

1. [Überblick](#überblick)
2. [Dateien](#dateien)
3. [Installation](#installation)
4. [Benutzeroberfläche](#benutzeroberfläche)
5. [Rastereinheit (RE)](#rastereinheit-re)
6. [Bildschirmformat-Erkennung](#bildschirmformat-erkennung)
7. [Grid-System & Offset-Logik](#grid-system--offset-logik)
8. [Zeitraum & Zeiteinheiten](#zeitraum--zeiteinheiten)
9. [Layout-Parameter](#layout-parameter)
10. [Spaltenbreite & Verteilung](#spaltenbreite--verteilung)
11. [Phasen](#phasen)
12. [Heute-Linie](#heute-linie)
13. [GANTT-Erstellung – Technischer Ablauf](#gantt-erstellung--technischer-ablauf)
14. [Positionierung & Snapping](#positionierung--snapping)
15. [Visuelle Struktur](#visuelle-struktur)
16. [Konstanten & Konfiguration](#konstanten--konfiguration)
17. [Funktionsreferenz (taskpane.js)](#funktionsreferenz-taskpanejs)
18. [Format-Tabelle (Grid_Resize_Tool v15)](#format-tabelle-grid_resize_tool-v15)
19. [Versionshistorie](#versionshistorie)
20. [Bekannte Einschränkungen](#bekannte-einschränkungen)
21. [Lizenz](#lizenz)

---

## Überblick

Der **GANTT Generator** ist ein **PowerPoint Office Web Add-in** zur Erstellung professioneller GANTT-Diagramme im DROEGE GROUP Corporate Design. Das Add-in erzeugt native PowerPoint-Shapes (Rechtecke, Linien, Textfelder), die nach der Erstellung frei editierbar sind.

### Kernmerkmale

- **Rasterbasiert:** Alle Positionen und Größen basieren auf Rastereinheiten (RE), die exakt dem PowerPoint-Raster entsprechen
- **Format-adaptiv:** Automatische Erkennung des Folienformats (16:9, 4:3, 16:10, A4 quer, Breitbild) mit format-spezifischen Grid-Offsets
- **API 1.10:** Nutzt PowerPointApi 1.10 zum dynamischen Auslesen der Foliengröße
- **Grid_Resize_Tool v15 Logik:** Integriert die Offset-Tabelle und 4-Stufen-Erkennung

---

## Dateien

| Datei | Beschreibung |
|---|---|
| `taskpane.js` | Kernlogik – GANTT-Berechnung, Grid-System, PowerPoint-API-Aufrufe |
| `taskpane.html` | Benutzeroberfläche – Eingabefelder, Buttons, Statusanzeige |
| `taskpane.css` | Styling – DROEGE GROUP Design, responsive Taskpane-Darstellung |
| `manifest-PPT-Addin.xml` | Office Add-in Manifest – Registrierung, Berechtigungen, Endpunkte |
| `README.md` | Diese Dokumentation |

---

## Installation

### Voraussetzungen
- PowerPoint (Desktop oder Online) mit Office Add-in-Unterstützung
- PowerPointApi 1.10 (für Foliengrößen-Erkennung)
- Webserver oder localhost zum Hosten der Add-in-Dateien

### Schritte
1. Alle Dateien (`taskpane.js`, `taskpane.html`, `taskpane.css`, `manifest-PPT-Addin.xml`) auf einen Webserver oder localhost legen
2. In PowerPoint: **Einfügen → Add-ins → Mein Add-in hochladen → `manifest-PPT-Addin.xml`** auswählen
3. Das Add-in öffnet sich im rechten Taskpane

### Manifest-Konfiguration
Die URLs im Manifest müssen auf den tatsächlichen Hosting-Ort zeigen:
```xml
<bt:Url id="Taskpane.Url" DefaultValue="https://DEIN-SERVER/taskpane.html"/>
```

---

## Benutzeroberfläche

Das Taskpane gliedert sich in folgende Bereiche (von oben nach unten):

### 1. Header
- Titel: **GANTT Generator DROEGE GROUP**
- Info-Leiste: Version, API-Status, Datum/Uhrzeit

### 2. Rastereinheit (RE)
- Eingabefeld für die Rastereinheit in cm
- Preset-Buttons: `0,2117` | `0,4233` | `0,6350` | `2,1167`
- Aktiver Preset wird visuell hervorgehoben

### 3. Bildschirmformat
- Automatische Anzeige des erkannten Formats
- Button **🔎 Erkennen** zum manuellen Auslösen der Format-Erkennung
- Anzeige der Grid-Offsets (X/Y in cm)

### 4. Zeitraum
- **Start-Datum** und **End-Datum** (Datumspicker)
- **Einheit:** Tage | Wochen | Monate | Quartale

### 5. Layout (in RE)
- **Label-Breite:** Breite der Phasennamen-Spalte (Standard: 20 RE)
- **Header-Höhe:** Höhe der Spaltenüberschriften (Standard: 3 RE)
- **Zeilen-Höhe:** Gesamthöhe einer Phasenzeile (Standard: 5 RE)
- **Balken-Höhe:** Höhe des farbigen Balkens innerhalb der Zeile (Standard: 3 RE)

### 6. Spaltenbreite
- **Auto-Verteilung:** Gleichmäßige Verteilung auf verfügbare Breite
- **Feste Breite:** Manuelle RE-Eingabe pro Spalte

### 7. Phasen
- Button **+ Phase** zum Hinzufügen
- Je Phase: Name | Start | Ende | Farbe | ×(Löschen)
- Beliebig viele Phasen möglich

### 8. Optionen
- Checkbox: **Heute-Linie anzeigen**

### 9. Erstellen
- Button **GANTT erstellen** – erzeugt das Diagramm auf der letzten Folie
- Statusanzeige mit Ergebnis-Informationen

---

## Rastereinheit (RE)

Die Rastereinheit ist die fundamentale Maßeinheit des Add-ins. Sie entspricht dem in PowerPoint eingestellten Raster.

### Standard-Rastereinheit
```
RE = 0,2117 cm = 6 pt / 28,3465
```

### Umrechnung
```
1 cm = 28,3464567 Points (exakt)
RE_PT = gridUnitCm × CM_PT
```

### Presets
| Preset | cm | Points | Verwendung |
|---|---|---|---|
| 0,2117 | 0,2117 cm | ~6,0 pt | Standard (fein) |
| 0,4233 | 0,4233 cm | ~12,0 pt | Mittel |
| 0,6350 | 0,6350 cm | ~18,0 pt | Grob |
| 2,1167 | 2,1167 cm | ~60,0 pt | Sehr grob |

### Auswirkung
Die RE bestimmt:
- Alle Positionsberechnungen (`reToX()`, `reToY()`)
- Alle Größenberechnungen (`re2pt()`)
- Das Snapping von Balken und Linien (`snapToGrid()`)

---

## Bildschirmformat-Erkennung

### Automatische Erkennung beim Start
Das Add-in liest beim Laden automatisch die Foliengröße via `pageSetup` (API 1.10) aus und bestimmt das Format über die integrierte **FORMAT_TABLE**.

### 4-Stufen-Erkennung (aus Grid_Resize_Tool v15)

| Stufe | Methode | Toleranz | Ergebnis |
|---|---|---|---|
| 1 | **Exakter Match** | ±6 pt auf Breite und Höhe | `"16:9"` |
| 2 | **Aspect-Ratio** | ±0,5 % auf das Seitenverhältnis | `"16:9 (Ratio)"` |
| 3 | **Nearest-Neighbor** | Distanz < 50 pt | `"16:9 (NN)"` |
| 4 | **Unbekannt** | Kein Match | `"Unbekannt"` (Offset 0/0) |

### Manuelle Erkennung
Button **🔎 Erkennen** löst `detectSlideSize()` aus und aktualisiert Format + Offsets.

---

## Grid-System & Offset-Logik

### Prinzip
Das Grid-System kompensiert den Versatz zwischen dem PowerPoint-Rasterursprung und der Folienecke (0,0). Dieser Versatz ist **format-abhängig**.

### Formeln
```
Position_X(re) = GRID_OFFSET_X + (re × RE_PT)
Position_Y(re) = GRID_OFFSET_Y + (re × RE_PT)
Snap(pos)      = GRID_OFFSET + round((pos - GRID_OFFSET) / RE_PT) × RE_PT
Größe(re)      = re × RE_PT     (reine Distanz, ohne Offset)
```

### Offset-Unterscheidung
- **Positionen** (left, top): Verwenden `reToX()` / `reToY()` → **mit** Offset
- **Größen** (width, height): Verwenden `re2pt()` → **ohne** Offset

---

## Zeitraum & Zeiteinheiten

### Zeiteinheiten

| Einheit | Spalten-Label | Monatszeile | Beispiel |
|---|---|---|---|
| **Tage** | `TT.MM` | Ja (Monat + Jahr) | `01.03`, `02.03` |
| **Wochen** | KW-Nummer | Ja (Monat + Jahr) | `9`, `10`, `11` |
| **Monate** | `Mon JJJJ` | Nein | `Mär 2026`, `Apr 2026` |
| **Quartale** | `Qn JJJJ` | Ja (Quartal + Jahr) | `Q1 2026`, `Q2 2026` |

### Monatszeile
Bei Tagen, Wochen und Quartalen wird über den Spaltenüberschriften eine zusätzliche **Gruppierungs-Zeile** angezeigt, die zusammengehörige Spalten unter Monats-/Quartals-Labels zusammenfasst.

### KW-Berechnung
Die Kalenderwochen werden nach **ISO 8601** berechnet (`getISOWeek()`).

---

## Layout-Parameter

### Standard-Layout (in RE)

| Parameter | Standard | Beschreibung |
|---|---|---|
| `GANTT_LEFT_RE` | 7 | Linker Startpunkt des GANTT |
| `GANTT_TOP_RE` | 17 | Oberer Startpunkt des GANTT |
| `labelWidthRE` | 20 | Breite der Phasennamen-Spalte |
| `headerHeightRE` | 3 | Höhe der Spaltenüberschriften |
| `rowHeightRE` | 5 | Gesamthöhe einer Phasenzeile |
| `barHeightRE` | 3 | Höhe des farbigen Balkens |

### Balken-Padding
Der Balken wird vertikal innerhalb der Zeile zentriert:
```
barPadding = max(2, round((rowHeightPt - barHeightPt) / 2))
```

---

## Spaltenbreite & Verteilung

### Auto-Verteilung (empfohlen)
```
slideWidthRE = floor(SLIDE_WIDTH_PT / RE_PT)
autoMaxRE    = slideWidthRE - GANTT_LEFT_RE - 6   (6 RE rechter Rand)
availableRE  = autoMaxRE - labelWidthRE
colWidthRE   = floor(availableRE / anzahlSpalten)
```

**Ergebnis:** Die Spalten füllen die verfügbare Breite gleichmäßig aus, mit mindestens 6 RE Abstand zum rechten Folienrand.

### Feste Breite
Manuelle Eingabe der Spaltenbreite in RE. Bei Überschreitung der verfügbaren Breite werden Spalten **abgeschnitten** (Truncation).

### Truncation
Wenn die berechnete GANTT-Breite die verfügbare Fläche überschreitet:
- Spalten werden abgeschnitten
- Ein gelber Hinweis-Shape wird unterhalb des GANTT platziert:  
  `⚠ Darstellung abgeschnitten (max. 118 RE Breite)`

---

## Phasen

### Definition
Jede Phase besteht aus:
- **Name:** Freitext (wird in der Label-Spalte angezeigt)
- **Start:** Datum (Anfang des Balkens)
- **Ende:** Datum (Ende des Balkens)
- **Farbe:** Hex-Farbcode (Farbe des Balkens)

### Standard-Phasen (beim Start)
| Phase | Dauer | Farbe |
|---|---|---|
| Konzeption | Heute + 21 Tage | `#2e86c1` (Blau) |
| Umsetzung | Tag 21 – Tag 63 | `#27ae60` (Grün) |
| Abnahme | Tag 63 – 3 Monate | `#e94560` (Rot) |

### Balken-Berechnung
```
startFrac = (clampStart - projStart) / (projEnd - projStart)
endFrac   = (clampEnd   - projStart) / (projEnd - projStart)
barLeft   = startFrac × chartWidth   → auf Raster gesnappt
barWidth  = (endFrac - startFrac) × chartWidth → auf Raster gesnappt (min. 1 RE)
```

Phasen, die außerhalb des Projektzeitraums liegen, werden auf den sichtbaren Bereich **geclampt**.

---

## Heute-Linie

### Funktion
Eine vertikale rote Linie markiert das aktuelle Datum im GANTT-Diagramm.

### Darstellung
- **Typ:** Echte PowerPoint-Linie (`addLine`, `ConnectorType.straight`)
- **Farbe:** `#FF0000` (Rot)
- **Stärke:** 1,5 pt
- **Breite:** 0,01 pt (verhindert Auto-Routing/Schiefwerden)
- **Höhe:** Gesamte GANTT-Höhe + 2 RE nach unten

### Datum-Label
Unterhalb der Linie wird ein rotes Rechteck mit dem Datum (`TT.MM`) in weißer Schrift platziert:
- Breite: 6 RE, Höhe: 2 RE
- Schriftgröße: 8 pt, fett

### Bedingung
Die Linie wird nur angezeigt, wenn:
1. Die Checkbox aktiviert ist **und**
2. Das heutige Datum innerhalb des Projektzeitraums liegt

---

## GANTT-Erstellung – Technischer Ablauf

### Schritt-für-Schritt

```
1. Eingaben validieren (Zeitraum, Phasen)
2. Zeiteinheiten berechnen (computeTimeUnits)
3. Spaltenbreite berechnen (Auto/Feste Breite)
4. PowerPoint.run() starten
   ├── Foliengröße lesen (pageSetup)
   ├── Grid-System aktualisieren (updateGridSystem)
   ├── Letzte Folie referenzieren
   └── drawGantt() aufrufen:
       ├── 1. Hintergrund-Rechteck
       ├── 2a. Monatszeile (wenn nötig)
       ├── 2b. Header-Zellen (Spaltenüberschriften)
       ├── 3. Phasenzeilen:
       │   ├── Label-Zelle (links, grau)
       │   ├── Zeilen-Hintergrund (alternierend weiß/hellgrau)
       │   └── Balken (farbig, auf Raster gesnappt)
       ├── 4. Spalten-Rechtecke (transparent + Rahmen)
       ├── 5. Heute-Linie + Label
       └── 6. Truncation-Hinweis (wenn nötig)
5. ctx.sync() → Status anzeigen
```

### Shape-Hierarchie (Z-Order)
Die Shapes werden in dieser Reihenfolge erstellt (erste = unten):
1. Hintergrund (weiß)
2. Monatszeile (grau)
3. Header-Zellen (hellgrau)
4. Label-Zellen + Zeilen-Hintergründe
5. Balken (farbig)
6. Spalten-Rechtecke (transparent + Rahmen)
7. Heute-Linie + Label

---

## Positionierung & Snapping

### Position vs. Größe

| Funktion | Typ | Formel | Offset |
|---|---|---|---|
| `re2pt(re)` | Größe/Distanz | `re × RE_PT` | Nein |
| `reToX(re)` | X-Position | `GRID_OFFSET_X + re × RE_PT` | Ja |
| `reToY(re)` | Y-Position | `GRID_OFFSET_Y + re × RE_PT` | Ja |
| `snapToGrid(pos, offset)` | Position snappen | `offset + round((pos - offset) / RE_PT) × RE_PT` | Ja |
| `snapSize(size)` | Größe snappen | `round(size / RE_PT) × RE_PT` (min. 1 RE) | Nein |

### Wann wird gesnappt?
- **Balken-Position:** `snapToGrid()` auf die berechnete Pixelposition
- **Balken-Breite:** `round(barWidth / RE_PT) × RE_PT`
- **Heute-Linie:** `snapToGrid()` auf die berechnete Tagesposition
- **Alle anderen Shapes:** Direkt über `reToX()`/`reToY()` (exakt auf Raster)

---

## Visuelle Struktur

### Layout-Diagramm (v21.3)

```
┌─ Folie ──────────────────────────────────────────────────┐
│                                                           │
│  7 RE  ┌─ GANTT ───────────────────────────┐  6 RE       │
│ ←────→ │                                   │ ←────→      │
│        │┌─Monatszeile──────────────────────┐│             │
│        ││  Jan 2026     │   Feb 2026       ││             │
│        │├─Header────────┼──────────────────┤│             │
│        ││  KW5 │ KW6    │  KW7  │  KW8    ││             │
│ 17 RE  │├──────┼────────┼───────┼─────────┤│             │
│ ↕      ││Phase1│████████│       │         ││             │
│        │├──────┼────────┼───────┼─────────┤│             │
│        ││Phase2│        │███████│█████████ ││             │
│        │├──────┼────────┼───────┼─────────┤│             │
│        ││Phase3│        │       │  ███████ ││             │
│        │└──────┴────────┴───────┴─────────┘│             │
│        │         ↑ Transparente Rahmen-     │             │
│        │           Rechtecke je Spalte      │             │
│        └────────────────────────────────────┘             │
└──────────────────────────────────────────────────────────┘
```

### Spalten-Rechtecke (v21.3)
Ab v21.3 werden statt vertikaler Trennlinien **transparente Rahmen-Rechtecke** pro Zeiteinheit verwendet:
- **Füllung:** Vollständig transparent (`fill.transparency = 1`)
- **Rahmen:** Dezent grau (`#C0C0C0`, Stärke: `LINE_WEIGHT`)
- **Breite:** Exakte Spaltenbreite
- **Höhe:** Gesamte Phasen-Höhe
- **Vorteil:** Kein Schiefwerden beim Skalieren, saubere vertikale Struktur

### Farbschema

| Element | Farbe | Hex |
|---|---|---|
| Hintergrund | Weiß | `#FFFFFF` |
| Zeilen (gerade) | Weiß | `#FFFFFF` |
| Zeilen (ungerade) | Hellgrau | `#F8F8F8` |
| Label-Zellen | Hellgrau | `#F0F0F0` |
| Header-Zellen | Grau | `#D5D5D5` |
| Monatszeile | Dunkelgrau | `#B0B0B0` |
| Rahmenlinien | Mittelgrau | `#808080` |
| Spalten-Rahmen | Hellgrau | `#C0C0C0` |
| Heute-Linie | Rot | `#FF0000` |
| Truncation-Hinweis | Gelb | `#FFF3CD` |

### Schrift
- **Größe:** 11 pt (GANTT-Inhalte), 8 pt (Heute-Label)
- **Ausrichtung:** Labels linksbündig, Header/Monate zentriert
- **Textfeld-Ränder:** Links 0,1 cm, sonst 0

---

## Konstanten & Konfiguration

### Globale Konstanten (taskpane.js)

```javascript
var CM_PT = 28.3464567;           // 1 cm = 28.3464567 Points
var gridUnitCm = 0.2117;          // Standard-RE in cm
var RE_PT = gridUnitCm * CM_PT;   // RE in Points (~6.0 pt)

var SLIDE_WIDTH_PT  = 960;        // Default (wird überschrieben)
var SLIDE_HEIGHT_PT = 540;        // Default (wird überschrieben)

var GANTT_LEFT_RE = 7;            // Start links (in RE)
var GANTT_TOP_RE  = 17;           // Start oben (in RE)
var GANTT_MAX_WIDTH_RE = 118;     // Maximale Breite

var FONT_SIZE    = 11;            // Schriftgröße in pt
var LINE_WEIGHT  = 0.5;           // Rahmenstärke in pt

var TEXT_MARGIN_LEFT   = 0.1 * CM_PT;  // ~2.83 pt
var TEXT_MARGIN_RIGHT  = 0;
var TEXT_MARGIN_TOP    = 0;
var TEXT_MARGIN_BOTTOM = 0;
```

---

## Funktionsreferenz (taskpane.js)

### Konvertierung & Grid

| Funktion | Parameter | Rückgabe | Beschreibung |
|---|---|---|---|
| `re2pt(re)` | RE-Anzahl | Points | RE → Points (Distanz, ohne Offset) |
| `reToX(re)` | RE-Anzahl | Points | RE → X-Position (mit Offset) |
| `reToY(re)` | RE-Anzahl | Points | RE → Y-Position (mit Offset) |
| `cm2pt(cm)` | cm-Wert | Points | Zentimeter → Points |
| `c2p(cm)` | cm-Wert | Points | Alias für cm2pt |
| `p2c(pt)` | Points | cm | Points → Zentimeter |
| `snapToGrid(pos, offset)` | Position, Offset | Points | Position auf nächsten Rasterpunkt |
| `snapSize(size)` | Größe in pt | Points | Größe auf RE-Vielfaches (min. 1 RE) |
| `updateGridSystem()` | – | – | Grid-Offsets neu berechnen |

### Format-Erkennung

| Funktion | Parameter | Rückgabe | Beschreibung |
|---|---|---|---|
| `getGridOffsets(w, h)` | Breite, Höhe in pt | `{name, x, y}` | 4-Stufen Format-Erkennung |
| `detectSlideSize()` | – | – | Foliengröße via API lesen + Grid aktualisieren |

### Zeitberechnung

| Funktion | Parameter | Rückgabe | Beschreibung |
|---|---|---|---|
| `computeTimeUnits(start, end, unit)` | Dates, Einheit | Array | Zeiteinheiten mit Label + Datum |
| `computeMonthGroups(units, unit)` | TimeUnits, Einheit | Array | Monatsgruppen für Gruppenzeile |
| `getISOWeek(d)` | Date | Nummer | ISO 8601 Kalenderwoche |
| `getQuarter(d)` | Date | 1–4 | Quartal des Datums |

### Phasen-Verwaltung

| Funktion | Parameter | Beschreibung |
|---|---|---|
| `addPhaseRow(name, start, end, color)` | Strings/Dates | Phase zur UI hinzufügen |
| `getPhases()` | – | Alle Phasen aus DOM lesen → Array |

### GANTT-Erstellung

| Funktion | Parameter | Beschreibung |
|---|---|---|
| `createGanttChart()` | – | Hauptfunktion: Validierung + PowerPoint.run |
| `drawGantt(ctx, slide, ...)` | Viele | Zeichnet alle Shapes auf die Folie |

### UI & Hilfe

| Funktion | Beschreibung |
|---|---|
| `initUI()` | Event-Listener registrieren |
| `initDefaults()` | Standard-Zeitraum + 3 Phasen setzen |
| `updateInfoBar()` | Version, API-Status, Uhrzeit anzeigen |
| `updatePresetButtons(v)` | Aktiven Preset visuell markieren |
| `showStatus(msg, type)` | Status-Nachricht anzeigen (`info`/`success`/`error`/`warning`) |
| `formatTextFrame(shape, text, centered)` | Textformat auf Shape anwenden |

### Hilfsfunktionen

| Funktion | Beschreibung |
|---|---|
| `pad2(n)` | Zahl mit führender Null: `5` → `"05"` |
| `toISO(d)` | Date → `"YYYY-MM-DD"` |
| `addDays(d, n)` | Datum + n Tage |
| `escHtml(s)` | HTML-Sonderzeichen escapen |
| `randomColor()` | Zufällige Farbe aus DROEGE-Palette |

---

## Format-Tabelle (Grid_Resize_Tool v15)

Die Offset-Tabelle beschreibt den Versatz des Rasterursprungs zur Folienecke (0,0) je Bildschirmformat:

| Format | Breite (pt) | Höhe (pt) | Offset X (cm) | Offset Y (cm) |
|---|---|---|---|---|
| **16:9** | 720.0 | 405.0 | 0,00 | 0,16 |
| **4:3** | 720.0 | 540.0 | 0,00 | 0,00 |
| **16:10** | 720.0 | 450.0 | 0,00 | 0,00 |
| **A4 quer** | 780.0 | 540.0 | 0,00 | 0,00 |
| **Breitbild** | 960.0 | 540.0 | 0,00 | 0,00 |

> **Hinweis:** Das Format **16:9** hat einen Y-Offset von 0,16 cm. Das bedeutet, dass das Raster bei diesem Format nicht bei y=0 beginnt, sondern um ~4,5 pt nach unten versetzt ist. Alle `reToY()`-Berechnungen berücksichtigen dies automatisch.

---

## Versionshistorie

### v21.3 (aktuell)
- **Spalten-Rechtecke statt Trennlinien:** Vertikale Trennlinien (`addLine`) durch transparente Rahmen-Rechtecke (`addGeometricShape`) ersetzt. Füllung transparent, Rahmen dezent grau. Kein Schiefwerden beim Skalieren.
- **Heute-Linie Fix:** `width: 0` → `width: 0.01` verhindert Auto-Routing/Verschieben
- **Datumsfelder schmaler:** Phasen-Datumsfelder von 105px auf 86px (~0,5 cm schmaler)

### v21.2
- **Dynamische Positionierung:** Start bei RE 7 (statt 9), rechter Rand 6 RE, automatische Berechnung der verfügbaren Breite für alle Folienformate
- **Echte Linien:** Heute-Linie als echte PowerPoint-Linie (`addLine`)
- **Phasendefinition einzeilig:** Kompaktere Darstellung der Phasenfelder

### v21.1
- CSS-Optimierungen, Phasen-Styling

### v21.0
- **FORMAT_TABLE** aus Grid_Resize_Tool v15 integriert
- **4-Stufen-Erkennung** (exakt/ratio/NN/unbekannt) für Folienformate
- **Dynamisches Auslesen** der Foliengröße via `pageSetup` (API 1.10)
- **reToX()/reToY()** verwenden format-abhängige Offsets
- **updateGridSystem()** als zentrale Grid-Funktion
- Bildschirmformat-Anzeige mit Erkennen-Button

### v20.0
- Erste Version mit Office Web Add-in Architektur
- Grundlegende GANTT-Erstellung mit Rastereinheiten

### v2.14 (Legacy)
- Ursprüngliche Version mit fester 118 RE Maximalbreite
- Feste Positionierung (Left=9, Top=17 RE)
- Manueller Spaltenmodus

---

## Bekannte Einschränkungen

### Office.js API
- **Grid-Einstellungen nicht auslesbar:** Die in PowerPoint eingestellte Rastergröße und Snap-Optionen können über Office Web Add-ins nicht direkt gelesen werden. Daher muss die RE manuell im Add-in eingegeben werden.
- **View-Einstellungen nicht auslesbar:** Zoom, Fensterposition und Display-Informationen sind über Office.js nicht verfügbar.

### Workarounds
- RE wird über das Eingabefeld oder Preset-Buttons manuell gesetzt
- Folienformat wird über `pageSetup` erkannt (API 1.10)
- Für vollautomatische Grid-Erkennung wäre ein COM/VSTO/VBA-Ansatz nötig (nur Desktop Windows)

### Darstellung
- Bei sehr vielen Zeiteinheiten (z.B. 365 Tage) kann das GANTT abgeschnitten werden (Truncation-Warnung)
- Schriftgrößen werden nicht automatisch an die Spaltenbreite angepasst

---

## Lizenz

DROEGE GROUP · 2026  
Internes Tool – Vertraulich

---

*Generiert am 27.02.2026 | GANTT Generator v21.3*
