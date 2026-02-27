# Droege GANTT Generator – v20.0

## Überblick

PowerPoint Add-in zur Erstellung von GANTT-Diagrammen im **DROEGE GROUP** Corporate Design.  
Das Add-in wird als Office Task Pane geladen und erzeugt rasterbasierte GANTT-Charts direkt auf der aktuellen Folie.

**Version:** 20.0  
**Herausgeber:** DROEGE GROUP  
**Stand:** Februar 2026

---

## Dateien

| Datei | Beschreibung |
|---|---|
| `manifest.xml` | Office Add-in Manifest (Version 20.0.0.0) |
| `taskpane.html` | Benutzeroberfläche (Task Pane) |
| `taskpane.js` | Kernlogik – Raster, GANTT-Erstellung, Hilfsfunktionen |
| `taskpane.css` | Styling im DROEGE GROUP Design |
| `README.md` | Diese Dokumentation |

---

## Rastersystem (Grid_Resize_Tool v15 Logik)

Das Rastersystem basiert auf der **zentrierten Grid-Logik** aus dem Grid_Resize_Tool v15.

### Prinzip

Die Folie (960 × 540 pt bei 16:9) wird in ganzzahlige Rastereinheiten (RE) unterteilt.  
Der verbleibende Rest wird **symmetrisch** als Margin links/rechts und oben/unten verteilt,  
sodass das Raster **exakt zentriert** auf der Folie liegt.

### Formeln

```
RE_PT         = gridUnitCm × 28.3464567    (Rastereinheit in Points)
FULL_UNITS_X  = floor(960 / RE_PT)         (Ganzzahlige RE horizontal)
FULL_UNITS_Y  = floor(540 / RE_PT)         (Ganzzahlige RE vertikal)
GRID_MARGIN_L = (960 - FULL_UNITS_X × RE_PT) / 2
GRID_MARGIN_T = (540 - FULL_UNITS_Y × RE_PT) / 2
```

### Konvertierungsfunktionen

| Funktion | Beschreibung |
|---|---|
| `re2pt(re)` | RE → Points (reine Distanz, ohne Offset) |
| `reToX(re)` | RE → X-Position auf der Folie (mit linkem Grid-Margin) |
| `reToY(re)` | RE → Y-Position auf der Folie (mit oberem Grid-Margin) |
| `cm2pt(cm)` | cm → Points |
| `updateGridMargins()` | Neuberechnung der Margins bei RE-Änderung |

### Standard-Rastereinheit

- **0.21 cm** = 5.9527559 pt (Standard)
- Presets: 0.21, 0.42, 0.63, 2.10 cm

---

## GANTT-Diagramm Konfiguration

### Feste Layout-Parameter

| Parameter | Wert | Beschreibung |
|---|---|---|
| `GANTT_LEFT_RE` | 9 | Linke Position in RE |
| `GANTT_TOP_RE` | 17 | Obere Position in RE |
| `GANTT_MAX_WIDTH_RE` | 118 | Maximale Breite in RE (≈ 24.78 cm) |
| `FONT_SIZE` | 11 | Schriftgröße in pt |
| `LINE_WEIGHT` | 0.5 | Linienstärke in pt |

### Einstellbare Parameter (UI)

| Parameter | Standard | Bereich | Beschreibung |
|---|---|---|---|
| Label-Breite | 20 RE | 5–40 | Breite der Phasen-Labels |
| Kopfzeile | 3 RE | 2–10 | Höhe der Header-Zeile |
| Zeilenhöhe | 5 RE | 3–15 | Gesamthöhe einer Phasenzeile |
| Balkenhöhe | 3 RE | 1–10 | Höhe der Phasen-Balken |
| Spaltenbreite | 3 RE | 1–20 | Breite pro Zeitspalte (bei fester Breite) |

### Textfeld-Ränder

| Rand | Wert |
|---|---|
| Links | 0.1 cm (≈ 2.83 pt) |
| Rechts | 0 |
| Oben | 0 |
| Unten | 0 |

---

## Zeiteinheiten

| Einheit | Label-Format | Monatszeile |
|---|---|---|
| Tage | Tagesnummer (z.B. "15") | Ja |
| Kalenderwochen | KW-Nummer (z.B. "12") | Ja |
| Monate | Monats-Kürzel (z.B. "Jan") | Nein |
| Quartale | Quartals-Kürzel (z.B. "Q1") | Ja |

---

## Spaltenbreiten-Modi

1. **Feste Breite:** Manuelle RE-Eingabe pro Spalte
2. **Auto-Verteilung:** Gleichmäßige Verteilung auf verfügbare Breite  
   (ganze RE-Einheiten, min. 1, max. 10)

Bei Überschreitung der maximalen Breite (118 RE) werden Spalten **abgeschnitten** (Truncation).

---

## Features

- ✅ Rasterbasierte Positionierung (Grid_Resize_Tool v15)
- ✅ Zentriertes Grid mit symmetrischen Margins
- ✅ Flexible Rastereinheit (RE) mit Presets
- ✅ Mehrere Phasen mit individuellen Farben
- ✅ "Heute"-Linie (rot) mit Datum-Label
- ✅ Monatszeile für Tage/Wochen/Quartale
- ✅ Auto-Spaltenbreite oder feste Breite
- ✅ Truncation bei Überschreitung der Maximalbreite
- ✅ DROEGE GROUP Corporate Design
- ✅ Office Add-in API 1.10+

---

## Änderungshistorie

### v20.0 (Februar 2026)
- Rasterlogik aus Grid_Resize_Tool v15 integriert
- "Format setzen" Button und Funktion entfernt
- Alle Versionsnummern auf v20.0 aktualisiert
- Code bereinigt und geprüft
- Vollständige README-Dokumentation erstellt

### v2.25
- Textfelder: linker Rand 0.1 cm, alle anderen 0 cm
- Kalenderwochen: nur Nummer (ohne "KW")
- Schrift: nicht fett, überall schwarz
- Heute-Linie: Datum-Label am unteren Ende

### v2.14
- Basis-Version mit vollständiger GANTT-Funktionalität
- "Format setzen" Funktion (in v20.0 entfernt)

---

## Installation

1. Dateien auf einen Webserver deployen (z.B. GitHub Pages)
2. `manifest.xml` in PowerPoint sideloaden:
   - Einfügen → Meine Add-ins → Freigegebener Ordner
   - Oder über das Microsoft 365 Admin Center
3. Add-in über den "GANTT Generator" Button im Home-Tab öffnen

---

## Technische Anforderungen

- Microsoft PowerPoint (Desktop oder Online)
- Office Add-in API: PowerPointApi 1.10+
- Moderner Browser (für Office Online)

---

## Lizenz

© 2026 DROEGE GROUP. Alle Rechte vorbehalten.  
Dieses Add-in ist ausschließlich für den internen Gebrauch der DROEGE GROUP bestimmt.
