# GANTT Generator – DROEGE GROUP v21.1

## Änderungen v21.1 (Bugfix-Release)
- **KRITISCH BEHOBEN:** 13 fehlende CSS-Klassen ergänzt
  - `.btn-primary` (GANTT-Erstellen-Button war unsichtbar!)
  - `.btn-add` (Phase-Hinzufügen-Button)
  - `.status`, `.status.info/success/warning/error` (Statusanzeige)
  - `.param-grid` (Layout-Parameter-Raster)
  - `.unit-row`, `.width-mode-row` (Dropdowns)
  - `.date-field` (Datums-Eingaben)
  - `.check-row` (Checkbox Heute-Linie)
  - `.phase-name`, `.phase-start`, `.phase-end`, `.phase-color`, `.phase-del` (Phasen-Zeilen)
  - `.gantt-phase` (Phasen-Container)
- CSS komplett neu aufgebaut: 0 fehlende Klassen
- Alle Element-IDs zwischen HTML ↔ JS verifiziert ✅
- Event-Listener-Kette vollständig geprüft ✅

## Dateien
| Datei | Beschreibung |
|-------|-------------|
| `taskpane.html` | Benutzeroberfläche |
| `taskpane.js` | Kernlogik (v21) |
| `taskpane.css` | Styling (v21.1 – repariert) |
| `manifest-PPT-Addin.xml` | Office Add-in Manifest |
| `README.md` | Diese Datei |

## Installation
1. Alle Dateien auf einen Webserver oder localhost legen
2. In PowerPoint: Einfügen → Add-ins → manifest-PPT-Addin.xml laden
3. Add-in öffnet sich im Taskpane

## Bedienung
1. **Rastereinheit** wählen (Standard: 0,2117 cm)
2. **Bildschirmformat** wird automatisch erkannt
3. **Zeitraum** festlegen (Start/Ende/Einheit)
4. **Phasen** hinzufügen (+ Phase Button)
5. **GANTT erstellen** klicken → Diagramm wird auf aktueller Folie eingefügt
