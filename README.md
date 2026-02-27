# GANTT Generator – DROEGE GROUP v21.2

## Änderungen v21.2
### 1. Auto-Verteilung: Dynamische Positionierung
- **Start links:** Rastereinheit 7 (vorher 9)
- **Start oben:** Rastereinheit 17 (unverändert)
- **Breite:** Dynamisch berechnet → Folienbreite in RE - 7 (links) - 6 (rechts)
- Funktioniert für alle Bildschirmformate (Breitbild, Standard, A4 etc.)
- Rechter Rand: immer mindestens 6 RE Platz zum Folienrand

### 2. Echte Linien statt Rechtecke
- **Vertikale Trennlinien** werden jetzt als echte PowerPoint-Linien (`addLine`) erzeugt
- **Heute-Linie** ebenfalls als echte Linie
- Vorteil: Linien lassen sich nachträglich leicht in Länge/Position anpassen
- Farbe Trennlinien: #C0C0C0, Stärke: 0.5 pt
- Farbe Heute-Linie: #FF0000, Stärke: 1.5 pt

### 3. Phasendefinition einzeilig
- Felder Phase | Start | Ende | Farbe | × jetzt in einer Zeile
- Kompaktere Darstellung, kein Umbruch mehr

## Dateien
| Datei | Beschreibung |
|-------|-------------|
| `taskpane.html` | Benutzeroberfläche |
| `taskpane.js` | Kernlogik (v21.2) |
| `taskpane.css` | Styling (v21.2) |
| `manifest-PPT-Addin.xml` | Office Add-in Manifest |
| `README.md` | Diese Datei |

## Installation
1. Alle Dateien auf einen Webserver oder localhost legen
2. In PowerPoint: Einfügen → Add-ins → manifest-PPT-Addin.xml laden
3. Add-in öffnet sich im Taskpane

## Layout-Logik (Auto-Verteilung)
```
┌─ Folie ────────────────────────────────────────────────┐
│                                                         │
│  7 RE  ┌─ GANTT ──────────────────────┐  6 RE          │
│ ←────→ │                              │ ←────→         │
│        │  Labels │ Spalten...         │                 │
│ 17 RE  │        │                     │                 │
│ ↕      │        │                     │                 │
│        └────────────────────────────────┘                │
└─────────────────────────────────────────────────────────┘
```
