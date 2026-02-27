# GANTT Generator – DROEGE GROUP v21.3

## Änderungen v21.3 (basierend auf v21.2)

### 1. Spalten-Rechtecke statt Trennlinien
- **Vertikale Trennlinien entfernt** (`addLine` → `addGeometricShape`)
- Stattdessen: **Transparente Rahmen-Rechtecke** je Zeiteinheit (Monat/KW/Tag)
- Breite = Spaltenbreite der gewählten Zeiteinheit
- Höhe = gesamte GANTT-Höhe über alle Phasen
- **Füllung:** transparent (`fill.transparency = 1`)
- **Rahmen:** dezent grau (`#C0C0C0`, Stärke: `LINE_WEIGHT`)
- Vorteil: Kein Schiefwerden, saubere Skalierung beim Breiterziehen

### 2. Heute-Linie Fix
- `width: 0` → `width: 0.01` verhindert Auto-Routing/Schiefwerden

### 3. Datumsfelder schmaler
- Phasen-Datumsfelder (Von/Bis) von `105px` auf `86px` reduziert (~0,5 cm schmaler)

## Änderungen v21.2
### 1. Auto-Verteilung: Dynamische Positionierung
- **Start links:** Rastereinheit 7 (vorher 9)
- **Start oben:** Rastereinheit 17 (unverändert)
- **Breite:** Dynamisch berechnet → Folienbreite in RE - 7 (links) - 6 (rechts)
- Funktioniert für alle Bildschirmformate (Breitbild, Standard, A4 etc.)
- Rechter Rand: immer mindestens 6 RE Platz zum Folienrand

### 2. Echte Linien statt Rechtecke (v21.2)
- Heute-Linie als echte PowerPoint-Linie (`addLine`)
- Farbe Heute-Linie: #FF0000, Stärke: 1.5 pt

### 3. Phasendefinition einzeilig
- Felder Phase | Start | Ende | Farbe | × jetzt in einer Zeile
- Kompaktere Darstellung, kein Umbruch mehr

## Layout-Konzept (v21.3)
```
┌──────┬──Jan──┬──Feb──┬──Mär──┬──Apr──┐
│      │┌─────┐│┌─────┐│┌─────┐│┌─────┐│
│Phase1││     │││█████ │││     │││     ││
│      │└─────┘│└─────┘│└─────┘│└─────┘│
│Phase2││     │││     │││██████│││█████ ││
│      │└─────┘│└─────┘│└─────┘│└─────┘│
└──────┴───────┴───────┴───────┴───────┘
         ↑ Transparente Rahmen-Rechtecke
```

## Dateien
| Datei | Beschreibung |
|-------|-------------|
| `taskpane.html` | Benutzeroberfläche |
| `taskpane.js` | Kernlogik (v21.3) |
| `taskpane.css` | Styling (v21.3) |
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
