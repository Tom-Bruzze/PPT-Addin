/*
 ═══════════════════════════════════════════════════════
 Droege GANTT Generator  –  taskpane.js  v21.3

 ÄNDERUNGEN v21.0 (basierend auf v20.0):
  - FORMAT_TABLE aus Grid_Resize_Tool v15 integriert
  - getGridOffsets() mit 4-Stufen-Erkennung (exakt/ratio/NN/unbekannt)
  - Dynamisches Auslesen der Foliengröße via pageSetup (API 1.10)
  - reToX() / reToY() verwenden jetzt Format-abhängige Offsets
  - Neues UI: Bildschirmformat-Anzeige mit "Erkennen"-Button
  - updateGridSystem() als zentrale Funktion für alle Rasterberechnungen
  - Alle Shape-Positionen nutzen offsetX/offsetY + RE*gPt
  - "Format setzen" Button entfernt (wie in v20)
  - API 1.10 Check für pageSetup

 DROEGE GROUP · 2026
 ═══════════════════════════════════════════════════════
*/

var VERSION = "21.0";

// ═══════════════════════════════════════════════════════
// KONSTANTEN
// ═══════════════════════════════════════════════════════
var CM_PT = 28.3464567;           // 1 cm = 28.3464567 Points (exakt)
var gridUnitCm = 0.2117;          // Standard-Rastereinheit (6 pt / CM_PT)
var RE_PT = gridUnitCm * CM_PT;   // Rastereinheit in Points
var apiOk = false;                // PowerPointApi 1.10 verfügbar?

var ganttPhaseCount = 0;

// ═══════════════════════════════════════════════════════
// FORMAT-ERKENNUNG & GRID-OFFSETS  (Grid_Resize_Tool v15)
//
// Offsets in cm – geprüft und vom Benutzer bestätigt.
// Die Offsets beschreiben die Verschiebung des Rasterursprungs
// relativ zur Folienecke links oben.
// ═══════════════════════════════════════════════════════
var FORMAT_TABLE = [
    { name: "16:9",      w: 720.0, h: 405.0, offXcm: 0.00, offYcm: 0.16 },
    { name: "4:3",       w: 720.0, h: 540.0, offXcm: 0.00, offYcm: 0.00 },
    { name: "16:10",     w: 720.0, h: 450.0, offXcm: 0.00, offYcm: 0.00 },
    { name: "A4 quer",   w: 780.0, h: 540.0, offXcm: 0.00, offYcm: 0.00 },
    { name: "Breitbild", w: 960.0, h: 540.0, offXcm: 0.00, offYcm: 0.00 }
];

// ═══════════════════════════════════════════════════════
// GRID-OFFSET-ERKENNUNG (4-Stufen, aus Grid_Resize_Tool v15)
// ═══════════════════════════════════════════════════════
function c2p(cm) { return cm * CM_PT; }
function p2c(pt) { return pt / CM_PT; }

function getGridOffsets(slideWidthPt, slideHeightPt) {
    /* 1) Exakter Match mit Toleranz ±6 pt */
    var tol = 6;
    for (var i = 0; i < FORMAT_TABLE.length; i++) {
        var f = FORMAT_TABLE[i];
        if (Math.abs(slideWidthPt - f.w) <= tol && Math.abs(slideHeightPt - f.h) <= tol) {
            return { name: f.name, x: c2p(f.offXcm), y: c2p(f.offYcm) };
        }
    }

    /* 2) Aspect-Ratio-Match (±0.5 %) */
    var ratio = slideWidthPt / slideHeightPt;
    var bestRatio = null, bestRatioDiff = Infinity;
    for (var i = 0; i < FORMAT_TABLE.length; i++) {
        var f = FORMAT_TABLE[i];
        var fRatio = f.w / f.h;
        var diff = Math.abs(ratio - fRatio) / fRatio;
        if (diff < 0.005 && diff < bestRatioDiff) {
            bestRatioDiff = diff;
            bestRatio = f;
        }
    }
    if (bestRatio) {
        return { name: bestRatio.name + " (Ratio)", x: c2p(bestRatio.offXcm), y: c2p(bestRatio.offYcm) };
    }

    /* 3) Nearest-Neighbor Fallback (Distanz < 50 pt) */
    var bestNN = null, bestNNdist = Infinity;
    for (var i = 0; i < FORMAT_TABLE.length; i++) {
        var f = FORMAT_TABLE[i];
        var d = Math.abs(slideWidthPt - f.w) + Math.abs(slideHeightPt - f.h);
        if (d < bestNNdist) { bestNNdist = d; bestNN = f; }
    }
    if (bestNN && bestNNdist < 50) {
        return { name: bestNN.name + " (NN)", x: c2p(bestNN.offXcm), y: c2p(bestNN.offYcm) };
    }

    /* 4) Unbekannt – kein Offset */
    return { name: "Unbekannt", x: 0, y: 0 };
}

// ═══════════════════════════════════════════════════════
// GRID-SYSTEM: Dynamisch (Foliengröße + Format-Offsets)
// ═══════════════════════════════════════════════════════
//
// Statt fester 960×540 pt wird die tatsächliche
// Foliengröße per pageSetup ausgelesen.
// Die Grid-Offsets kommen aus der FORMAT_TABLE.
//
// Formel (Grid_Resize_Tool v15):
//   gPt           = gridUnitCm * CM_PT      (Raster in pt)
//   formatOffset  = getGridOffsets(w, h)     (Format-Offset)
//   GRID_MARGIN_X = formatOffset.x           (X-Offset)
//   GRID_MARGIN_Y = formatOffset.y           (Y-Offset)
//   Position(re)  = GRID_MARGIN + re * gPt
//   Snap(pos)     = GRID_MARGIN + round((pos - GRID_MARGIN) / gPt) * gPt
// ═══════════════════════════════════════════════════════

// Aktuelle Foliengröße (Defaults, wird beim Start überschrieben)
var SLIDE_WIDTH_PT  = 960;
var SLIDE_HEIGHT_PT = 540;

// Grid-Offsets (werden dynamisch berechnet)
var GRID_OFFSET_X = 0;
var GRID_OFFSET_Y = 0;
var DETECTED_FORMAT = "Breitbild";

// GANTT Layout (in RE, relativ zum Grid-Ursprung)
var GANTT_LEFT_RE = 7;
var GANTT_TOP_RE = 17;
var GANTT_MAX_WIDTH_RE = 118;

// Schriftgröße
var FONT_SIZE = 11;

// Linienstärke für Rechteck-Objekte
var LINE_WEIGHT = 0.5;

// Textfeld-Ränder (in Points)
var TEXT_MARGIN_LEFT   = 0.1 * CM_PT;  // 0.1 cm ≈ 2.83 pt
var TEXT_MARGIN_RIGHT  = 0;
var TEXT_MARGIN_TOP    = 0;
var TEXT_MARGIN_BOTTOM = 0;

// ═══════════════════════════════════════════════════════
// KONVERTIERUNGS-FUNKTIONEN (Grid_Resize_Tool v15 Logik)
// ═══════════════════════════════════════════════════════

/** RE → Points (reine Distanz, ohne Offset) */
function re2pt(re) {
  return re * RE_PT;
}

/** RE → X-Position auf der Folie (mit Format-Offset) */
function reToX(re) {
  return GRID_OFFSET_X + (re * RE_PT);
}

/** RE → Y-Position auf der Folie (mit Format-Offset) */
function reToY(re) {
  return GRID_OFFSET_Y + (re * RE_PT);
}

/** cm → Points */
function cm2pt(cm) {
  return cm * CM_PT;
}

/** Position auf nächsten Rasterpunkt snappen (wie Grid_Resize_Tool v15) */
function snapToGrid(pos, offset) {
  return offset + Math.round((pos - offset) / RE_PT) * RE_PT;
}

/** Größe auf nächstes Raster-Vielfaches snappen */
function snapSize(size) {
  var snapped = Math.round(size / RE_PT) * RE_PT;
  return snapped < RE_PT ? RE_PT : snapped;
}

/**
 * Grid-System aktualisieren.
 * Wird aufgerufen wenn sich RE, Foliengröße oder Format ändert.
 */
function updateGridSystem() {
  RE_PT = gridUnitCm * CM_PT;
  var off = getGridOffsets(SLIDE_WIDTH_PT, SLIDE_HEIGHT_PT);
  GRID_OFFSET_X = off.x;
  GRID_OFFSET_Y = off.y;
  DETECTED_FORMAT = off.name;

  // UI aktualisieren
  var fmtEl = document.getElementById("detectedFormat");
  if (fmtEl) {
    fmtEl.textContent = DETECTED_FORMAT + " (" +
      p2c(SLIDE_WIDTH_PT).toFixed(2) + " × " +
      p2c(SLIDE_HEIGHT_PT).toFixed(2) + " cm)";
  }
  var offEl = document.getElementById("gridOffsetInfo");
  if (offEl) {
    offEl.textContent = "Offset X: " + p2c(GRID_OFFSET_X).toFixed(3) +
      " cm, Y: " + p2c(GRID_OFFSET_Y).toFixed(3) + " cm";
  }

  console.log("Grid-System aktualisiert:",
    "Format:", DETECTED_FORMAT,
    "| Slide:", SLIDE_WIDTH_PT.toFixed(1) + "×" + SLIDE_HEIGHT_PT.toFixed(1) + " pt",
    "| RE:", gridUnitCm.toFixed(4), "cm =", RE_PT.toFixed(7), "pt",
    "| Offset X:", GRID_OFFSET_X.toFixed(4), "Y:", GRID_OFFSET_Y.toFixed(4));
}

/**
 * Foliengröße per API 1.10 auslesen und Grid-System aktualisieren.
 */
function detectSlideSize() {
  if (!apiOk) {
    showStatus("PowerPointApi 1.10 nicht verfügbar – verwende Defaults", "warning");
    updateGridSystem();
    return;
  }
  PowerPoint.run(function(ctx) {
    var ps = ctx.presentation.pageSetup;
    ps.load(["slideWidth", "slideHeight"]);
    return ctx.sync().then(function() {
      SLIDE_WIDTH_PT = ps.slideWidth;
      SLIDE_HEIGHT_PT = ps.slideHeight;
      updateGridSystem();
      showStatus("Format erkannt: " + DETECTED_FORMAT + " (" +
        p2c(SLIDE_WIDTH_PT).toFixed(2) + " × " +
        p2c(SLIDE_HEIGHT_PT).toFixed(2) + " cm)", "success");
    });
  }).catch(function(e) {
    console.log("detectSlideSize error:", e);
    showStatus("Format-Erkennung fehlgeschlagen: " + e.message, "error");
  });
}

// ═══════════════════════════════════════════════════════
// OFFICE READY
// ═══════════════════════════════════════════════════════
Office.onReady(function(info) {
  if (info.host === Office.HostType.PowerPoint) {
    // API 1.10 Check (wie Grid_Resize_Tool v15)
    if (Office.context.requirements && Office.context.requirements.isSetSupported) {
      apiOk = Office.context.requirements.isSetSupported("PowerPointApi", "1.10");
    } else {
      apiOk = (typeof PowerPoint !== "undefined" && PowerPoint.run && typeof PowerPoint.run === "function");
    }

    initUI();
    updateInfoBar();

    // Foliengröße beim Start automatisch erkennen
    detectSlideSize();

    if (!apiOk) {
      showStatus("PowerPointApi 1.10 nicht verfügbar", "warning");
    } else {
      showStatus("Bereit", "success");
    }
    console.log("GANTT v21.0 geladen | API 1.10:", apiOk);
  }
});

function updateInfoBar() {
  var now = new Date();
  var d = pad2(now.getDate()) + "." + pad2(now.getMonth() + 1) + "." + now.getFullYear();
  var t = pad2(now.getHours()) + ":" + pad2(now.getMinutes());
  var el = document.getElementById("infoDateTime");
  if (el) el.textContent = d + " " + t;
  var elV = document.getElementById("infoVersion");
  if (elV) elV.textContent = "v" + VERSION;
  var elApi = document.getElementById("infoApi");
  if (elApi) elApi.textContent = apiOk ? "API 1.10 ✓" : "API 1.10 ✗";
}

function initUI() {
  // Rastereinheit-Eingabe
  var gi = document.getElementById("gridUnit");
  if (gi) {
    gi.addEventListener("change", function() {
      var v = parseFloat(this.value);
      if (!isNaN(v) && v > 0) {
        gridUnitCm = v;
        updateGridSystem();
        updatePresetButtons(v);
        showStatus("RE: " + v.toFixed(4) + " cm = " + RE_PT.toFixed(4) + " pt", "info");
      }
    });
  }
  
  // Preset-Buttons für RE
  document.querySelectorAll(".pre").forEach(function(b) {
    b.addEventListener("click", function() {
      var v = parseFloat(this.dataset.value);
      gridUnitCm = v;
      updateGridSystem();
      if (gi) gi.value = v;
      updatePresetButtons(v);
      showStatus("RE: " + v.toFixed(4) + " cm = " + RE_PT.toFixed(4) + " pt", "info");
    });
  });

  // Format erkennen Button
  var btnDetect = document.getElementById("detectFormat");
  if (btnDetect) btnDetect.addEventListener("click", detectSlideSize);

  // Spaltenbreiten-Modus
  var widthMode = document.getElementById("ganttWidthMode");
  if (widthMode) {
    widthMode.addEventListener("change", function() {
      var colWidthDiv = document.getElementById("colWidthDiv");
      if (this.value === "auto") {
        colWidthDiv.style.display = "none";
      } else {
        colWidthDiv.style.display = "block";
      }
    });
  }

  // GANTT erstellen
  var btnGantt = document.getElementById("createGantt");
  if (btnGantt) btnGantt.addEventListener("click", createGanttChart);
  
  // Phase hinzufügen
  var btnAdd = document.getElementById("ganttAddPhase");
  if (btnAdd) btnAdd.addEventListener("click", function() {
    var start = new Date(document.getElementById("ganttStart").value);
    if (isNaN(start.getTime())) start = new Date();
    addPhaseRow("Phase " + (ganttPhaseCount + 1), start, addDays(start, 14), randomColor());
  });

  initDefaults();
}

function updatePresetButtons(v) {
  document.querySelectorAll(".pre").forEach(function(x) {
    x.classList.toggle("active", Math.abs(parseFloat(x.dataset.value) - v) < 0.001);
  });
}

function initDefaults() {
  var today = new Date();
  var end = new Date(today);
  end.setMonth(end.getMonth() + 3);
  
  document.getElementById("ganttStart").value = toISO(today);
  document.getElementById("ganttEnd").value = toISO(end);
  
  addPhaseRow("Konzeption", today, addDays(today, 21), "#2e86c1");
  addPhaseRow("Umsetzung", addDays(today, 21), addDays(today, 63), "#27ae60");
  addPhaseRow("Abnahme", addDays(today, 63), end, "#e94560");
}

// ═══════════════════════════════════════════════════════
// STATUS & HILFSFUNKTIONEN
// ═══════════════════════════════════════════════════════
function showStatus(msg, type) {
  var el = document.getElementById("status");
  if (el) {
    el.textContent = msg;
    el.className = "status " + (type || "info");
  }
  console.log("[" + (type || "info") + "] " + msg);
}

function pad2(n) { return n < 10 ? "0" + n : "" + n; }
function toISO(d) { return d.getFullYear() + "-" + pad2(d.getMonth()+1) + "-" + pad2(d.getDate()); }
function addDays(d, n) { var r = new Date(d); r.setDate(r.getDate() + n); return r; }
function escHtml(s) { return s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }
function randomColor() {
  var colors = ["#2e86c1","#27ae60","#e94560","#f39c12","#8e44ad","#1abc9c","#d35400","#2c3e50"];
  return colors[Math.floor(Math.random() * colors.length)];
}

// ═══════════════════════════════════════════════════════
// PHASEN-VERWALTUNG
// ═══════════════════════════════════════════════════════
function addPhaseRow(name, start, end, color) {
  ganttPhaseCount++;
  var container = document.getElementById("ganttPhases");
  var div = document.createElement("div");
  div.className = "gantt-phase";
  div.innerHTML = 
    '<input type="text" value="' + escHtml(name) + '" class="phase-name" placeholder="Phasenname">' +
    '<input type="date" value="' + toISO(start) + '" class="phase-start">' +
    '<input type="date" value="' + toISO(end) + '" class="phase-end">' +
    '<input type="color" value="' + color + '" class="phase-color">' +
    '<button class="phase-del" onclick="this.parentNode.remove()">×</button>';
  container.appendChild(div);
}

function getPhases() {
  var arr = [];
  document.querySelectorAll(".gantt-phase").forEach(function(div) {
    var nameEl = div.querySelector(".phase-name");
    var startEl = div.querySelector(".phase-start");
    var endEl = div.querySelector(".phase-end");
    var colorEl = div.querySelector(".phase-color");
    if (nameEl && startEl && endEl && colorEl) {
      var s = new Date(startEl.value);
      var e = new Date(endEl.value);
      if (!isNaN(s.getTime()) && !isNaN(e.getTime())) {
        arr.push({ name: nameEl.value, start: s, end: e, color: colorEl.value });
      }
    }
  });
  return arr;
}

// ═══════════════════════════════════════════════════════
// ZEITEINHEITEN-BERECHNUNG
// ═══════════════════════════════════════════════════════
function getISOWeek(d) {
  var tmp = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  tmp.setDate(tmp.getDate() + 3 - ((tmp.getDay() + 6) % 7));
  var jan4 = new Date(tmp.getFullYear(), 0, 4);
  return 1 + Math.round(((tmp - jan4) / 86400000 - 3 + ((jan4.getDay() + 6) % 7)) / 7);
}

function getQuarter(d) { return Math.floor(d.getMonth() / 3) + 1; }

function computeTimeUnits(projStart, projEnd, unit) {
  var units = [];
  var cur = new Date(projStart);
  
  if (unit === "day") {
    while (cur <= projEnd) {
      units.push({ date: new Date(cur), label: pad2(cur.getDate()) + "." + pad2(cur.getMonth()+1) });
      cur.setDate(cur.getDate() + 1);
    }
  } else if (unit === "week") {
    cur.setDate(cur.getDate() - ((cur.getDay() + 6) % 7));
    while (cur <= projEnd) {
      var w = getISOWeek(cur);
      units.push({ date: new Date(cur), label: "" + w });
      cur.setDate(cur.getDate() + 7);
    }
  } else if (unit === "month") {
    cur = new Date(cur.getFullYear(), cur.getMonth(), 1);
    var months = ["Jan","Feb","Mär","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez"];
    while (cur <= projEnd) {
      units.push({ date: new Date(cur), label: months[cur.getMonth()] + " " + cur.getFullYear() });
      cur.setMonth(cur.getMonth() + 1);
    }
  } else if (unit === "quarter") {
    cur = new Date(cur.getFullYear(), Math.floor(cur.getMonth()/3)*3, 1);
    while (cur <= projEnd) {
      var q = getQuarter(cur);
      units.push({ date: new Date(cur), label: "Q" + q + " " + cur.getFullYear() });
      cur.setMonth(cur.getMonth() + 3);
    }
  }
  return units;
}

function computeMonthGroups(timeUnits, unit) {
  var months = ["Jan","Feb","Mär","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez"];
  var groups = [];
  var current = null;
  for (var i = 0; i < timeUnits.length; i++) {
    var d = timeUnits[i].date;
    var key = d.getFullYear() + "-" + d.getMonth();
    var label;
    if (unit === "quarter") {
      key = d.getFullYear() + "-Q" + getQuarter(d);
      label = "Q" + getQuarter(d) + " " + d.getFullYear();
    } else {
      label = months[d.getMonth()] + " " + d.getFullYear();
    }
    if (current && current.key === key) {
      current.count++;
    } else {
      current = { key: key, label: label, count: 1 };
      groups.push(current);
    }
  }
  return groups;
}

// ═══════════════════════════════════════════════════════
// TEXTFORMAT-HILFSFUNKTION
// ═══════════════════════════════════════════════════════
function formatTextFrame(shape, text, centered) {
  try {
    shape.textFrame.textRange.text = text;
    shape.textFrame.textRange.font.size = FONT_SIZE;
    shape.textFrame.textRange.font.bold = false;
    shape.textFrame.textRange.font.color = "000000";
    shape.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
    if (centered) {
      shape.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
    }
    shape.textFrame.marginLeft   = TEXT_MARGIN_LEFT;
    shape.textFrame.marginRight  = TEXT_MARGIN_RIGHT;
    shape.textFrame.marginTop    = TEXT_MARGIN_TOP;
    shape.textFrame.marginBottom = TEXT_MARGIN_BOTTOM;
  } catch(e) {
    console.log("formatTextFrame error:", e);
  }
}

// ═══════════════════════════════════════════════════════
// GANTT ERSTELLEN (Hauptfunktion)
// ═══════════════════════════════════════════════════════
function createGanttChart() {
  if (!apiOk) {
    showStatus("PowerPointApi 1.10 nicht verfügbar!", "error");
    return;
  }

  var projStart = new Date(document.getElementById("ganttStart").value);
  var projEnd   = new Date(document.getElementById("ganttEnd").value);
  var unit      = document.getElementById("ganttUnit").value;
  var widthMode = document.getElementById("ganttWidthMode").value;
  var labelWidthRE  = parseInt(document.getElementById("ganttLabelW").value) || 20;
  var headerHeightRE = parseInt(document.getElementById("ganttHeaderH").value) || 3;
  var rowHeightRE   = parseInt(document.getElementById("ganttRowH").value) || 5;
  var barHeightRE   = parseInt(document.getElementById("ganttBarH").value) || 3;
  var showTodayLine = document.getElementById("ganttTodayLine").checked;

  if (isNaN(projStart.getTime()) || isNaN(projEnd.getTime()) || projEnd <= projStart) {
    showStatus("Ungültiger Zeitraum!", "error");
    return;
  }

  var phases = getPhases();
  if (phases.length === 0) {
    showStatus("Keine Phasen definiert!", "error");
    return;
  }

  var timeUnits = computeTimeUnits(projStart, projEnd, unit);
  if (timeUnits.length === 0) {
    showStatus("Zeitraum zu kurz für Einheit!", "error");
    return;
  }

  var colWidthRE;
  var totalDays = Math.ceil((projEnd - projStart) / 86400000);
  var visibleColumns = timeUnits.length;
  var truncated = false;

  if (widthMode === "auto") {
    // Auto: Start bei RE 7, rechts 6 RE Abstand zum Rand
    var slideWidthRE = Math.floor(SLIDE_WIDTH_PT / RE_PT);
    var autoMaxRE = slideWidthRE - GANTT_LEFT_RE - 6;  // 6 RE rechter Rand
    var availableRE = autoMaxRE - labelWidthRE;
    if (availableRE < 1) availableRE = 1;
    if (availableRE < timeUnits.length) {
      visibleColumns = availableRE;
      truncated = true;
      colWidthRE = 1;
    } else {
      colWidthRE = Math.floor(availableRE / timeUnits.length);
      if (colWidthRE < 1) colWidthRE = 1;
    }
  } else {
    colWidthRE = parseInt(document.getElementById("ganttColW").value) || 3;
    var slideWidthRE_m = Math.floor(SLIDE_WIDTH_PT / RE_PT);
    var manualMaxRE = slideWidthRE_m - GANTT_LEFT_RE - 6;
    var neededRE = labelWidthRE + (timeUnits.length * colWidthRE);
    if (neededRE > manualMaxRE) {
      visibleColumns = Math.floor((manualMaxRE - labelWidthRE) / colWidthRE);
      truncated = true;
    }
  }

  showStatus("Erstelle GANTT... Format: " + DETECTED_FORMAT, "info");

  PowerPoint.run(function(ctx) {
    // Foliengröße nochmal dynamisch lesen
    var ps = ctx.presentation.pageSetup;
    ps.load(["slideWidth", "slideHeight"]);

    var slides = ctx.presentation.slides;
    slides.load("items");
    return ctx.sync().then(function() {
      // Grid-System mit aktueller Foliengröße aktualisieren
      SLIDE_WIDTH_PT = ps.slideWidth;
      SLIDE_HEIGHT_PT = ps.slideHeight;
      updateGridSystem();

      if (slides.items.length === 0) {
        showStatus("Keine Folie vorhanden!", "error");
        return;
      }
      var slide = slides.items[slides.items.length - 1];
      
      drawGantt(ctx, slide, projStart, projEnd, unit, phases, timeUnits,
                labelWidthRE, headerHeightRE, barHeightRE, rowHeightRE,
                colWidthRE, totalDays, showTodayLine, visibleColumns, truncated);
      
      return ctx.sync().then(function() {
        var msg = "GANTT erstellt ✓ | " + DETECTED_FORMAT +
          " | " + visibleColumns + " Spalten" +
          " | Offset X:" + p2c(GRID_OFFSET_X).toFixed(3) +
          " Y:" + p2c(GRID_OFFSET_Y).toFixed(3) + " cm";
        if (truncated) msg += " | ⚠ abgeschnitten";
        showStatus(msg, "success");
      });
    });
  }).catch(function(e) {
    showStatus("Fehler: " + e.message, "error");
    console.error("createGanttChart error:", e);
  });
}

// ═══════════════════════════════════════════════════════
// DRAW GANTT (mit Grid_Resize_Tool v15 Positionierung)
//
// Alle Positionen verwenden reToX()/reToY() die den
// Format-abhängigen Offset berücksichtigen.
// Alle Größen sind exakte RE-Vielfache.
// ═══════════════════════════════════════════════════════
function drawGantt(ctx, slide, projStart, projEnd, unit, phases, timeUnits, 
                   labelWidthRE, headerHeightRE, barHeightRE, rowHeightRE, 
                   colWidthRE, totalDays, showTodayLine, visibleColumns, truncated) {
  
  // Positionen über reToX/reToY (mit Format-Offset)
  var GANTT_LEFT_PT = reToX(GANTT_LEFT_RE);
  var GANTT_TOP_PT  = reToY(GANTT_TOP_RE);
  
  // Dimensionen in Points (reine Distanzen via re2pt)
  var labelWidthPt    = re2pt(labelWidthRE);
  var headerHeightPt  = re2pt(headerHeightRE);
  var barHeightPt     = re2pt(barHeightRE);
  var rowHeightPt     = re2pt(rowHeightRE);
  var colWidthPt      = re2pt(colWidthRE);
  
  var barPadding = Math.max(2, Math.round((rowHeightPt - barHeightPt) / 2));
  
  var chartLeft  = GANTT_LEFT_PT + labelWidthPt;
  var chartWidth = visibleColumns * colWidthPt;
  var totalWidth = labelWidthPt + chartWidth;
  
  var needsMonthRow = (unit === "day" || unit === "week" || unit === "quarter");
  var monthRowHeightPt = needsMonthRow ? headerHeightPt : 0;
  
  var totalHeaderHeight = monthRowHeightPt + headerHeightPt;
  var chartTop    = GANTT_TOP_PT + totalHeaderHeight;
  var totalHeight = totalHeaderHeight + (phases.length * rowHeightPt);
  var lineHeight  = phases.length * rowHeightPt;

  // 1. Hintergrund
  var bg = slide.shapes.addGeometricShape(
    PowerPoint.GeometricShapeType.rectangle,
    { left: GANTT_LEFT_PT, top: GANTT_TOP_PT, width: totalWidth, height: totalHeight }
  );
  bg.fill.setSolidColor("FFFFFF");
  bg.lineFormat.color = "808080";
  bg.lineFormat.weight = LINE_WEIGHT;
  
  // 2a. Monatszeile
  if (needsMonthRow) {
    var monthGroups = computeMonthGroups(timeUnits, unit);
    var monthX = 0;
    for (var m = 0; m < monthGroups.length; m++) {
      var mg = monthGroups[m];
      var monthWidth = mg.count * colWidthPt;
      
      var monthCell = slide.shapes.addGeometricShape(
        PowerPoint.GeometricShapeType.rectangle,
        { left: chartLeft + monthX, top: GANTT_TOP_PT, width: monthWidth, height: monthRowHeightPt }
      );
      monthCell.fill.setSolidColor("B0B0B0");
      monthCell.lineFormat.color = "808080";
      monthCell.lineFormat.weight = LINE_WEIGHT;
      
      formatTextFrame(monthCell, mg.label, true);
      
      monthX += monthWidth;
    }
  }
  
  // 2b. Header-Zellen
  var linePositions = [];
  var headerTop = GANTT_TOP_PT + monthRowHeightPt;
  
  for (var c = 0; c < visibleColumns; c++) {
    var colX = c * colWidthPt;
    linePositions.push(chartLeft + colX);
    
    var hdr = slide.shapes.addGeometricShape(
      PowerPoint.GeometricShapeType.rectangle,
      { left: chartLeft + colX, top: headerTop, width: colWidthPt, height: headerHeightPt }
    );
    hdr.fill.setSolidColor("D5D5D5");
    hdr.lineFormat.color = "808080";
    hdr.lineFormat.weight = LINE_WEIGHT;
    
    formatTextFrame(hdr, timeUnits[c].label, true);
  }
  
  // 3. Phasenzeilen
  for (var p = 0; p < phases.length; p++) {
    var phase = phases[p];
    var rowTop = chartTop + (p * rowHeightPt);
    
    // Label-Zelle
    var lbl = slide.shapes.addGeometricShape(
      PowerPoint.GeometricShapeType.rectangle,
      { left: GANTT_LEFT_PT, top: rowTop, width: labelWidthPt, height: rowHeightPt }
    );
    lbl.fill.setSolidColor("F0F0F0");
    lbl.lineFormat.color = "808080";
    lbl.lineFormat.weight = LINE_WEIGHT;
    formatTextFrame(lbl, phase.name, false);
    
    // Zeilen-Hintergrund (Chart-Bereich)
    var rowBg = slide.shapes.addGeometricShape(
      PowerPoint.GeometricShapeType.rectangle,
      { left: chartLeft, top: rowTop, width: chartWidth, height: rowHeightPt }
    );
    rowBg.fill.setSolidColor(p % 2 === 0 ? "FFFFFF" : "F8F8F8");
    rowBg.lineFormat.color = "808080";
    rowBg.lineFormat.weight = LINE_WEIGHT;
    
    // Balken berechnen
    var phaseStartMs = phase.start.getTime();
    var phaseEndMs   = phase.end.getTime();
    var projStartMs  = projStart.getTime();
    var projEndMs    = projEnd.getTime();
    
    if (phaseEndMs > projStartMs && phaseStartMs < projEndMs) {
      var clampStart = Math.max(phaseStartMs, projStartMs);
      var clampEnd   = Math.min(phaseEndMs, projEndMs);
      
      var startFrac = (clampStart - projStartMs) / (projEndMs - projStartMs);
      var endFrac   = (clampEnd - projStartMs) / (projEndMs - projStartMs);
      
      var barLeftPx = startFrac * chartWidth;
      var barWidthPx = (endFrac - startFrac) * chartWidth;
      
      // Balken auf Raster snappen
      var snappedBarLeft = Math.round(barLeftPx / RE_PT) * RE_PT;
      var snappedBarWidth = Math.max(RE_PT, Math.round(barWidthPx / RE_PT) * RE_PT);
      
      if (snappedBarWidth > 0) {
        var bar = slide.shapes.addGeometricShape(
          PowerPoint.GeometricShapeType.rectangle,
          { left: chartLeft + snappedBarLeft, top: rowTop + barPadding, 
            width: snappedBarWidth, height: barHeightPt }
        );
        bar.fill.setSolidColor(phase.color.replace("#",""));
        bar.lineFormat.visible = false;
      }
    }
  }
  
  // 4. Vertikale Trennlinien (echte Linien, keine Rechtecke)
  for (var v = 1; v < visibleColumns; v++) {
    var lineX = chartLeft + (v * colWidthPt);
    var vLine = slide.shapes.addLine(
      PowerPoint.ConnectorType.straight,
      { left: lineX, top: chartTop, width: 0.01, height: lineHeight }
    );
    vLine.lineFormat.color = "C0C0C0";
    vLine.lineFormat.weight = 0.5;
    vLine.lineFormat.dashStyle = PowerPoint.ShapeLineDashStyle.solid;
  }
  
  // 5. Heute-Linie
  if (showTodayLine) {
    var today = new Date();
    var todayMs = today.getTime();
    var projStartMs2 = projStart.getTime();
    var projEndMs2 = projEnd.getTime();
    
    if (todayMs >= projStartMs2 && todayMs <= projEndMs2) {
      var todayFrac = (todayMs - projStartMs2) / (projEndMs2 - projStartMs2);
      var todayX = chartLeft + (todayFrac * chartWidth);
      
      // Heute-Linie auf Raster snappen
      todayX = snapToGrid(todayX, GRID_OFFSET_X);
      
      var todayLine = slide.shapes.addLine(
        PowerPoint.ConnectorType.straight,
        { left: todayX, top: GANTT_TOP_PT, width: 0.01, height: totalHeight + re2pt(2) }
      );
      todayLine.lineFormat.color = "FF0000";
      todayLine.lineFormat.weight = 1.5;
      todayLine.lineFormat.dashStyle = PowerPoint.ShapeLineDashStyle.solid;
      
      // Datum-Label
      var todayLabel = slide.shapes.addGeometricShape(
        PowerPoint.GeometricShapeType.rectangle,
        { left: todayX - re2pt(3), top: GANTT_TOP_PT + totalHeight + re2pt(0.5), 
          width: re2pt(6), height: re2pt(2) }
      );
      todayLabel.fill.setSolidColor("FF0000");
      todayLabel.lineFormat.visible = false;
      
      try {
        todayLabel.textFrame.textRange.text = pad2(today.getDate()) + "." + pad2(today.getMonth()+1);
        todayLabel.textFrame.textRange.font.size = 8;
        todayLabel.textFrame.textRange.font.bold = true;
        todayLabel.textFrame.textRange.font.color = "FFFFFF";
        todayLabel.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
        todayLabel.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
        todayLabel.textFrame.marginLeft = 0;
        todayLabel.textFrame.marginRight = 0;
        todayLabel.textFrame.marginTop = 0;
        todayLabel.textFrame.marginBottom = 0;
      } catch(e) { console.log("today label error:", e); }
    }
  }

  // Truncation-Hinweis
  if (truncated) {
    var truncNote = slide.shapes.addGeometricShape(
      PowerPoint.GeometricShapeType.rectangle,
      { left: GANTT_LEFT_PT, top: GANTT_TOP_PT + totalHeight + re2pt(3), 
        width: totalWidth, height: re2pt(2) }
    );
    truncNote.fill.setSolidColor("FFF3CD");
    truncNote.lineFormat.color = "F0AD4E";
    truncNote.lineFormat.weight = 0.5;
    formatTextFrame(truncNote, "⚠ Darstellung abgeschnitten (max. " + GANTT_MAX_WIDTH_RE + " RE Breite)", true);
  }
}
