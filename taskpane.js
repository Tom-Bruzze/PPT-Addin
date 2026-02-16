/*
 ═══════════════════════════════════════════════════════
 Droege GANTT Generator  –  taskpane.js  v2.25

 ÄNDERUNGEN v2.25 (basierend auf v2.24):
  - Textfelder: linker Rand 0.1 cm, alle anderen 0 cm
  - Kalenderwochen: nur Nummer (ohne "KW")
  - Schrift: nicht fett, überall schwarz
  - Heute-Linie: Datum-Label am unteren Ende

 DROEGE GROUP · 2026
 ═══════════════════════════════════════════════════════
*/

var VERSION = "2.25";

// ═══════════════════════════════════════════════════════
// RASTER DIREKT IN POINTS
// ═══════════════════════════════════════════════════════
var RE_PT = 5.9527559;  // Rastereinheit in Points (exakt 0.21 cm)
var CM_PT = 28.3464567; // Points pro cm (exakt)

var gridUnitCm = 0.21;
var ganttPhaseCount = 0;

// GANTT Layout
var GANTT_LEFT_RE = 9;
var GANTT_TOP_RE = 17;
var GANTT_MAX_WIDTH_RE = 118;

// Schriftgröße
var FONT_SIZE = 11;

// Linienstärke für Rechteck-Objekte
var LINE_WEIGHT = 0.5;

// Textfeld-Ränder (in Points)
var TEXT_MARGIN_LEFT = 0.1 * CM_PT;  // 0.1 cm = ~2.83 pt
var TEXT_MARGIN_RIGHT = 0;
var TEXT_MARGIN_TOP = 0;
var TEXT_MARGIN_BOTTOM = 0;

// Feste Foliengröße (16:9 Standard)
var SLIDE_WIDTH_PT = 960;
var SLIDE_HEIGHT_PT = 540;

// Grid-Margins
var FULL_UNITS_X = Math.floor(SLIDE_WIDTH_PT / RE_PT);
var FULL_UNITS_Y = Math.floor(SLIDE_HEIGHT_PT / RE_PT);
var GRID_MARGIN_LEFT = (SLIDE_WIDTH_PT - (FULL_UNITS_X * RE_PT)) / 2;
var GRID_MARGIN_TOP = (SLIDE_HEIGHT_PT - (FULL_UNITS_Y * RE_PT)) / 2;

// ═══════════════════════════════════════════════════════
// KONVERTIERUNGS-FUNKTIONEN
// ═══════════════════════════════════════════════════════
function re2pt(re) { return re * RE_PT; }
function reToX(re) { return GRID_MARGIN_LEFT + (re * RE_PT); }
function reToY(re) { return GRID_MARGIN_TOP + (re * RE_PT); }
function cm2pt(cm) { return cm * CM_PT; }

// ═══════════════════════════════════════════════════════
// OFFICE READY
// ═══════════════════════════════════════════════════════
Office.onReady(function(info) {
  if (info.host === Office.HostType.PowerPoint) {
    initUI();
    updateInfoBar();
    showStatus("Bereit", "success");
    console.log("GANTT v2.25 - RE_PT:", RE_PT);
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
}

function initUI() {
  var gi = document.getElementById("gridUnit");
  if (gi) {
    gi.addEventListener("change", function() {
      var v = parseFloat(this.value);
      if (!isNaN(v) && v > 0) {
        gridUnitCm = v;
        RE_PT = v * CM_PT;
        updateGridMargins();
        updatePresetButtons(v);
      }
    });
  }
  
  document.querySelectorAll(".pre").forEach(function(b) {
    b.addEventListener("click", function() {
      var v = parseFloat(this.dataset.value);
      gridUnitCm = v;
      RE_PT = v * CM_PT;
      updateGridMargins();
      if (gi) gi.value = v;
      updatePresetButtons(v);
    });
  });

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

  var btnSlide = document.getElementById("setSlide");
  if (btnSlide) btnSlide.addEventListener("click", setSlideSize);
  
  var btnGantt = document.getElementById("createGantt");
  if (btnGantt) btnGantt.addEventListener("click", createGanttChart);
  
  var btnAdd = document.getElementById("ganttAddPhase");
  if (btnAdd) btnAdd.addEventListener("click", function() {
    var start = new Date(document.getElementById("ganttStart").value);
    if (isNaN(start.getTime())) start = new Date();
    addPhaseRow("Phase " + (ganttPhaseCount + 1), start, addDays(start, 14), randomColor());
  });

  initDefaults();
}

function updateGridMargins() {
  FULL_UNITS_X = Math.floor(SLIDE_WIDTH_PT / RE_PT);
  FULL_UNITS_Y = Math.floor(SLIDE_HEIGHT_PT / RE_PT);
  GRID_MARGIN_LEFT = (SLIDE_WIDTH_PT - (FULL_UNITS_X * RE_PT)) / 2;
  GRID_MARGIN_TOP = (SLIDE_HEIGHT_PT - (FULL_UNITS_Y * RE_PT)) / 2;
}

function updatePresetButtons(v) {
  document.querySelectorAll(".pre").forEach(function(x) {
    x.classList.toggle("active", Math.abs(parseFloat(x.dataset.value) - v) < 0.01);
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
        arr.push({ name: nameEl.value || "Phase", start: s, end: e, color: colorEl.value });
      }
    }
  });
  return arr;
}

function setSlideSize() {
  showStatus("Folienformat: 16:9 (960 x 540 pt)", "success");
}

// ═══════════════════════════════════════════════════════
// CREATE GANTT CHART
// ═══════════════════════════════════════════════════════
function createGanttChart() {
  console.log("=== createGanttChart START v2.25 ===");
  
  var projStart = new Date(document.getElementById("ganttStart").value);
  var projEnd = new Date(document.getElementById("ganttEnd").value);
  var unit = document.getElementById("ganttUnit").value;
  var labelWidthRE = parseInt(document.getElementById("ganttLabelW").value) || 20;
  var headerHeightRE = parseInt(document.getElementById("ganttHeaderH").value) || 3;
  var barHeightRE = parseInt(document.getElementById("ganttBarH").value) || 3;
  var rowHeightRE = parseInt(document.getElementById("ganttRowH").value) || 5;
  var showTodayLine = document.getElementById("ganttTodayLine").checked;
  
  var widthMode = document.getElementById("ganttWidthMode").value;
  var colWidthRE = parseInt(document.getElementById("ganttColW").value) || 3;

  if (isNaN(projStart.getTime()) || isNaN(projEnd.getTime())) {
    ganttInfo("Bitte Start- und Enddatum eingeben!", true);
    return;
  }
  if (projEnd <= projStart) {
    ganttInfo("Ende muss nach Start liegen!", true);
    return;
  }

  var phases = getPhases();
  if (phases.length === 0) {
    ganttInfo("Mindestens eine Phase hinzufügen!", true);
    return;
  }

  var timeUnits = computeTimeUnits(projStart, projEnd, unit);
  if (timeUnits.length === 0 || timeUnits.length > 200) {
    ganttInfo("Zeitraum anpassen oder andere Einheit wählen!", true);
    return;
  }

  var totalDays = daysBetween(projStart, projEnd);
  if (totalDays < 1) totalDays = 1;

  var availableWidthRE = GANTT_MAX_WIDTH_RE - labelWidthRE;
  
  if (widthMode === "auto") {
    colWidthRE = Math.floor(availableWidthRE / timeUnits.length);
    if (colWidthRE < 1) colWidthRE = 1;
    if (colWidthRE > 10) colWidthRE = 10;
  }
  
  var maxVisibleColumns = Math.floor(availableWidthRE / colWidthRE);
  var visibleColumns = Math.min(timeUnits.length, maxVisibleColumns);
  var truncated = visibleColumns < timeUnits.length;

  var unitNames = {day:"Tage", week:"Wochen", month:"Monate", quarter:"Quartale"};
  var infoText = "<b>" + phases.length + "</b> Phasen, <b>" + visibleColumns + "</b> " + unitNames[unit];
  if (truncated) {
    infoText += " <span style='color:#e94560'>(von " + timeUnits.length + " – abgeschnitten)</span>";
  }
  infoText += "<br>Spaltenbreite: <b>" + colWidthRE + " RE</b> (" + re2pt(colWidthRE).toFixed(2) + " pt)";
  ganttInfo(infoText, false);

  showStatus("Erstelle GANTT auf aktueller Folie...", "working");

  PowerPoint.run(function(ctx) {
    var selectedSlides = ctx.presentation.getSelectedSlides();
    selectedSlides.load("items");
    
    return ctx.sync().then(function() {
      var slide;
      if (selectedSlides.items.length > 0) {
        slide = selectedSlides.items[0];
        console.log("Verwende ausgewählte Folie");
      } else {
        slide = ctx.presentation.slides.getItemAt(0);
        console.log("Fallback: erste Folie");
      }
      
      drawGantt(ctx, slide, projStart, projEnd, unit, phases, timeUnits, 
                labelWidthRE, headerHeightRE, barHeightRE, rowHeightRE, 
                colWidthRE, totalDays, showTodayLine, visibleColumns, truncated);
      
      return ctx.sync();
    });
  }).then(function() {
    console.log("=== createGanttChart DONE ===");
    showStatus("GANTT erstellt!", "success");
  }).catch(function(err) {
    console.error("Fehler:", err);
    showStatus("Fehler: " + err.message, "error");
  });
}

// ═══════════════════════════════════════════════════════
// TEXTFORMAT-HILFSFUNKTION (v2.25)
// ═══════════════════════════════════════════════════════
function formatTextFrame(shape, text, centered) {
  try {
    shape.textFrame.textRange.text = text;
    shape.textFrame.textRange.font.size = FONT_SIZE;
    shape.textFrame.textRange.font.bold = false;  // Nicht fett
    shape.textFrame.textRange.font.color = "000000";  // Immer schwarz
    shape.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
    if (centered) {
      shape.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
    }
    // Textfeld-Ränder
    shape.textFrame.marginLeft = TEXT_MARGIN_LEFT;
    shape.textFrame.marginRight = TEXT_MARGIN_RIGHT;
    shape.textFrame.marginTop = TEXT_MARGIN_TOP;
    shape.textFrame.marginBottom = TEXT_MARGIN_BOTTOM;
  } catch(e) {
    console.log("formatTextFrame error:", e);
  }
}

// ═══════════════════════════════════════════════════════
// DRAW GANTT
// ═══════════════════════════════════════════════════════
function drawGantt(ctx, slide, projStart, projEnd, unit, phases, timeUnits, 
                   labelWidthRE, headerHeightRE, barHeightRE, rowHeightRE, 
                   colWidthRE, totalDays, showTodayLine, visibleColumns, truncated) {
  
  var GANTT_LEFT_PT = reToX(GANTT_LEFT_RE);
  var GANTT_TOP_PT = reToY(GANTT_TOP_RE);
  
  var labelWidthPt = re2pt(labelWidthRE);
  var headerHeightPt = re2pt(headerHeightRE);
  var barHeightPt = re2pt(barHeightRE);
  var rowHeightPt = re2pt(rowHeightRE);
  var colWidthPt = re2pt(colWidthRE);
  
  var barPadding = Math.max(2, Math.round((rowHeightPt - barHeightPt) / 2));
  
  var chartLeft = GANTT_LEFT_PT + labelWidthPt;
  var chartWidth = visibleColumns * colWidthPt;
  var totalWidth = labelWidthPt + chartWidth;
  
  var needsMonthRow = (unit === "day" || unit === "week" || unit === "quarter");
  var monthRowHeightPt = needsMonthRow ? headerHeightPt : 0;
  
  var totalHeaderHeight = monthRowHeightPt + headerHeightPt;
  var chartTop = GANTT_TOP_PT + totalHeaderHeight;
  var totalHeight = totalHeaderHeight + (phases.length * rowHeightPt);
  var lineHeight = phases.length * rowHeightPt;

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
  
  // 3. Trennlinien (unverändert)
  for (var li = 0; li < linePositions.length; li++) {
    var line = slide.shapes.addLine(PowerPoint.ConnectorType.straight,
      { left: linePositions[li], top: chartTop, width: 0.01, height: lineHeight });
    line.lineFormat.color = "AAAAAA";
    line.lineFormat.weight = 0.5;
  }
  
  // 4. Phasen
  for (var p = 0; p < phases.length; p++) {
    var phase = phases[p];
    var rowTop = chartTop + (p * rowHeightPt);
    
    // Label
    var label = slide.shapes.addGeometricShape(
      PowerPoint.GeometricShapeType.rectangle,
      { left: GANTT_LEFT_PT, top: rowTop, width: labelWidthPt, height: rowHeightPt }
    );
    label.fill.setSolidColor("F0F0F0");
    label.lineFormat.color = "808080";
    label.lineFormat.weight = LINE_WEIGHT;
    
    formatTextFrame(label, phase.name, false);
    
    // Balken
    var phaseStart = daysBetween(projStart, phase.start);
    var phaseEnd = daysBetween(projStart, phase.end);
    if (phaseStart < 0) phaseStart = 0;
    if (phaseEnd > totalDays) phaseEnd = totalDays;
    
    if (phaseEnd > phaseStart) {
      var barX = (phaseStart / totalDays) * chartWidth;
      var barW = ((phaseEnd - phaseStart) / totalDays) * chartWidth;
      
      if (barX < chartWidth && barW > 0) {
        if (barX + barW > chartWidth) barW = chartWidth - barX;
        
        var bar = slide.shapes.addGeometricShape(
          PowerPoint.GeometricShapeType.rectangle,
          { left: chartLeft + barX, top: rowTop + barPadding, width: barW, height: barHeightPt }
        );
        
        var colorHex = phase.color.replace("#", "");
        bar.fill.setSolidColor(colorHex);
        bar.lineFormat.color = colorHex;
        bar.lineFormat.weight = LINE_WEIGHT;
        
        // Balken-Text: weiße Schrift auf farbigem Hintergrund
        try {
          bar.textFrame.textRange.text = phase.name;
          bar.textFrame.textRange.font.size = FONT_SIZE;
          bar.textFrame.textRange.font.bold = false;
          bar.textFrame.textRange.font.color = "FFFFFF";  // Weiß für Balken
          bar.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
          bar.textFrame.marginLeft = TEXT_MARGIN_LEFT;
          bar.textFrame.marginRight = TEXT_MARGIN_RIGHT;
          bar.textFrame.marginTop = TEXT_MARGIN_TOP;
          bar.textFrame.marginBottom = TEXT_MARGIN_BOTTOM;
        } catch(e) {}
      }
    }
  }
  
  // 5. Heute-Linie mit Datum-Label am unteren Ende
  if (showTodayLine) {
    var today = new Date();
    var todayDays = daysBetween(projStart, today);
    if (todayDays >= 0 && todayDays <= totalDays) {
      var todayX = (todayDays / totalDays) * chartWidth;
      if (todayX <= chartWidth) {
        // Linie
        var tl = slide.shapes.addLine(PowerPoint.ConnectorType.straight,
          { left: chartLeft + todayX, top: GANTT_TOP_PT, width: 0.01, height: totalHeight });
        tl.lineFormat.color = "FF0000";
        tl.lineFormat.weight = 2;
        
        // Datum-Label am unteren Ende
        var todayStr = pad2(today.getDate()) + "." + pad2(today.getMonth() + 1) + "." + today.getFullYear();
        var labelWidth = 60;  // Breite für Datum-Box
        var labelHeight = 14; // Höhe für Datum-Box
        var labelLeft = chartLeft + todayX - (labelWidth / 2);  // Zentriert unter der Linie
        var labelTop = GANTT_TOP_PT + totalHeight + 2;  // Direkt unter dem Diagramm
        
        var todayLabel = slide.shapes.addGeometricShape(
          PowerPoint.GeometricShapeType.rectangle,
          { left: labelLeft, top: labelTop, width: labelWidth, height: labelHeight }
        );
        todayLabel.fill.setSolidColor("FF0000");
        todayLabel.lineFormat.color = "FF0000";
        todayLabel.lineFormat.weight = LINE_WEIGHT;
        
        try {
          todayLabel.textFrame.textRange.text = todayStr;
          todayLabel.textFrame.textRange.font.size = 9;
          todayLabel.textFrame.textRange.font.bold = false;
          todayLabel.textFrame.textRange.font.color = "FFFFFF";
          todayLabel.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
          todayLabel.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
          todayLabel.textFrame.marginLeft = 0;
          todayLabel.textFrame.marginRight = 0;
          todayLabel.textFrame.marginTop = 0;
          todayLabel.textFrame.marginBottom = 0;
        } catch(e) {}
      }
    }
  }
}

// ═══════════════════════════════════════════════════════
// HILFSFUNKTIONEN
// ═══════════════════════════════════════════════════════
function computeTimeUnits(start, end, unit) {
  var units = [];
  var current = new Date(start);
  var maxUnits = 200;
  
  while (current < end && units.length < maxUnits) {
    var label = "";
    var unitStart = new Date(current);
    var unitEnd;
    
    switch (unit) {
      case "day":
        label = current.getDate() + "";
        unitEnd = addDays(current, 1);
        break;
      case "week":
        // v2.25: Nur Nummer, ohne "KW"
        label = "" + getWeekNumber(current);
        unitEnd = addDays(current, 7);
        break;
      case "month":
        label = getMonthShort(current.getMonth());
        unitEnd = new Date(current.getFullYear(), current.getMonth() + 1, 1);
        break;
      case "quarter":
        var q = Math.floor(current.getMonth() / 3) + 1;
        label = "Q" + q;
        unitEnd = new Date(current.getFullYear(), (q * 3), 1);
        break;
      default:
        label = current.getDate() + "";
        unitEnd = addDays(current, 1);
    }
    
    units.push({ label: label, start: unitStart, end: unitEnd });
    current = unitEnd;
  }
  
  return units;
}

function computeMonthGroups(timeUnits, unit) {
  var groups = [];
  var currentMonth = -1;
  var currentYear = -1;
  var count = 0;
  var label = "";
  
  for (var i = 0; i < timeUnits.length; i++) {
    var tu = timeUnits[i];
    var m = tu.start.getMonth();
    var y = tu.start.getFullYear();
    
    if (m !== currentMonth || y !== currentYear) {
      if (count > 0) groups.push({ label: label, count: count });
      currentMonth = m;
      currentYear = y;
      label = getMonthShort(m) + " " + y;
      count = 1;
    } else {
      count++;
    }
  }
  if (count > 0) groups.push({ label: label, count: count });
  
  return groups;
}

function daysBetween(d1, d2) {
  return Math.round((d2 - d1) / (24 * 60 * 60 * 1000));
}

function addDays(d, n) {
  var r = new Date(d);
  r.setDate(r.getDate() + n);
  return r;
}

function getWeekNumber(d) {
  var onejan = new Date(d.getFullYear(), 0, 1);
  return Math.ceil((((d - onejan) / 86400000) + onejan.getDay() + 1) / 7);
}

function getMonthShort(m) {
  return ["Jan","Feb","Mär","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez"][m];
}

function toISO(d) {
  return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
}

function pad2(n) {
  return n < 10 ? "0" + n : "" + n;
}

function escHtml(s) {
  return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

function randomColor() {
  var colors = ["#2e86c1","#27ae60","#e94560","#f39c12","#9b59b6","#1abc9c","#e74c3c","#3498db"];
  return colors[Math.floor(Math.random() * colors.length)];
}

function showStatus(msg, type) {
  var el = document.getElementById("status");
  if (el) {
    el.textContent = msg;
    el.className = "status " + type;
  }
  console.log("[Status] " + msg);
}

function ganttInfo(html, isError) {
  var el = document.getElementById("ganttInfo");
  if (el) {
    el.innerHTML = html;
    el.style.color = isError ? "#e94560" : "#333";
  }
}
