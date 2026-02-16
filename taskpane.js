/*
 ═══════════════════════════════════════════════════════
 Droege GANTT Generator  –  taskpane.js  v2.22

 UPDATES v2.22 (basierend auf v2.14):
  - Linienstärke 0.5 pt für alle Objekte (außer senkrechte Trennlinien)
  - REST-RAND Berücksichtigung: Grid-Margins für korrekte Positionierung
  - Feste Foliengröße 16:9 (960 x 540 pt) für konsistente Berechnung

 DROEGE GROUP · 2026
 ═══════════════════════════════════════════════════════
*/

var VERSION = "2.22";

// ═══════════════════════════════════════════════════════
// KONSTANTEN
// ═══════════════════════════════════════════════════════
var CM = 28.3464567;           // Points pro cm (exakt)
var gridUnitCm = 0.21;         // Rastereinheit in cm
var ganttPhaseCount = 0;

// GANTT Layout - FEST in Rastereinheiten
var GANTT_LEFT_RE = 9;         // Links: 9 RE
var GANTT_TOP_RE = 17;         // Oben: 17 RE
var GANTT_MAX_WIDTH_RE = 118;  // Max. Breite: 118 RE

// Schriftgröße für alle Texte
var FONT_SIZE = 11;

// LINIENSTÄRKE - NEU: 0.5 pt für alle Ränder
var LINE_WEIGHT = 0.5;

// Feste Foliengröße (16:9 Standard)
var SLIDE_WIDTH_PT = 960;
var SLIDE_HEIGHT_PT = 540;

// ═══════════════════════════════════════════════════════
// BERECHNETE GRID-MARGINS (REST-RAND)
// Diese Werte werden beim Start berechnet
// ═══════════════════════════════════════════════════════
var GRID_MARGIN_LEFT_PT = 0;
var GRID_MARGIN_TOP_PT = 0;

function calculateGridMargins() {
  var rePt = gridUnitCm * CM;
  var fullUnitsX = Math.floor(SLIDE_WIDTH_PT / rePt);
  var fullUnitsY = Math.floor(SLIDE_HEIGHT_PT / rePt);
  GRID_MARGIN_LEFT_PT = (SLIDE_WIDTH_PT - (fullUnitsX * rePt)) / 2;
  GRID_MARGIN_TOP_PT = (SLIDE_HEIGHT_PT - (fullUnitsY * rePt)) / 2;
  console.log("Grid-Margins berechnet:", {
    rePt: rePt,
    fullUnitsX: fullUnitsX,
    fullUnitsY: fullUnitsY,
    GRID_MARGIN_LEFT_PT: GRID_MARGIN_LEFT_PT,
    GRID_MARGIN_TOP_PT: GRID_MARGIN_TOP_PT
  });
}

// ═══════════════════════════════════════════════════════
// KONVERTIERUNGS-FUNKTIONEN MIT REST-RAND
// ═══════════════════════════════════════════════════════

// Konvertiert Rastereinheiten zu Points (ohne Margin - für Breiten/Höhen)
function re2pt(re) {
  return re * gridUnitCm * CM;
}

// Konvertiert cm zu Points
function cm2pt(c) {
  return c * CM;
}

// Konvertiert RE zu absoluter X-Position (MIT Grid-Margin)
function reToAbsoluteX(re) {
  return GRID_MARGIN_LEFT_PT + (re * gridUnitCm * CM);
}

// Konvertiert RE zu absoluter Y-Position (MIT Grid-Margin)
function reToAbsoluteY(re) {
  return GRID_MARGIN_TOP_PT + (re * gridUnitCm * CM);
}

// ═══════════════════════════════════════════════════════
// OFFICE READY
// ═══════════════════════════════════════════════════════
Office.onReady(function(info) {
  if (info.host === Office.HostType.PowerPoint) {
    calculateGridMargins();
    initUI();
    updateInfoBar();
    showStatus("Bereit", "success");
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
  // Grid Unit Input
  var gi = document.getElementById("gridUnit");
  if (gi) {
    gi.addEventListener("change", function() {
      var v = parseFloat(this.value);
      if (!isNaN(v) && v > 0) {
        gridUnitCm = v;
        calculateGridMargins(); // Neu berechnen bei Änderung
        updatePresetButtons(v);
      }
    });
  }
  
  // Preset Buttons
  document.querySelectorAll(".pre").forEach(function(b) {
    b.addEventListener("click", function() {
      var v = parseFloat(this.dataset.value);
      gridUnitCm = v;
      calculateGridMargins(); // Neu berechnen bei Änderung
      if (gi) gi.value = v;
      updatePresetButtons(v);
    });
  });

  // Width Mode Toggle
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

  // Buttons
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

// ═══════════════════════════════════════════════════════
// SLIDE SIZE
// ═══════════════════════════════════════════════════════
function setSlideSize() {
  showStatus("Setze Folienformat...", "working");
  PowerPoint.run(function(ctx) {
    // 16:9 Format
    // In der JS API nicht direkt setzbar - nur Info anzeigen
    return ctx.sync();
  }).then(function() {
    showStatus("Folienformat: 16:9 (Standard)", "success");
  }).catch(function(err) {
    showStatus("Fehler: " + err.message, "error");
  });
}

// ═══════════════════════════════════════════════════════
// CREATE GANTT CHART
// ═══════════════════════════════════════════════════════
function createGanttChart() {
  console.log("=== createGanttChart START v2.22 ===");
  
  // Grid-Margins neu berechnen (für aktuelle gridUnitCm)
  calculateGridMargins();
  
  // Eingaben lesen
  var projStart = new Date(document.getElementById("ganttStart").value);
  var projEnd = new Date(document.getElementById("ganttEnd").value);
  var unit = document.getElementById("ganttUnit").value;
  var labelWidthRE = parseInt(document.getElementById("ganttLabelW").value) || 20;
  var headerHeightRE = parseInt(document.getElementById("ganttHeaderH").value) || 3;
  var barHeightRE = parseInt(document.getElementById("ganttBarH").value) || 3;
  var rowHeightRE = parseInt(document.getElementById("ganttRowH").value) || 5;
  var showTodayLine = document.getElementById("ganttTodayLine").checked;
  
  // Spaltenbreiten-Modus
  var widthMode = document.getElementById("ganttWidthMode").value;
  var colWidthRE = parseInt(document.getElementById("ganttColW").value) || 3;

  // Validierung
  if (isNaN(projStart.getTime()) || isNaN(projEnd.getTime())) {
    ganttInfo("Bitte Start- und Enddatum eingeben!", true);
    return;
  }
  if (projEnd <= projStart) {
    ganttInfo("Ende muss nach Start liegen!", true);
    return;
  }

  var phases = getPhases();
  console.log("Phasen:", phases.length);
  
  if (phases.length === 0) {
    ganttInfo("Mindestens eine Phase hinzufügen!", true);
    return;
  }

  // Zeiteinheiten berechnen
  var timeUnits = computeTimeUnits(projStart, projEnd, unit);
  console.log("Zeiteinheiten:", timeUnits.length);
  
  if (timeUnits.length === 0 || timeUnits.length > 200) {
    ganttInfo("Zeitraum anpassen oder andere Einheit wählen!", true);
    return;
  }

  var totalDays = daysBetween(projStart, projEnd);
  if (totalDays < 1) totalDays = 1;

  // ═══ SPALTENBREITE BERECHNEN ═══
  var availableWidthRE = GANTT_MAX_WIDTH_RE - labelWidthRE;
  
  if (widthMode === "auto") {
    colWidthRE = Math.floor(availableWidthRE / timeUnits.length);
    if (colWidthRE < 1) colWidthRE = 1;
    if (colWidthRE > 10) colWidthRE = 10;
    console.log("Auto-Modus: Spaltenbreite = " + colWidthRE + " RE");
  }
  
  var maxVisibleColumns = Math.floor(availableWidthRE / colWidthRE);
  var visibleColumns = Math.min(timeUnits.length, maxVisibleColumns);
  var truncated = visibleColumns < timeUnits.length;
  
  console.log("Sichtbare Spalten: " + visibleColumns + " von " + timeUnits.length);

  // Info anzeigen
  var unitNames = {day:"Tage", week:"Wochen", month:"Monate", quarter:"Quartale"};
  var infoText = "<b>" + phases.length + "</b> Phasen, <b>" + visibleColumns + "</b> " + unitNames[unit];
  if (truncated) {
    infoText += " <span style='color:#e94560'>(von " + timeUnits.length + " – abgeschnitten)</span>";
  }
  infoText += "<br>Spaltenbreite: <b>" + colWidthRE + " RE</b>";
  infoText += "<br>Grid-Margin: " + GRID_MARGIN_LEFT_PT.toFixed(2) + " x " + GRID_MARGIN_TOP_PT.toFixed(2) + " pt";
  ganttInfo(infoText, false);

  showStatus("Erstelle GANTT auf aktueller Folie...", "working");

  PowerPoint.run(function(ctx) {
    var slide = ctx.presentation.slides.getItemAt(0);
    
    drawGantt(ctx, slide, projStart, projEnd, unit, phases, timeUnits, 
              labelWidthRE, headerHeightRE, barHeightRE, rowHeightRE, 
              colWidthRE, totalDays, showTodayLine, visibleColumns, truncated);
    
    return ctx.sync();
  }).then(function() {
    console.log("=== createGanttChart DONE ===");
    showStatus("GANTT erstellt!", "success");
  }).catch(function(err) {
    console.error("Fehler:", err);
    showStatus("Fehler: " + err.message, "error");
  });
}

// ═══════════════════════════════════════════════════════
// DRAW GANTT
// ═══════════════════════════════════════════════════════
function drawGantt(ctx, slide, projStart, projEnd, unit, phases, timeUnits, 
                   labelWidthRE, headerHeightRE, barHeightRE, rowHeightRE, 
                   colWidthRE, totalDays, showTodayLine, visibleColumns, truncated) {
  
  // ═══ POSITIONIERUNG MIT REST-RAND ═══
  var GANTT_LEFT_PT = reToAbsoluteX(GANTT_LEFT_RE);
  var GANTT_TOP_PT = reToAbsoluteY(GANTT_TOP_RE);
  
  // Dimensionen in Points (ohne Margin - nur Größen)
  var labelWidthPt = re2pt(labelWidthRE);
  var headerHeightPt = re2pt(headerHeightRE);
  var barHeightPt = re2pt(barHeightRE);
  var rowHeightPt = re2pt(rowHeightRE);
  var colWidthPt = re2pt(colWidthRE);
  
  // Balken-Padding
  var barPadding = Math.max(2, Math.round((rowHeightPt - barHeightPt) / 2));
  
  var chartLeft = GANTT_LEFT_PT + labelWidthPt;
  var chartWidth = visibleColumns * colWidthPt;
  var totalWidth = labelWidthPt + chartWidth;
  
  // Monatszeile
  var needsMonthRow = (unit === "day" || unit === "week" || unit === "quarter");
  var monthRowHeightPt = needsMonthRow ? headerHeightPt : 0;
  
  var totalHeaderHeight = monthRowHeightPt + headerHeightPt;
  var chartTop = GANTT_TOP_PT + totalHeaderHeight;
  var totalHeight = totalHeaderHeight + (phases.length * rowHeightPt);
  var chartBottom = GANTT_TOP_PT + totalHeight;
  var lineHeight = chartBottom - (GANTT_TOP_PT + monthRowHeightPt + headerHeightPt);
  
  console.log("Layout v2.22:", {
    GANTT_LEFT_PT: GANTT_LEFT_PT,
    GANTT_TOP_PT: GANTT_TOP_PT,
    GRID_MARGIN_LEFT_PT: GRID_MARGIN_LEFT_PT,
    GRID_MARGIN_TOP_PT: GRID_MARGIN_TOP_PT,
    totalWidth: totalWidth,
    totalHeight: totalHeight
  });

  // ═══ 1. HINTERGRUND (mit 0.5pt Rand) ═══
  console.log("1. Hintergrund");
  var bg = slide.shapes.addGeometricShape(
    PowerPoint.GeometricShapeType.rectangle,
    {
      left: GANTT_LEFT_PT,
      top: GANTT_TOP_PT,
      width: totalWidth,
      height: totalHeight
    }
  );
  bg.fill.setSolidColor("FFFFFF");
  bg.lineFormat.color = "808080";
  bg.lineFormat.weight = LINE_WEIGHT;
  
  // ═══ 2. MONATSZEILE (wenn benötigt) ═══
  if (needsMonthRow) {
    console.log("2a. Monatszeile");
    var monthGroups = computeMonthGroups(timeUnits, unit);
    var monthX = 0;
    
    for (var m = 0; m < monthGroups.length; m++) {
      var mg = monthGroups[m];
      var monthWidth = mg.count * colWidthPt;
      
      var monthCell = slide.shapes.addGeometricShape(
        PowerPoint.GeometricShapeType.rectangle,
        {
          left: chartLeft + monthX,
          top: GANTT_TOP_PT,
          width: monthWidth,
          height: monthRowHeightPt
        }
      );
      monthCell.fill.setSolidColor("B0B0B0");
      monthCell.lineFormat.color = "808080";
      monthCell.lineFormat.weight = LINE_WEIGHT;
      
      try {
        monthCell.textFrame.textRange.text = mg.label;
        monthCell.textFrame.textRange.font.size = FONT_SIZE;
        monthCell.textFrame.textRange.font.bold = true;
        monthCell.textFrame.textRange.font.color = "000000";
        monthCell.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
        monthCell.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
      } catch(e) {}
      
      monthX += monthWidth;
    }
  }
  
  // ═══ 2b. HEADER-ZELLEN ═══
  console.log("2b. Header-Zellen: " + timeUnits.length);
  var linePositions = [];
  var headerTop = GANTT_TOP_PT + monthRowHeightPt;
  
  for (var c = 0; c < visibleColumns; c++) {
    var colX = c * colWidthPt;
    
    var hdr = slide.shapes.addGeometricShape(
      PowerPoint.GeometricShapeType.rectangle,
      {
        left: chartLeft + colX,
        top: headerTop,
        width: colWidthPt,
        height: headerHeightPt
      }
    );
    hdr.fill.setSolidColor("E0E0E0");
    hdr.lineFormat.color = "808080";
    hdr.lineFormat.weight = LINE_WEIGHT;
    
    try {
      hdr.textFrame.textRange.text = timeUnits[c].label;
      hdr.textFrame.textRange.font.size = FONT_SIZE;
      hdr.textFrame.textRange.font.bold = false;
      hdr.textFrame.textRange.font.color = "000000";
      hdr.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
      hdr.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
    } catch(e) {}
    
    linePositions.push(chartLeft + colX);
  }
  
  // ═══ 3. SENKRECHTE TRENNLINIEN (OHNE Linienstärke-Einstellung = Standard) ═══
  console.log("3. Senkrechte Linien: " + linePositions.length);
  var lineTop = GANTT_TOP_PT + monthRowHeightPt + headerHeightPt;
  
  for (var i = 0; i < linePositions.length; i++) {
    var line = slide.shapes.addLine(
      PowerPoint.ConnectorType.straight,
      {
        left: linePositions[i],
        top: lineTop,
        width: 0,
        height: lineHeight
      }
    );
    line.lineFormat.color = "D0D0D0";
    // KEINE lineFormat.weight Einstellung = Standard-Linienstärke
  }
  
  // Rechte Abschlusslinie
  var rightLine = slide.shapes.addLine(
    PowerPoint.ConnectorType.straight,
    {
      left: chartLeft + chartWidth,
      top: lineTop,
      width: 0,
      height: lineHeight
    }
  );
  rightLine.lineFormat.color = "D0D0D0";
  // KEINE lineFormat.weight Einstellung
  
  // ═══ 4. LABEL-SPALTE ═══
  console.log("4. Label-Spalte");
  
  // Label-Header
  var labelHeader = slide.shapes.addGeometricShape(
    PowerPoint.GeometricShapeType.rectangle,
    {
      left: GANTT_LEFT_PT,
      top: GANTT_TOP_PT,
      width: labelWidthPt,
      height: totalHeaderHeight
    }
  );
  labelHeader.fill.setSolidColor("C0C0C0");
  labelHeader.lineFormat.color = "808080";
  labelHeader.lineFormat.weight = LINE_WEIGHT;
  
  try {
    labelHeader.textFrame.textRange.text = "Phase";
    labelHeader.textFrame.textRange.font.size = FONT_SIZE;
    labelHeader.textFrame.textRange.font.bold = true;
    labelHeader.textFrame.textRange.font.color = "000000";
    labelHeader.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
    labelHeader.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
  } catch(e) {}
  
  // ═══ 5. PHASEN-ZEILEN ═══
  console.log("5. Phasen-Zeilen: " + phases.length);
  
  for (var p = 0; p < phases.length; p++) {
    var phase = phases[p];
    var rowTop = chartTop + (p * rowHeightPt);
    
    // Label-Zelle
    var labelCell = slide.shapes.addGeometricShape(
      PowerPoint.GeometricShapeType.rectangle,
      {
        left: GANTT_LEFT_PT,
        top: rowTop,
        width: labelWidthPt,
        height: rowHeightPt
      }
    );
    labelCell.fill.setSolidColor("F5F5F5");
    labelCell.lineFormat.color = "808080";
    labelCell.lineFormat.weight = LINE_WEIGHT;
    
    try {
      labelCell.textFrame.textRange.text = phase.name;
      labelCell.textFrame.textRange.font.size = FONT_SIZE;
      labelCell.textFrame.textRange.font.color = "000000";
      labelCell.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
      labelCell.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.left;
      labelCell.textFrame.marginLeft = 5;
    } catch(e) {}
    
    // Balken berechnen
    var barStartDay = daysBetween(projStart, phase.start);
    var barEndDay = daysBetween(projStart, phase.end);
    
    if (barStartDay < 0) barStartDay = 0;
    if (barEndDay > totalDays) barEndDay = totalDays;
    
    if (barEndDay > barStartDay) {
      var barStartX = (barStartDay / totalDays) * chartWidth;
      var barWidth = ((barEndDay - barStartDay) / totalDays) * chartWidth;
      
      // Auf sichtbaren Bereich begrenzen
      if (barStartX + barWidth > chartWidth) {
        barWidth = chartWidth - barStartX;
      }
      if (barStartX < chartWidth && barWidth > 0) {
        var bar = slide.shapes.addGeometricShape(
          PowerPoint.GeometricShapeType.rectangle,
          {
            left: chartLeft + barStartX,
            top: rowTop + barPadding,
            width: barWidth,
            height: barHeightPt
          }
        );
        bar.fill.setSolidColor(phase.color.replace("#", ""));
        bar.lineFormat.color = phase.color.replace("#", "");
        bar.lineFormat.weight = LINE_WEIGHT;
      }
    }
  }
  
  // ═══ 6. HEUTE-LINIE ═══
  if (showTodayLine) {
    var today = new Date();
    var todayDay = daysBetween(projStart, today);
    
    if (todayDay >= 0 && todayDay <= totalDays) {
      var todayX = (todayDay / totalDays) * chartWidth;
      
      if (todayX <= chartWidth) {
        console.log("6. Heute-Linie bei Tag " + todayDay);
        var todayLine = slide.shapes.addLine(
          PowerPoint.ConnectorType.straight,
          {
            left: chartLeft + todayX,
            top: chartTop,
            width: 0,
            height: phases.length * rowHeightPt
          }
        );
        todayLine.lineFormat.color = "E94560";
        todayLine.lineFormat.weight = 2; // Heute-Linie dicker für Sichtbarkeit
      }
    }
  }
  
  // ═══ 7. ABSCHNITT-INDIKATOR ═══
  if (truncated) {
    console.log("7. Truncated-Indikator");
    var truncIndicator = slide.shapes.addGeometricShape(
      PowerPoint.GeometricShapeType.rectangle,
      {
        left: chartLeft + chartWidth - 20,
        top: GANTT_TOP_PT,
        width: 20,
        height: totalHeight
      }
    );
    truncIndicator.fill.setSolidColor("FFEEEE");
    truncIndicator.lineFormat.color = "E94560";
    truncIndicator.lineFormat.weight = LINE_WEIGHT;
    truncIndicator.fill.transparency = 0.5;
    
    try {
      truncIndicator.textFrame.textRange.text = "...";
      truncIndicator.textFrame.textRange.font.size = 14;
      truncIndicator.textFrame.textRange.font.bold = true;
      truncIndicator.textFrame.textRange.font.color = "E94560";
      truncIndicator.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
      truncIndicator.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
    } catch(e) {}
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
        label = "KW" + getWeekNumber(current);
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
    
    units.push({
      label: label,
      start: unitStart,
      end: unitEnd
    });
    
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
      if (count > 0) {
        groups.push({ label: label, count: count });
      }
      currentMonth = m;
      currentYear = y;
      label = getMonthShort(m) + " " + y;
      count = 1;
    } else {
      count++;
    }
  }
  
  if (count > 0) {
    groups.push({ label: label, count: count });
  }
  
  return groups;
}

function daysBetween(d1, d2) {
  var oneDay = 24 * 60 * 60 * 1000;
  return Math.round((d2 - d1) / oneDay);
}

function addDays(date, days) {
  var result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

function getWeekNumber(d) {
  var onejan = new Date(d.getFullYear(), 0, 1);
  var millisecsInDay = 86400000;
  return Math.ceil((((d - onejan) / millisecsInDay) + onejan.getDay() + 1) / 7);
}

function getMonthShort(m) {
  var months = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"];
  return months[m];
}

function toISO(d) {
  return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
}

function pad2(n) {
  return n < 10 ? "0" + n : "" + n;
}

function escHtml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function randomColor() {
  var colors = ["#2e86c1", "#27ae60", "#e94560", "#f39c12", "#9b59b6", "#1abc9c", "#e74c3c", "#3498db"];
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
