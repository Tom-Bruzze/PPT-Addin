/*
 ═══════════════════════════════════════════════════════
 Droege GANTT Generator  –  taskpane.js  v2.17

 UPDATES v2.17:
  - MARGIN-BERECHNUNG DIREKT IN drawGantt() - garantiert korrekt
  - Einheitliche Linienstärke 0.5 Pt für alle Objekte
  - Keine globalen Variablen für Margins mehr

 DROEGE GROUP · 2026
 ═══════════════════════════════════════════════════════
*/

var VERSION = "2.17";

// ═══════════════════════════════════════════
// EXAKTE KONSTANTEN (keine Rundung!)
// ═══════════════════════════════════════════
var POINTS_PER_CM = 28.3464567;
var RE_CM = 0.21;
var RE_PT = RE_CM * POINTS_PER_CM;  // 5.95275590551 points exakt

// Legacy-Kompatibilität
var CM = POINTS_PER_CM;
var gridUnitCm = RE_CM;
var ganttPhaseCount = 0;

// GANTT Layout - FEST in Rastereinheiten
var GANTT_LEFT_RE = 9;         // Links: 9 RE vom Raster-Ursprung
var GANTT_TOP_RE = 17;         // Oben: 17 RE vom Raster-Ursprung
var GANTT_MAX_WIDTH_RE = 118;  // Max. Breite: 118 RE

// Schrift und Linien
var FONT_SIZE = 11;
var LINE_WEIGHT = 0.5;  // Einheitliche Linienstärke für alle Ränder

// ═══════════════════════════════════════════
// HELPER: RE zu Points (für Größen)
// ═══════════════════════════════════════════
function re2pt(re) {
  return re * RE_PT;
}

// ═══════════════════════════════════════════
// OFFICE INIT
// ═══════════════════════════════════════════

Office.onReady(function(info) {
  if (info.host === Office.HostType.PowerPoint) {
    initUI();
    updateInfoBar();
    showStatus("Bereit (v" + VERSION + ")", "success");
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
        updatePresetButtons(v);
      }
    });
  }
  
  // Preset Buttons
  document.querySelectorAll(".pre").forEach(function(b) {
    b.addEventListener("click", function() {
      var v = parseFloat(this.dataset.value);
      gridUnitCm = v;
      document.getElementById("gridUnit").value = v;
      updatePresetButtons(v);
    });
  });
  
  // Insert Rectangle
  var ir = document.getElementById("insertRect");
  if (ir) ir.addEventListener("click", insertRectangle);
  
  // Insert Table
  var it = document.getElementById("insertTable");
  if (it) it.addEventListener("click", insertTable);
  
  // Set Slide Size
  var ss = document.getElementById("setSlide");
  if (ss) ss.addEventListener("click", function() {
    var sel = document.getElementById("slideSize").value;
    setSlideSize(sel);
  });
  
  // GANTT Generate - KORRIGIERTE ID
  var btnGantt = document.getElementById("createGantt");
  if (btnGantt) {
    btnGantt.addEventListener("click", generateGantt);
    console.log("✓ GANTT Button gebunden");
  } else {
    console.error("✗ GANTT Button nicht gefunden!");
  }
  
  // GANTT Add Phase
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
    '<input type="text" value="' + escHtml(name) + '" placeholder="Name">' +
    '<input type="date" value="' + toISO(start) + '">' +
    '<input type="date" value="' + toISO(end) + '">' +
    '<input type="color" value="' + color + '">' +
    '<button class="gantt-del">&times;</button>';
  container.appendChild(div);
  
  div.querySelector(".gantt-del").addEventListener("click", function() {
    div.remove();
  });
}

// ═══════════════════════════════════════════
// HELPER
// ═══════════════════════════════════════════

function pad2(n) { return n < 10 ? "0" + n : String(n); }
function toISO(d) { return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate()); }
function addDays(d, n) { var r = new Date(d); r.setDate(r.getDate() + n); return r; }
function daysBetween(a, b) { return Math.floor((b - a) / 86400000); }
function randomColor() { 
  var cols = ["#2e86c1","#27ae60","#e94560","#f39c12","#9b59b6","#1abc9c","#34495e"];
  return cols[Math.floor(Math.random() * cols.length)];
}
function escHtml(s) { return String(s).replace(/"/g, "&quot;"); }

function showStatus(msg, type) {
  var el = document.getElementById("status");
  if (el) {
    el.textContent = msg;
    el.className = "status " + type;
  }
}

function ganttInfo(msg, isError) {
  var el = document.getElementById("ganttInfo");
  if (el) {
    el.innerHTML = msg;
    el.className = "gantt-info" + (isError ? " error" : "");
  }
}

function getPhases() {
  var phases = [];
  var rows = document.querySelectorAll(".gantt-phase");
  for (var i = 0; i < rows.length; i++) {
    var inputs = rows[i].querySelectorAll("input");
    var name = inputs[0].value.trim() || "Phase";
    var start = new Date(inputs[1].value);
    var end = new Date(inputs[2].value);
    var color = inputs[3].value || "#2e86c1";
    
    if (!isNaN(start.getTime()) && !isNaN(end.getTime()) && end > start) {
      phases.push({ name: name, start: start, end: end, color: color });
    }
  }
  return phases;
}

function getWeekNumber(d) {
  var tmp = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  tmp.setDate(tmp.getDate() + 3 - ((tmp.getDay() + 6) % 7));
  var week1 = new Date(tmp.getFullYear(), 0, 4);
  return 1 + Math.round(((tmp - week1) / 86400000 - 3 + ((week1.getDay() + 6) % 7)) / 7);
}

// ═══════════════════════════════════════════
// SLIDE SIZE
// ═══════════════════════════════════════════

async function setSlideSize(size) {
  showStatus("Setze Foliengröße...", "info");
  try {
    await PowerPoint.run(async function(context) {
      var pres = context.presentation;
      if (size === "16:9") {
        pres.slideWidth = 33.867 * CM;
        pres.slideHeight = 19.05 * CM;
      } else if (size === "4:3") {
        pres.slideWidth = 25.4 * CM;
        pres.slideHeight = 19.05 * CM;
      } else if (size === "A4") {
        pres.slideWidth = 29.7 * CM;
        pres.slideHeight = 21.0 * CM;
      }
      await context.sync();
    });
    showStatus("Foliengröße gesetzt", "success");
  } catch (e) {
    showStatus("Fehler: " + e.message, "error");
  }
}

// ═══════════════════════════════════════════
// INSERT RECTANGLE
// ═══════════════════════════════════════════

async function insertRectangle() {
  showStatus("Füge Rechteck ein...", "info");
  try {
    await PowerPoint.run(async function(context) {
      var presentation = context.presentation;
      presentation.load("slideWidth,slideHeight");
      await context.sync();
      
      // Berechne Margin HIER
      var slideWidth = presentation.slideWidth;
      var slideHeight = presentation.slideHeight;
      var fullUnitsX = Math.floor(slideWidth / RE_PT);
      var fullUnitsY = Math.floor(slideHeight / RE_PT);
      var marginLeft = (slideWidth - (fullUnitsX * RE_PT)) / 2;
      var marginTop = (slideHeight - (fullUnitsY * RE_PT)) / 2;
      
      var slide = presentation.getSelectedSlides().getItemAt(0);
      
      // Position mit Margin
      var leftPt = marginLeft + (GANTT_LEFT_RE * RE_PT);
      var topPt = marginTop + (GANTT_TOP_RE * RE_PT);
      var widthPt = re2pt(10);
      var heightPt = re2pt(5);
      
      var shape = slide.shapes.addGeometricShape(
        PowerPoint.GeometricShapeType.rectangle,
        {
          left: leftPt,
          top: topPt,
          width: widthPt,
          height: heightPt
        }
      );
      shape.fill.setSolidColor("1A3A5C");
      shape.lineFormat.color = "1A3A5C";
      shape.lineFormat.weight = LINE_WEIGHT;
      
      await context.sync();
    });
    showStatus("Rechteck eingefügt", "success");
  } catch (e) {
    showStatus("Fehler: " + e.message, "error");
  }
}

// ═══════════════════════════════════════════
// INSERT TABLE
// ═══════════════════════════════════════════

async function insertTable() {
  showStatus("Füge Tabelle ein...", "info");
  try {
    await PowerPoint.run(async function(context) {
      var presentation = context.presentation;
      presentation.load("slideWidth,slideHeight");
      await context.sync();
      
      // Berechne Margin HIER
      var slideWidth = presentation.slideWidth;
      var slideHeight = presentation.slideHeight;
      var fullUnitsX = Math.floor(slideWidth / RE_PT);
      var fullUnitsY = Math.floor(slideHeight / RE_PT);
      var marginLeft = (slideWidth - (fullUnitsX * RE_PT)) / 2;
      var marginTop = (slideHeight - (fullUnitsY * RE_PT)) / 2;
      
      var slide = presentation.getSelectedSlides().getItemAt(0);
      
      var leftPt = marginLeft + (GANTT_LEFT_RE * RE_PT);
      var topPt = marginTop + (GANTT_TOP_RE * RE_PT);
      
      var cellWidth = re2pt(8);
      var cellHeight = re2pt(3);
      
      for (var row = 0; row < 3; row++) {
        for (var col = 0; col < 3; col++) {
          var x = leftPt + (col * cellWidth);
          var y = topPt + (row * cellHeight);
          
          var cell = slide.shapes.addGeometricShape(
            PowerPoint.GeometricShapeType.rectangle,
            { left: x, top: y, width: cellWidth, height: cellHeight }
          );
          cell.fill.setSolidColor("FFFFFF");
          cell.lineFormat.color = "1A3A5C";
          cell.lineFormat.weight = LINE_WEIGHT;
        }
      }
      
      await context.sync();
    });
    showStatus("Tabelle eingefügt", "success");
  } catch (e) {
    showStatus("Fehler: " + e.message, "error");
  }
}

// ═══════════════════════════════════════════
// GANTT GENERATOR
// ═══════════════════════════════════════════

async function generateGantt() {
  console.log("generateGantt() aufgerufen");
  showStatus("Generiere GANTT...", "info");
  ganttInfo("Generiere...", false);
  
  var phases = getPhases();
  if (phases.length === 0) {
    ganttInfo("Keine gültigen Phasen definiert!", true);
    showStatus("Fehler: Keine Phasen", "error");
    return;
  }
  
  var projStart = new Date(document.getElementById("ganttStart").value);
  var projEnd = new Date(document.getElementById("ganttEnd").value);
  
  if (isNaN(projStart.getTime()) || isNaN(projEnd.getTime()) || projEnd <= projStart) {
    ganttInfo("Ungültiger Zeitraum!", true);
    showStatus("Fehler: Zeitraum", "error");
    return;
  }
  
  var unit = document.getElementById("ganttUnit").value;
  var timeUnits = computeTimeUnits(projStart, projEnd, unit);
  
  if (timeUnits.length === 0) {
    ganttInfo("Keine Zeiteinheiten berechnet!", true);
    return;
  }
  
  try {
    await PowerPoint.run(async function(context) {
      // ═══════════════════════════════════════════
      // MARGIN DIREKT HIER BERECHNEN!
      // ═══════════════════════════════════════════
      var presentation = context.presentation;
      presentation.load("slideWidth,slideHeight");
      await context.sync();
      
      var slideWidth = presentation.slideWidth;
      var slideHeight = presentation.slideHeight;
      
      // Berechne Rest-Rand
      var fullUnitsX = Math.floor(slideWidth / RE_PT);
      var fullUnitsY = Math.floor(slideHeight / RE_PT);
      var marginLeft = (slideWidth - (fullUnitsX * RE_PT)) / 2;
      var marginTop = (slideHeight - (fullUnitsY * RE_PT)) / 2;
      
      console.log("═══════════════════════════════════════════");
      console.log("GANTT v2.17 - MARGIN BERECHNUNG");
      console.log("  Folie: " + (slideWidth/CM).toFixed(2) + " x " + (slideHeight/CM).toFixed(2) + " cm");
      console.log("  Margin: " + (marginLeft/CM).toFixed(4) + " x " + (marginTop/CM).toFixed(4) + " cm");
      console.log("═══════════════════════════════════════════");
      
      await drawGantt(context, phases, timeUnits, projStart, projEnd, unit, marginLeft, marginTop);
    });
    ganttInfo("GANTT erstellt: " + phases.length + " Phasen, " + timeUnits.length + " Spalten", false);
    showStatus("GANTT erstellt", "success");
  } catch (e) {
    ganttInfo("Fehler: " + e.message, true);
    showStatus("Fehler: " + e.message, "error");
    console.error(e);
  }
}

async function drawGantt(ctx, phases, timeUnits, projStart, projEnd, unit, marginLeft, marginTop) {
  var slide = ctx.presentation.getSelectedSlides().getItemAt(0);
  
  // Layout in RE
  var labelWidthRE = 15;
  var headerHeightRE = 3;
  var barHeightRE = 2;
  var rowHeightRE = 3;
  var colWidthRE = 4;
  
  // Berechne verfügbare Breite und sichtbare Spalten
  var availableWidthRE = GANTT_MAX_WIDTH_RE - labelWidthRE;
  var visibleColumns = Math.min(timeUnits.length, Math.floor(availableWidthRE / colWidthRE));
  
  if (visibleColumns < timeUnits.length) {
    timeUnits = timeUnits.slice(0, visibleColumns);
    console.log("Spalten beschränkt auf " + visibleColumns);
  }
  
  // ═══ ABSOLUTE POSITIONIERUNG MIT ÜBERGEBENEM MARGIN ═══
  var GANTT_LEFT_PT = marginLeft + (GANTT_LEFT_RE * RE_PT);
  var GANTT_TOP_PT = marginTop + (GANTT_TOP_RE * RE_PT);
  
  console.log("GANTT Position: " + (GANTT_LEFT_PT/CM).toFixed(4) + " x " + (GANTT_TOP_PT/CM).toFixed(4) + " cm");
  
  // Dimensionen in Points
  var labelWidthPt = re2pt(labelWidthRE);
  var headerHeightPt = re2pt(headerHeightRE);
  var barHeightPt = re2pt(barHeightRE);
  var rowHeightPt = re2pt(rowHeightRE);
  var colWidthPt = re2pt(colWidthRE);
  
  var barPadding = (rowHeightPt - barHeightPt) / 2;
  var chartLeft = GANTT_LEFT_PT + labelWidthPt;
  var chartWidth = visibleColumns * colWidthPt;
  var totalWidth = labelWidthPt + chartWidth;
  
  // Monatszeile bei bestimmten Einheiten
  var needsMonthRow = (unit === "day" || unit === "week" || unit === "quarter");
  var monthRowHeightPt = needsMonthRow ? headerHeightPt : 0;
  var totalHeaderHeight = monthRowHeightPt + headerHeightPt;
  var chartTop = GANTT_TOP_PT + totalHeaderHeight;
  var totalHeight = totalHeaderHeight + (phases.length * rowHeightPt);
  
  // ═══ 1. HINTERGRUND ═══
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
  
  // ═══ 2. MONATSZEILE ═══
  if (needsMonthRow) {
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
  
  // ═══ 3. HEADER-ZELLEN ═══
  var colX = 0;
  var linePositions = [];
  var headerTop = GANTT_TOP_PT + monthRowHeightPt;
  
  for (var c = 0; c < timeUnits.length; c++) {
    var hdr = slide.shapes.addGeometricShape(
      PowerPoint.GeometricShapeType.rectangle,
      {
        left: chartLeft + colX,
        top: headerTop,
        width: colWidthPt,
        height: headerHeightPt
      }
    );
    hdr.fill.setSolidColor("D5D5D5");
    hdr.lineFormat.color = "808080";
    hdr.lineFormat.weight = LINE_WEIGHT;
    
    try {
      hdr.textFrame.textRange.text = timeUnits[c].label;
      hdr.textFrame.textRange.font.size = FONT_SIZE;
      hdr.textFrame.textRange.font.bold = true;
      hdr.textFrame.textRange.font.color = "333333";
      hdr.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
      hdr.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
    } catch(e) {}
    
    if (c > 0) {
      linePositions.push(chartLeft + colX);
    }
    
    colX += colWidthPt;
  }
  
  // ═══ 4. ZEILEN UND LABELS ═══
  var totalDays = daysBetween(projStart, projEnd);
  
  for (var i = 0; i < phases.length; i++) {
    var p = phases[i];
    var rowTop = chartTop + (i * rowHeightPt);
    
    // Zeilenhintergrund
    var rowBg = slide.shapes.addGeometricShape(
      PowerPoint.GeometricShapeType.rectangle,
      {
        left: GANTT_LEFT_PT,
        top: rowTop,
        width: totalWidth,
        height: rowHeightPt
      }
    );
    rowBg.fill.setSolidColor(i % 2 === 0 ? "F8F8F8" : "FFFFFF");
    rowBg.lineFormat.color = "D0D0D0";
    rowBg.lineFormat.weight = LINE_WEIGHT;
    
    // Label
    var label = slide.shapes.addGeometricShape(
      PowerPoint.GeometricShapeType.rectangle,
      {
        left: GANTT_LEFT_PT,
        top: rowTop,
        width: labelWidthPt,
        height: rowHeightPt
      }
    );
    label.fill.setSolidColor("E8E8E8");
    label.lineFormat.color = "808080";
    label.lineFormat.weight = LINE_WEIGHT;
    
    try {
      label.textFrame.textRange.text = p.name;
      label.textFrame.textRange.font.size = FONT_SIZE;
      label.textFrame.textRange.font.bold = true;
      label.textFrame.textRange.font.color = "333333";
      label.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
      label.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.left;
      label.textFrame.marginLeft = 5;
    } catch(e) {}
    
    // ═══ BALKEN ═══
    var phaseStartDay = Math.max(0, daysBetween(projStart, p.start));
    var phaseEndDay = Math.min(totalDays, daysBetween(projStart, p.end));
    
    if (phaseEndDay > phaseStartDay) {
      var barStartPct = phaseStartDay / totalDays;
      var barEndPct = phaseEndDay / totalDays;
      
      var barX = chartLeft + (barStartPct * chartWidth);
      var barW = (barEndPct - barStartPct) * chartWidth;
      var barY = rowTop + barPadding;
      
      if (barW < 4) barW = 4;
      
      var fillColor = p.color.replace("#", "");
      
      var bar = slide.shapes.addGeometricShape(
        PowerPoint.GeometricShapeType.rectangle,
        {
          left: barX,
          top: barY,
          width: barW,
          height: barHeightPt
        }
      );
      bar.fill.setSolidColor(fillColor);
      bar.lineFormat.color = fillColor;
      bar.lineFormat.weight = LINE_WEIGHT;
    }
  }
  
  // ═══ 5. VERTIKALE TRENNLINIEN ═══
  for (var l = 0; l < linePositions.length; l++) {
    var line = slide.shapes.addLine(
      PowerPoint.ConnectorType.straight,
      {
        left: linePositions[l],
        top: chartTop,
        width: 0,
        height: (phases.length * rowHeightPt)
      }
    );
    line.lineFormat.color = "C0C0C0";
    line.lineFormat.weight = LINE_WEIGHT;
  }
  
  // ═══ 6. HEUTE-LINIE ═══
  var today = new Date();
  if (today >= projStart && today <= projEnd) {
    var todayDays = daysBetween(projStart, today);
    var todayPct = todayDays / totalDays;
    var todayX = chartLeft + (todayPct * chartWidth);
    
    var todayLine = slide.shapes.addLine(
      PowerPoint.ConnectorType.straight,
      {
        left: todayX,
        top: GANTT_TOP_PT,
        width: 0,
        height: totalHeight
      }
    );
    todayLine.lineFormat.color = "FF0000";
    todayLine.lineFormat.weight = 2;
    
    try {
      var todayLabel = slide.shapes.addGeometricShape(
        PowerPoint.GeometricShapeType.rectangle,
        {
          left: todayX - 20,
          top: GANTT_TOP_PT + totalHeight + 4,
          width: 40,
          height: 16
        }
      );
      todayLabel.fill.setSolidColor("FF0000");
      todayLabel.lineFormat.color = "FF0000";
      todayLabel.lineFormat.weight = LINE_WEIGHT;
      todayLabel.textFrame.textRange.text = "Heute";
      todayLabel.textFrame.textRange.font.size = 9;
      todayLabel.textFrame.textRange.font.color = "FFFFFF";
      todayLabel.textFrame.textRange.font.bold = true;
      todayLabel.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
      todayLabel.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
    } catch(e) {}
  }
  
  return ctx.sync();
}

// ═══════════════════════════════════════════
// COMPUTE FUNCTIONS
// ═══════════════════════════════════════════

function computeMonthGroups(timeUnits, unit) {
  var groups = [];
  var months = ["Jan","Feb","Mrz","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez"];
  
  if (timeUnits.length === 0) return groups;
  
  var currentMonth = -1;
  var currentYear = -1;
  var currentCount = 0;
  
  for (var i = 0; i < timeUnits.length; i++) {
    var tu = timeUnits[i];
    var m = tu.monthIndex;
    var y = tu.year;
    
    if (m === currentMonth && y === currentYear) {
      currentCount++;
    } else {
      if (currentCount > 0) {
        groups.push({
          label: months[currentMonth] + " " + currentYear,
          count: currentCount
        });
      }
      currentMonth = m;
      currentYear = y;
      currentCount = 1;
    }
  }
  
  if (currentCount > 0) {
    groups.push({
      label: months[currentMonth] + " " + currentYear,
      count: currentCount
    });
  }
  
  return groups;
}

function computeTimeUnits(start, end, unit) {
  var units = [];
  var totalDays = daysBetween(start, end);
  
  if (unit === "day") {
    for (var i = 0; i < totalDays && i < 200; i++) {
      var d = addDays(start, i);
      units.push({
        label: pad2(d.getDate()),
        days: 1,
        monthIndex: d.getMonth(),
        year: d.getFullYear()
      });
    }
  } 
  else if (unit === "week") {
    var cur = new Date(start);
    while (cur < end) {
      var weekEnd = addDays(cur, 7);
      if (weekEnd > end) weekEnd = new Date(end);
      var days = daysBetween(cur, weekEnd);
      if (days > 0) {
        units.push({
          label: String(getWeekNumber(cur)),
          days: days,
          monthIndex: cur.getMonth(),
          year: cur.getFullYear()
        });
      }
      cur = weekEnd;
    }
  }
  else if (unit === "month") {
    var months = ["Jan","Feb","Mrz","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez"];
    var cur = new Date(start.getFullYear(), start.getMonth(), 1);
    while (cur < end) {
      var mStart = cur < start ? new Date(start) : new Date(cur);
      var mEnd = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
      if (mEnd > end) mEnd = new Date(end);
      var days = daysBetween(mStart, mEnd);
      if (days > 0) {
        units.push({
          label: months[cur.getMonth()],
          days: days,
          monthIndex: cur.getMonth(),
          year: cur.getFullYear()
        });
      }
      cur = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
    }
  }
  else if (unit === "quarter") {
    var cur = new Date(start.getFullYear(), Math.floor(start.getMonth() / 3) * 3, 1);
    while (cur < end) {
      var qStart = cur < start ? new Date(start) : new Date(cur);
      var qEnd = new Date(cur.getFullYear(), cur.getMonth() + 3, 1);
      if (qEnd > end) qEnd = new Date(end);
      var days = daysBetween(qStart, qEnd);
      if (days > 0) {
        var qNum = Math.floor(cur.getMonth() / 3) + 1;
        units.push({
          label: "Q" + qNum,
          days: days,
          monthIndex: cur.getMonth(),
          year: cur.getFullYear()
        });
      }
      cur = new Date(cur.getFullYear(), cur.getMonth() + 3, 1);
    }
  }
  
  return units;
}
