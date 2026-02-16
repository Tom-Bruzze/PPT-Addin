/*
 ═══════════════════════════════════════════════════════
 Droege GANTT Generator  –  taskpane.js  v2.19

 UPDATES v2.19:
  - DEBUG-AUSGABE DIREKT IM PANEL (ganttInfo)
  - Zeigt berechnete Positionen an
  - Alles in einer Funktion

 DROEGE GROUP · 2026
 ═══════════════════════════════════════════════════════
*/

var VERSION = "2.19";

// ═══════════════════════════════════════════
// KONSTANTEN
// ═══════════════════════════════════════════
var POINTS_PER_CM = 28.3464567;
var RE_CM = 0.21;
var RE_PT = RE_CM * POINTS_PER_CM;  // 5.95275590551

var CM = POINTS_PER_CM;
var gridUnitCm = RE_CM;
var ganttPhaseCount = 0;

// GANTT Layout
var GANTT_LEFT_RE = 9;
var GANTT_TOP_RE = 17;
var GANTT_MAX_WIDTH_RE = 118;

var FONT_SIZE = 11;
var LINE_WEIGHT = 0.5;

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
  
  document.querySelectorAll(".pre").forEach(function(b) {
    b.addEventListener("click", function() {
      var v = parseFloat(this.dataset.value);
      gridUnitCm = v;
      document.getElementById("gridUnit").value = v;
      updatePresetButtons(v);
    });
  });
  
  var ss = document.getElementById("setSlide");
  if (ss) ss.addEventListener("click", function() {
    var sel = document.getElementById("slideSize").value;
    setSlideSize(sel);
  });
  
  var btnGantt = document.getElementById("createGantt");
  if (btnGantt) {
    btnGantt.addEventListener("click", generateGantt);
  }
  
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
// GANTT GENERATOR
// ═══════════════════════════════════════════

async function generateGantt() {
  var debugLog = [];
  debugLog.push("GANTT v2.19 gestartet");
  
  showStatus("Generiere GANTT...", "info");
  ganttInfo("Generiere...", false);
  
  // 1. Phasen sammeln
  var phases = getPhases();
  debugLog.push("Phasen: " + phases.length);
  
  if (phases.length === 0) {
    ganttInfo("Keine gültigen Phasen definiert!", true);
    showStatus("Fehler: Keine Phasen", "error");
    return;
  }
  
  // 2. Zeitraum
  var projStart = new Date(document.getElementById("ganttStart").value);
  var projEnd = new Date(document.getElementById("ganttEnd").value);
  
  if (isNaN(projStart.getTime()) || isNaN(projEnd.getTime()) || projEnd <= projStart) {
    ganttInfo("Ungültiger Zeitraum!", true);
    showStatus("Fehler: Zeitraum", "error");
    return;
  }
  
  // 3. Zeiteinheiten
  var unit = document.getElementById("ganttUnit").value;
  var timeUnits = computeTimeUnits(projStart, projEnd, unit);
  debugLog.push("Zeiteinheiten: " + timeUnits.length);
  
  if (timeUnits.length === 0) {
    ganttInfo("Keine Zeiteinheiten berechnet!", true);
    return;
  }
  
  // 4. GANTT zeichnen
  try {
    await PowerPoint.run(async function(context) {
      debugLog.push("PowerPoint.run OK");
      
      // A) FOLIENGRÖSSE
      var presentation = context.presentation;
      presentation.load("slideWidth,slideHeight");
      await context.sync();
      
      var slideWidth = presentation.slideWidth;
      var slideHeight = presentation.slideHeight;
      debugLog.push("Folie: " + slideWidth.toFixed(1) + " x " + slideHeight.toFixed(1) + " pt");
      
      // B) MARGIN
      var fullUnitsX = Math.floor(slideWidth / RE_PT);
      var fullUnitsY = Math.floor(slideHeight / RE_PT);
      var marginLeft = (slideWidth - (fullUnitsX * RE_PT)) / 2;
      var marginTop = (slideHeight - (fullUnitsY * RE_PT)) / 2;
      debugLog.push("Margin: " + marginLeft.toFixed(2) + " x " + marginTop.toFixed(2) + " pt");
      
      // C) GANTT POSITION
      var ganttLeftPt = marginLeft + (GANTT_LEFT_RE * RE_PT);
      var ganttTopPt = marginTop + (GANTT_TOP_RE * RE_PT);
      debugLog.push("GANTT bei: " + ganttLeftPt.toFixed(1) + " x " + ganttTopPt.toFixed(1) + " pt");
      debugLog.push("GANTT bei: " + (ganttLeftPt/CM).toFixed(2) + " x " + (ganttTopPt/CM).toFixed(2) + " cm");
      
      // D) LAYOUT
      var labelWidthRE = 15;
      var headerHeightRE = 3;
      var barHeightRE = 2;
      var rowHeightRE = 3;
      var colWidthRE = 4;
      
      var labelWidthPt = labelWidthRE * RE_PT;
      var headerHeightPt = headerHeightRE * RE_PT;
      var barHeightPt = barHeightRE * RE_PT;
      var rowHeightPt = rowHeightRE * RE_PT;
      var colWidthPt = colWidthRE * RE_PT;
      
      var availableWidthRE = GANTT_MAX_WIDTH_RE - labelWidthRE;
      var visibleColumns = Math.min(timeUnits.length, Math.floor(availableWidthRE / colWidthRE));
      if (visibleColumns < timeUnits.length) {
        timeUnits = timeUnits.slice(0, visibleColumns);
      }
      
      var chartLeft = ganttLeftPt + labelWidthPt;
      var chartWidth = visibleColumns * colWidthPt;
      var totalWidth = labelWidthPt + chartWidth;
      
      var needsMonthRow = (unit === "day" || unit === "week" || unit === "quarter");
      var monthRowHeightPt = needsMonthRow ? headerHeightPt : 0;
      var totalHeaderHeight = monthRowHeightPt + headerHeightPt;
      var chartTop = ganttTopPt + totalHeaderHeight;
      var totalHeight = totalHeaderHeight + (phases.length * rowHeightPt);
      var barPadding = (rowHeightPt - barHeightPt) / 2;
      
      // E) FOLIE
      var slide = presentation.getSelectedSlides().getItemAt(0);
      
      // F) HINTERGRUND
      debugLog.push("Zeichne BG bei L=" + ganttLeftPt.toFixed(1) + ", T=" + ganttTopPt.toFixed(1));
      
      var bg = slide.shapes.addGeometricShape(
        PowerPoint.GeometricShapeType.rectangle,
        {
          left: ganttLeftPt,
          top: ganttTopPt,
          width: totalWidth,
          height: totalHeight
        }
      );
      bg.fill.setSolidColor("FFFFFF");
      bg.lineFormat.color = "808080";
      bg.lineFormat.weight = LINE_WEIGHT;
      
      // G) HEADER
      var headerTop = ganttTopPt + monthRowHeightPt;
      
      for (var c = 0; c < timeUnits.length; c++) {
        var colX = chartLeft + (c * colWidthPt);
        
        var hdr = slide.shapes.addGeometricShape(
          PowerPoint.GeometricShapeType.rectangle,
          {
            left: colX,
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
      }
      
      // H) ZEILEN
      var totalDays = daysBetween(projStart, projEnd);
      
      for (var i = 0; i < phases.length; i++) {
        var p = phases[i];
        var rowTop = chartTop + (i * rowHeightPt);
        
        // Zeilenhintergrund
        var rowBg = slide.shapes.addGeometricShape(
          PowerPoint.GeometricShapeType.rectangle,
          {
            left: ganttLeftPt,
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
            left: ganttLeftPt,
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
        
        // Balken
        var phaseStartDay = Math.max(0, daysBetween(projStart, p.start));
        var phaseEndDay = Math.min(totalDays, daysBetween(projStart, p.end));
        
        if (phaseEndDay > phaseStartDay && totalDays > 0) {
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
      
      // I) HEUTE-LINIE
      var today = new Date();
      if (today >= projStart && today <= projEnd && totalDays > 0) {
        var todayDays = daysBetween(projStart, today);
        var todayPct = todayDays / totalDays;
        var todayX = chartLeft + (todayPct * chartWidth);
        
        var todayLine = slide.shapes.addLine(
          PowerPoint.ConnectorType.straight,
          {
            left: todayX,
            top: ganttTopPt,
            width: 0,
            height: totalHeight
          }
        );
        todayLine.lineFormat.color = "FF0000";
        todayLine.lineFormat.weight = 2;
      }
      
      debugLog.push("Sync...");
      await context.sync();
      debugLog.push("FERTIG!");
    });
    
    // DEBUG-AUSGABE IM PANEL
    var debugHtml = "<strong>DEBUG:</strong><br>" + debugLog.join("<br>");
    ganttInfo(debugHtml, false);
    showStatus("GANTT erstellt", "success");
    
  } catch (e) {
    debugLog.push("FEHLER: " + e.message);
    var debugHtml = "<strong>DEBUG:</strong><br>" + debugLog.join("<br>");
    ganttInfo(debugHtml, true);
    showStatus("Fehler: " + e.message, "error");
  }
}

// ═══════════════════════════════════════════
// COMPUTE FUNCTIONS
// ═══════════════════════════════════════════

function computeTimeUnits(start, end, unit) {
  var units = [];
  var totalDays = daysBetween(start, end);
  
  if (unit === "day") {
    for (var i = 0; i < totalDays && i < 200; i++) {
      var d = addDays(start, i);
      units.push({
        label: pad2(d.getDate()),
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
      units.push({
        label: String(getWeekNumber(cur)),
        monthIndex: cur.getMonth(),
        year: cur.getFullYear()
      });
      cur = weekEnd;
    }
  }
  else if (unit === "month") {
    var months = ["Jan","Feb","Mrz","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez"];
    var cur = new Date(start.getFullYear(), start.getMonth(), 1);
    while (cur < end) {
      units.push({
        label: months[cur.getMonth()],
        monthIndex: cur.getMonth(),
        year: cur.getFullYear()
      });
      cur = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
    }
  }
  else if (unit === "quarter") {
    var cur = new Date(start.getFullYear(), Math.floor(start.getMonth() / 3) * 3, 1);
    while (cur < end) {
      var qNum = Math.floor(cur.getMonth() / 3) + 1;
      units.push({
        label: "Q" + qNum,
        monthIndex: cur.getMonth(),
        year: cur.getFullYear()
      });
      cur = new Date(cur.getFullYear(), cur.getMonth() + 3, 1);
    }
  }
  
  return units;
}
