/*
 ═══════════════════════════════════════════════════════
 Droege GANTT Generator  –  taskpane.js  v2.14
 KORREKTUR: Keine Rundung bei re2pt()
 ═══════════════════════════════════════════════════════
*/

var VERSION = "2.14";
var CM = 28.3465;
var gridUnitCm = 0.21;
var ganttPhaseCount = 0;

var GANTT_LEFT_RE = 9;
var GANTT_TOP_RE = 17;
var GANTT_MAX_WIDTH_RE = 118;
var FONT_SIZE = 11;

Office.onReady(function(info) {
  if (info.host === Office.HostType.PowerPoint) {
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
// RE to Points - OHNE RUNDUNG!
// 3 RE × 0.21 cm × 28.3465 = 17.858295 pt (exakt)
// ═══════════════════════════════════════════
function re2pt(re) {
  return re * gridUnitCm * CM;
}

function setSlideSize() {
  showStatus("Setze Format...", "info");
  PowerPoint.run(function(ctx) {
    ctx.presentation.pageSetup.slideWidth = 786;
    ctx.presentation.pageSetup.slideHeight = 547;
    return ctx.sync();
  }).then(function() {
    showStatus("Format gesetzt: 27.73 x 19.30 cm", "success");
  }).catch(function(err) {
    showStatus("Fehler: " + err.message, "error");
  });
}

function createGanttChart() {
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

  var chartAreaWidthRE = GANTT_MAX_WIDTH_RE - labelWidthRE;
  var visibleColumns = timeUnits.length;
  var truncated = false;
  
  if (widthMode === "auto") {
    colWidthRE = Math.floor(chartAreaWidthRE / timeUnits.length);
    if (colWidthRE < 1) {
      colWidthRE = 1;
      visibleColumns = Math.floor(chartAreaWidthRE / 1);
      truncated = visibleColumns < timeUnits.length;
    }
    if (colWidthRE > 10) colWidthRE = 10;
  } else {
    var requiredWidth = timeUnits.length * colWidthRE;
    if (requiredWidth > chartAreaWidthRE) {
      visibleColumns = Math.floor(chartAreaWidthRE / colWidthRE);
      truncated = true;
    }
  }
  
  var actualChartWidthRE = visibleColumns * colWidthRE;

  var unitNames = {day:"Tage", week:"Wochen", month:"Monate", quarter:"Quartale"};
  var infoText = "<b>" + phases.length + "</b> Phasen, <b>" + visibleColumns + "</b> " + unitNames[unit];
  if (truncated) infoText += " <span style='color:#e94560'>(abgeschnitten)</span>";
  infoText += "<br>Spalte: <b>" + colWidthRE + " RE</b> (" + (colWidthRE * gridUnitCm).toFixed(2) + " cm)";
  ganttInfo(infoText, false);

  showStatus("Erstelle GANTT...", "info");

  var visibleTimeUnits = timeUnits.slice(0, visibleColumns);

  PowerPoint.run(function(ctx) {
    var selectedSlides = ctx.presentation.getSelectedSlides();
    selectedSlides.load("items");
    
    return ctx.sync().then(function() {
      var slide;
      if (selectedSlides.items && selectedSlides.items.length > 0) {
        slide = selectedSlides.items[0];
      } else {
        var allSlides = ctx.presentation.slides;
        allSlides.load("items");
        return ctx.sync().then(function() {
          if (allSlides.items.length > 0) slide = allSlides.items[0];
          return drawGantt(ctx, slide, projStart, projEnd, unit, phases, visibleTimeUnits, 
                          labelWidthRE, headerHeightRE, barHeightRE, rowHeightRE, 
                          colWidthRE, totalDays, showTodayLine, actualChartWidthRE);
        });
      }
      return drawGantt(ctx, slide, projStart, projEnd, unit, phases, visibleTimeUnits, 
                      labelWidthRE, headerHeightRE, barHeightRE, rowHeightRE, 
                      colWidthRE, totalDays, showTodayLine, actualChartWidthRE);
    });
  }).then(function() {
    showStatus("GANTT erstellt!", "success");
  }).catch(function(err) {
    showStatus("Fehler: " + err.message, "error");
  });
}

function drawGantt(ctx, slide, projStart, projEnd, unit, phases, timeUnits, 
                   labelWidthRE, headerHeightRE, barHeightRE, rowHeightRE, 
                   colWidthRE, totalDays, showTodayLine, actualChartWidthRE) {
  
  var GANTT_LEFT_PT = re2pt(GANTT_LEFT_RE);
  var GANTT_TOP_PT = re2pt(GANTT_TOP_RE);
  
  var labelWidthPt = re2pt(labelWidthRE);
  var headerHeightPt = re2pt(headerHeightRE);
  var barHeightPt = re2pt(barHeightRE);
  var rowHeightPt = re2pt(rowHeightRE);
  var colWidthPt = re2pt(colWidthRE);
  var chartWidthPt = re2pt(actualChartWidthRE);
  
  var barPadding = (rowHeightPt - barHeightPt) / 2;
  if (barPadding < 2) barPadding = 2;
  
  var chartLeft = GANTT_LEFT_PT + labelWidthPt;
  var totalWidth = labelWidthPt + chartWidthPt;
  
  var needsMonthRow = (unit === "day" || unit === "week" || unit === "quarter");
  var monthRowHeightPt = needsMonthRow ? headerHeightPt : 0;
  var totalHeaderHeight = monthRowHeightPt + headerHeightPt;
  var chartTop = GANTT_TOP_PT + totalHeaderHeight;
  var totalHeight = totalHeaderHeight + (phases.length * rowHeightPt);
  var chartBottom = GANTT_TOP_PT + totalHeight;
  var lineHeight = chartBottom - chartTop;

  // 1. Hintergrund
  var bg = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle,
    { left: GANTT_LEFT_PT, top: GANTT_TOP_PT, width: totalWidth, height: totalHeight });
  bg.fill.setSolidColor("FFFFFF");
  
  // 2a. Monatszeile
  if (needsMonthRow) {
    var monthGroups = computeMonthGroups(timeUnits, unit);
    var monthX = 0;
    for (var m = 0; m < monthGroups.length; m++) {
      var mg = monthGroups[m];
      var monthWidthPt = mg.count * colWidthPt;
      var mc = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle,
        { left: chartLeft + monthX, top: GANTT_TOP_PT, width: monthWidthPt, height: monthRowHeightPt });
      mc.fill.setSolidColor("B0B0B0");
      try {
        mc.textFrame.textRange.text = mg.label;
        mc.textFrame.textRange.font.size = FONT_SIZE;
        mc.textFrame.textRange.font.bold = true;
        mc.textFrame.textRange.font.color = "000000";
        mc.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
        mc.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
      } catch(e) {}
      monthX += monthWidthPt;
    }
  }
  
  // 2b. Header-Zellen
  var colX = 0;
  var linePositions = [];
  var headerTop = GANTT_TOP_PT + monthRowHeightPt;
  
  for (var c = 0; c < timeUnits.length; c++) {
    var hdr = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle,
      { left: chartLeft + colX, top: headerTop, width: colWidthPt, height: headerHeightPt });
    hdr.fill.setSolidColor("D5D5D5");
    try {
      hdr.textFrame.textRange.text = timeUnits[c].label;
      hdr.textFrame.textRange.font.size = FONT_SIZE;
      hdr.textFrame.textRange.font.bold = true;
      hdr.textFrame.textRange.font.color = "000000";
      hdr.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
      hdr.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
    } catch(e) {}
    if (c < timeUnits.length - 1) linePositions.push(chartLeft + colX + colWidthPt);
    colX += colWidthPt;
  }
  
  // 3. Trennlinien
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
    
    var label = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle,
      { left: GANTT_LEFT_PT, top: rowTop, width: labelWidthPt, height: rowHeightPt });
    label.fill.setSolidColor("F0F0F0");
    try {
      label.textFrame.textRange.text = " " + phase.name;
      label.textFrame.textRange.font.size = FONT_SIZE;
      label.textFrame.textRange.font.bold = true;
      label.textFrame.textRange.font.color = "000000";
      label.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
    } catch(e) {}

    var phaseStartDay = daysBetween(projStart, phase.start);
    var phaseEndDay = daysBetween(projStart, phase.end);
    if (phaseStartDay < 0) phaseStartDay = 0;
    if (phaseEndDay > totalDays) phaseEndDay = totalDays;
    
    if (phaseEndDay > phaseStartDay) {
      var barLeft = chartLeft + (phaseStartDay / totalDays) * chartWidthPt;
      var barWidth = ((phaseEndDay - phaseStartDay) / totalDays) * chartWidthPt;
      var maxRight = chartLeft + chartWidthPt;
      if (barLeft + barWidth > maxRight) barWidth = maxRight - barLeft;
      if (barLeft > maxRight) continue;
      if (barWidth < 10) barWidth = 10;
      var barTop = rowTop + barPadding;
      
      var bar = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle,
        { left: barLeft, top: barTop, width: barWidth, height: barHeightPt });
      bar.fill.setSolidColor(phase.color.replace("#", ""));
      try {
        bar.textFrame.textRange.text = phase.name;
        bar.textFrame.textRange.font.size = FONT_SIZE;
        bar.textFrame.textRange.font.color = "000000";
        bar.textFrame.textRange.font.bold = true;
        bar.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
        bar.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
      } catch(e) {}
    }
  }

  // 5. Heute-Linie
  if (showTodayLine) {
    var today = new Date();
    var todayDay = daysBetween(projStart, today);
    if (todayDay >= 0 && todayDay <= totalDays) {
      var todayX = chartLeft + (todayDay / totalDays) * chartWidthPt;
      if (todayX <= chartLeft + chartWidthPt) {
        var tl = slide.shapes.addLine(PowerPoint.ConnectorType.straight,
          { left: todayX, top: GANTT_TOP_PT, width: 0.01, height: totalHeight });
        tl.lineFormat.color = "FF0000";
        tl.lineFormat.weight = 2;
        try {
          var tlb = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle,
            { left: todayX - 20, top: GANTT_TOP_PT + totalHeight + 4, width: 40, height: 16 });
          tlb.fill.setSolidColor("FF0000");
          tlb.textFrame.textRange.text = "Heute";
          tlb.textFrame.textRange.font.size = 9;
          tlb.textFrame.textRange.font.color = "FFFFFF";
          tlb.textFrame.textRange.font.bold = true;
          tlb.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
          tlb.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
        } catch(e) {}
      }
    }
  }
  
  return ctx.sync();
}

function computeMonthGroups(timeUnits, unit) {
  var groups = [];
  var months = ["Jan","Feb","Mrz","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez"];
  if (timeUnits.length === 0) return groups;
  var currentMonth = -1, currentYear = -1, currentCount = 0;
  for (var i = 0; i < timeUnits.length; i++) {
    var m = timeUnits[i].monthIndex, y = timeUnits[i].year;
    if (m === currentMonth && y === currentYear) {
      currentCount++;
    } else {
      if (currentCount > 0) groups.push({ label: months[currentMonth] + " " + currentYear, count: currentCount });
      currentMonth = m; currentYear = y; currentCount = 1;
    }
  }
  if (currentCount > 0) groups.push({ label: months[currentMonth] + " " + currentYear, count: currentCount });
  return groups;
}

function computeTimeUnits(start, end, unit) {
  var units = [];
  var totalDays = daysBetween(start, end);
  
  if (unit === "day") {
    for (var i = 0; i < totalDays && i < 200; i++) {
      var d = addDays(start, i);
      units.push({ label: pad2(d.getDate()), days: 1, monthIndex: d.getMonth(), year: d.getFullYear() });
    }
  } else if (unit === "week") {
    var cur = new Date(start);
    while (cur < end) {
      var weekEnd = addDays(cur, 7);
      if (weekEnd > end) weekEnd = new Date(end);
      var days = daysBetween(cur, weekEnd);
      if (days > 0) units.push({ label: String(getWeekNumber(cur)), days: days, monthIndex: cur.getMonth(), year: cur.getFullYear() });
      cur = weekEnd;
    }
  } else if (unit === "month") {
    var months = ["Jan","Feb","Mrz","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez"];
    var cur = new Date(start.getFullYear(), start.getMonth(), 1);
    while (cur < end) {
      var mStart = cur < start ? new Date(start) : new Date(cur);
      var mEnd = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
      if (mEnd > end) mEnd = new Date(end);
      var days = daysBetween(mStart, mEnd);
      if (days > 0) units.push({ label: months[cur.getMonth()], days: days, monthIndex: cur.getMonth(), year: cur.getFullYear() });
      cur = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
    }
  } else if (unit === "quarter") {
    var cur = new Date(start.getFullYear(), Math.floor(start.getMonth() / 3) * 3, 1);
    while (cur < end) {
      var qStart = cur < start ? new Date(start) : new Date(cur);
      var qEnd = new Date(cur.getFullYear(), cur.getMonth() + 3, 1);
      if (qEnd > end) qEnd = new Date(end);
      var days = daysBetween(qStart, qEnd);
      var q = Math.floor(cur.getMonth() / 3) + 1;
      if (days > 0) units.push({ label: "Q" + q, days: days, monthIndex: cur.getMonth(), year: cur.getFullYear() });
      cur = new Date(cur.getFullYear(), cur.getMonth() + 3, 1);
    }
  }
  return units;
}
