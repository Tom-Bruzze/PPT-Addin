/*
 ═══════════════════════════════════════════════════════
 Droege GANTT Generator  –  taskpane.js  v2.12

 UPDATES v2.12:
  - Spaltenbreite als Dropdown auswählbar
  - Monatszeile ohne Jahreszahl
  - Meilenstein-Dreieck zeigt nach oben (0°)

 DROEGE GROUP · 2026
 ═══════════════════════════════════════════════════════
*/

var VERSION = "2.12";
var CM = 28.3465;
var gridUnitCm = 0.21;
var ganttPhaseCount = 0;

// GANTT Layout in Points
var GANTT_LEFT_PT = 48;
var GANTT_TOP_PT = 100;
var GANTT_WIDTH_PT = 700;

// Schriftgröße für alle Texte
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
      if (gi) gi.value = v;
      updatePresetButtons(v);
    });
  });

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
// HELPER FUNCTIONS
// ═══════════════════════════════════════════

function pad2(n) { return n < 10 ? "0" + n : String(n); }

function toISO(d) { 
  return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate()); 
}

function addDays(d, n) { 
  var r = new Date(d); 
  r.setDate(r.getDate() + n); 
  return r; 
}

function daysBetween(a, b) { 
  return Math.round((b.getTime() - a.getTime()) / 86400000); 
}

function escHtml(s) { 
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); 
}

function randomColor() {
  var colors = ["#2e86c1", "#27ae60", "#e94560", "#f39c12", "#8e44ad", "#1abc9c"];
  return colors[Math.floor(Math.random() * colors.length)];
}

function showStatus(msg, type) {
  var el = document.getElementById("status");
  if (el) {
    el.textContent = msg;
    el.className = "sts " + (type || "info");
  }
}

function ganttInfo(msg, isError) {
  var el = document.getElementById("ganttInfo");
  if (el) {
    el.innerHTML = msg;
    el.className = "gantt-info" + (isError ? " err" : "");
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

// RE to Points
function re2pt(re) {
  return Math.round(re * gridUnitCm * CM);
}

// ═══════════════════════════════════════════
// SLIDE SIZE
// ═══════════════════════════════════════════

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

// ═══════════════════════════════════════════
// GANTT CHART CREATION
// ═══════════════════════════════════════════

function createGanttChart() {
  console.log("=== createGanttChart START v2.12 ===");
  
  // Eingaben lesen
  var projStart = new Date(document.getElementById("ganttStart").value);
  var projEnd = new Date(document.getElementById("ganttEnd").value);
  var unit = document.getElementById("ganttUnit").value;
  var labelWidthRE = parseInt(document.getElementById("ganttLabelW").value) || 20;
  var headerHeightRE = parseInt(document.getElementById("ganttHeaderH").value) || 3;
  var barHeightRE = parseInt(document.getElementById("ganttBarH").value) || 3;
  var rowHeightRE = parseInt(document.getElementById("ganttRowH").value) || 5;
  var showTodayLine = document.getElementById("ganttTodayLine").checked;
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

  // Zeiteinheiten berechnen - jetzt mit fester Breite in RE
  var timeUnits = computeTimeUnits(projStart, projEnd, unit);
  console.log("Zeiteinheiten:", timeUnits.length);
  
  if (timeUnits.length === 0 || timeUnits.length > 100) {
    ganttInfo("Zeitraum anpassen oder andere Einheit wählen!", true);
    return;
  }

  var totalDays = daysBetween(projStart, projEnd);
  if (totalDays < 1) totalDays = 1;

  // Info anzeigen
  var unitNames = {day:"Tage", week:"Wochen", month:"Monate", quarter:"Quartale"};
  ganttInfo("<b>" + phases.length + "</b> Phasen, <b>" + timeUnits.length + "</b> " + unitNames[unit], false);

  showStatus("Erstelle GANTT auf aktueller Folie...", "info");

  // GANTT zeichnen
  PowerPoint.run(function(ctx) {
    console.log("PowerPoint.run gestartet");
    
    var selectedSlides = ctx.presentation.getSelectedSlides();
    selectedSlides.load("items");
    
    return ctx.sync().then(function() {
      var slide;
      
      if (selectedSlides.items && selectedSlides.items.length > 0) {
        slide = selectedSlides.items[0];
        console.log("Verwende ausgewählte Folie");
      } else {
        var allSlides = ctx.presentation.slides;
        allSlides.load("items");
        return ctx.sync().then(function() {
          if (allSlides.items.length === 0) {
            throw new Error("Keine Folie vorhanden");
          }
          return drawGanttOnSlide(ctx, allSlides.items[0], projStart, projEnd, totalDays, 
                                   timeUnits, phases, labelWidthRE, headerHeightRE, barHeightRE, rowHeightRE, showTodayLine, colWidthRE, unit);
        });
      }
      
      return drawGanttOnSlide(ctx, slide, projStart, projEnd, totalDays, 
                               timeUnits, phases, labelWidthRE, headerHeightRE, barHeightRE, rowHeightRE, showTodayLine, colWidthRE, unit);
    });
    
  }).catch(function(err) {
    console.error("FEHLER:", err);
    showStatus("Fehler: " + err.message, "error");
  });
}

function drawGanttOnSlide(ctx, slide, projStart, projEnd, totalDays, timeUnits, phases, labelWidthRE, headerHeightRE, barHeightRE, rowHeightRE, showTodayLine, colWidthRE, unit) {
  console.log("drawGanttOnSlide gestartet");
  
  // Layout berechnen in Points
  var labelWidthPt = re2pt(labelWidthRE);
  var headerHeightPt = re2pt(headerHeightRE);
  var barHeightPt = re2pt(barHeightRE);
  var rowHeightPt = re2pt(rowHeightRE);
  var colWidthPt = re2pt(colWidthRE);  // Feste Spaltenbreite in RE
  
  // Balken-Padding (zentriert in Zeile)
  var barPadding = Math.max(2, Math.round((rowHeightPt - barHeightPt) / 2));
  
  var chartLeft = GANTT_LEFT_PT + labelWidthPt;
  
  // Chart-Breite basiert jetzt auf Anzahl Zeiteinheiten × Spaltenbreite
  var chartWidth = timeUnits.length * colWidthPt;
  var totalWidth = labelWidthPt + chartWidth;
  
  // Prüfen ob Monatszeile benötigt wird (bei Tage, Wochen, Quartale)
  var needsMonthRow = (unit === "day" || unit === "week" || unit === "quarter");
  var monthRowHeightPt = needsMonthRow ? headerHeightPt : 0;
  
  // Header-Bereich (Monatszeile + Hauptheader)
  var totalHeaderHeight = monthRowHeightPt + headerHeightPt;
  var chartTop = GANTT_TOP_PT + totalHeaderHeight;
  
  // Gesamthöhe berechnen
  var totalHeight = totalHeaderHeight + (phases.length * rowHeightPt);
  var chartBottom = GANTT_TOP_PT + totalHeight;
  
  // Höhe der Linien (vom Hauptheader bis zum Ende)
  var lineHeight = chartBottom - (GANTT_TOP_PT + monthRowHeightPt + headerHeightPt);
  
  console.log("Layout:", {
    labelWidthPt: labelWidthPt,
    headerHeightPt: headerHeightPt,
    monthRowHeightPt: monthRowHeightPt,
    barHeightPt: barHeightPt,
    rowHeightPt: rowHeightPt,
    colWidthPt: colWidthPt,
    chartLeft: chartLeft,
    chartWidth: chartWidth,
    totalHeight: totalHeight,
    needsMonthRow: needsMonthRow
  });

  // ═══ 1. HINTERGRUND ═══
  console.log("1. Hintergrund");
  var bg = slide.shapes.addGeometricShape(
    PowerPoint.GeometricShapeType.rectangle,
    {
      left: Math.round(GANTT_LEFT_PT),
      top: Math.round(GANTT_TOP_PT),
      width: Math.round(totalWidth),
      height: Math.round(totalHeight)
    }
  );
  bg.fill.setSolidColor("FFFFFF");
  
  // ═══ 2. MONATSZEILE (wenn benötigt) - OHNE JAHRESZAHL ═══
  if (needsMonthRow) {
    console.log("2a. Monatszeile (ohne Jahreszahl)");
    var monthGroups = computeMonthGroups(timeUnits, unit);
    var monthX = 0;
    
    for (var m = 0; m < monthGroups.length; m++) {
      var mg = monthGroups[m];
      var monthWidth = mg.count * colWidthPt;
      
      var monthCell = slide.shapes.addGeometricShape(
        PowerPoint.GeometricShapeType.rectangle,
        {
          left: Math.round(chartLeft + monthX),
          top: Math.round(GANTT_TOP_PT),
          width: Math.round(monthWidth),
          height: Math.round(monthRowHeightPt)
        }
      );
      monthCell.fill.setSolidColor("B0B0B0");
      
      try {
        monthCell.textFrame.textRange.text = mg.label;  // Nur Monatsname, keine Jahreszahl
        monthCell.textFrame.textRange.font.size = FONT_SIZE;
        monthCell.textFrame.textRange.font.bold = true;
        monthCell.textFrame.textRange.font.color = "000000";
        monthCell.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
        monthCell.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
      } catch(e) {}
      
      monthX += monthWidth;
    }
  }
  
  // ═══ 2b. HEADER-ZELLEN (Zeiteinheiten) ═══
  console.log("2b. Header-Zellen: " + timeUnits.length);
  var colX = 0;
  var linePositions = [];
  var headerTop = GANTT_TOP_PT + monthRowHeightPt;
  
  for (var c = 0; c < timeUnits.length; c++) {
    // Header-Zelle mit fester Breite
    var hdr = slide.shapes.addGeometricShape(
      PowerPoint.GeometricShapeType.rectangle,
      {
        left: Math.round(chartLeft + colX),
        top: Math.round(headerTop),
        width: Math.round(colWidthPt),
        height: Math.round(headerHeightPt)
      }
    );
    hdr.fill.setSolidColor("D5D5D5");
    
    try {
      hdr.textFrame.textRange.text = timeUnits[c].label;
      hdr.textFrame.textRange.font.size = FONT_SIZE;
      hdr.textFrame.textRange.font.bold = true;
      hdr.textFrame.textRange.font.color = "000000";
      hdr.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
      hdr.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
    } catch(e) {}
    
    // Position für Trennlinie speichern (außer nach letzter Spalte)
    if (c < timeUnits.length - 1) {
      linePositions.push(chartLeft + colX + colWidthPt);
    }
    
    colX += colWidthPt;
  }
  
  // ═══ 3. VERTIKALE TRENNLINIEN ═══
  console.log("3. Vertikale Trennlinien: " + linePositions.length);
  
  for (var i = 0; i < linePositions.length; i++) {
    var lineX = linePositions[i];
    
    var line = slide.shapes.addLine(
      PowerPoint.ConnectorType.straight,
      {
        left: lineX,
        top: chartTop,
        width: 0.01,
        height: lineHeight
      }
    );
    
    line.lineFormat.color = "CCCCCC";
    line.lineFormat.weight = 0.75;
  }

  // ═══ 4. PHASEN-ZEILEN, BALKEN UND MEILENSTEINE ═══
  console.log("4. Phasen: " + phases.length);
  
  for (var p = 0; p < phases.length; p++) {
    var phase = phases[p];
    var rowTop = chartTop + (p * rowHeightPt);
    
    console.log("Phase " + p + ": " + phase.name + " | rowTop=" + rowTop);
    
    // ─── Label-Zelle ───
    var label = slide.shapes.addGeometricShape(
      PowerPoint.GeometricShapeType.rectangle,
      {
        left: Math.round(GANTT_LEFT_PT),
        top: Math.round(rowTop),
        width: Math.round(labelWidthPt),
        height: Math.round(rowHeightPt)
      }
    );
    label.fill.setSolidColor("F0F0F0");
    
    try {
      label.textFrame.textRange.text = " " + phase.name;
      label.textFrame.textRange.font.size = FONT_SIZE;
      label.textFrame.textRange.font.bold = true;
      label.textFrame.textRange.font.color = "000000";
      label.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
    } catch(e) {}

    // ─── Balken berechnen ───
    var phaseStartDay = daysBetween(projStart, phase.start);
    var phaseEndDay = daysBetween(projStart, phase.end);
    
    if (phaseStartDay < 0) phaseStartDay = 0;
    if (phaseEndDay > totalDays) phaseEndDay = totalDays;
    
    console.log("  Tage: " + phaseStartDay + " - " + phaseEndDay);
    
    if (phaseEndDay > phaseStartDay) {
      // Berechne Position basierend auf fester Spaltenbreite
      var barLeft = chartLeft + (phaseStartDay / totalDays) * chartWidth;
      var barWidth = ((phaseEndDay - phaseStartDay) / totalDays) * chartWidth;
      
      if (barWidth < 10) barWidth = 10;
      
      var barTop = rowTop + barPadding;
      
      console.log("  Balken: left=" + Math.round(barLeft) + " top=" + Math.round(barTop) + 
                  " w=" + Math.round(barWidth) + " h=" + barHeightPt);
      
      var bar = slide.shapes.addGeometricShape(
        PowerPoint.GeometricShapeType.rectangle,
        {
          left: Math.round(barLeft),
          top: Math.round(barTop),
          width: Math.round(barWidth),
          height: Math.round(barHeightPt)
        }
      );
      
      var colorHex = phase.color.replace("#", "");
      bar.fill.setSolidColor(colorHex);
      
      try {
        bar.textFrame.textRange.text = phase.name;
        bar.textFrame.textRange.font.size = FONT_SIZE;
        bar.textFrame.textRange.font.color = "000000";
        bar.textFrame.textRange.font.bold = true;
        bar.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
        bar.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
      } catch(e) {}
      
      // ─── MEILENSTEIN-DREIECK am Balkenende (zeigt nach OBEN, 0°) ───
      var triangleSize = barHeightPt;  // Höhe und Breite = Balkenhöhe
      // Dreieck zentriert am Ende des Balkens, zeigt nach oben
      var triangleLeft = barLeft + barWidth - (triangleSize / 2);
      var triangleTop = barTop;
      
      var milestone = slide.shapes.addGeometricShape(
        PowerPoint.GeometricShapeType.isoscelesTriangle,
        {
          left: Math.round(triangleLeft),
          top: Math.round(triangleTop),
          width: Math.round(triangleSize),
          height: Math.round(triangleSize)
        }
      );
      milestone.fill.setSolidColor("4A4A4A");  // Dunkelgrau
      // Keine Rotation - Dreieck zeigt standardmäßig nach oben (0°)
      
      console.log("  Meilenstein (nach oben): left=" + Math.round(triangleLeft) + " size=" + Math.round(triangleSize));
    }
  }

  // ═══ 5. HEUTE-LINIE (rot) - Label UNTEN ═══
  if (showTodayLine) {
    var today = new Date();
    var todayDay = daysBetween(projStart, today);
    
    console.log("5. Heute-Linie: Tag " + todayDay + " von " + totalDays);
    
    if (todayDay >= 0 && todayDay <= totalDays) {
      var todayX = chartLeft + (todayDay / totalDays) * chartWidth;
      
      // Rote vertikale Linie
      var todayLine = slide.shapes.addLine(
        PowerPoint.ConnectorType.straight,
        {
          left: todayX,
          top: GANTT_TOP_PT,
          width: 0.01,
          height: totalHeight
        }
      );
      todayLine.lineFormat.color = "FF0000";
      todayLine.lineFormat.weight = 2;
      
      // "Heute" Label UNTEN
      try {
        var todayLabel = slide.shapes.addGeometricShape(
          PowerPoint.GeometricShapeType.rectangle,
          {
            left: Math.round(todayX - 20),
            top: Math.round(chartBottom + 2),
            width: 40,
            height: 14
          }
        );
        todayLabel.fill.setSolidColor("FF0000");
        todayLabel.textFrame.textRange.text = "Heute";
        todayLabel.textFrame.textRange.font.size = 8;
        todayLabel.textFrame.textRange.font.color = "FFFFFF";
        todayLabel.textFrame.textRange.font.bold = true;
        todayLabel.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
        todayLabel.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
      } catch(e) { console.log("Heute-Label Fehler:", e); }
      
      console.log("  Heute-Linie bei X=" + Math.round(todayX));
    }
  }

  console.log("6. Sync...");
  return ctx.sync().then(function() {
    console.log("=== FERTIG ===");
    showStatus("GANTT erstellt (" + phases.length + " Phasen) ✓", "success");
  });
}

// Berechne Monatsgruppen für die obere Zeile - OHNE JAHRESZAHL
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
        // NUR Monatsname, KEINE Jahreszahl
        groups.push({
          label: months[currentMonth],
          count: currentCount
        });
      }
      currentMonth = m;
      currentYear = y;
      currentCount = 1;
    }
  }
  
  // Letzte Gruppe hinzufügen
  if (currentCount > 0) {
    groups.push({
      label: months[currentMonth],
      count: currentCount
    });
  }
  
  return groups;
}

function computeTimeUnits(start, end, unit) {
  var units = [];
  var totalDays = daysBetween(start, end);
  
  if (unit === "day") {
    for (var i = 0; i < totalDays && i < 60; i++) {
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
      var q = Math.floor(cur.getMonth() / 3) + 1;
      if (days > 0) {
        units.push({
          label: "Q" + q,
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
