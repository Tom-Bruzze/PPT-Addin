/*
 ═══════════════════════════════════════════════════════
 Droege GANTT Generator  –  taskpane.js  v2.14

 UPDATES v2.14 (basierend auf v2.11):
  - Meilenstein-Dreieck ENTFERNT
  - Feste Positionierung: Links=9 RE, Oben=17 RE
  - Max. Breite: 118 RE (Gesamtbreite inkl. Label-Spalte)
  - Zwei Spaltenbreiten-Modi:
    1. Fest: Manuell wählbare Spaltenbreite in RE
    2. Auto: Berechnet größtmögliche ganze RE-Breite pro Spalte
       innerhalb der verfügbaren 118 RE

 DROEGE GROUP · 2026
 ═══════════════════════════════════════════════════════
*/

var VERSION = "2.14";
var CM = 28.3465;
var gridUnitCm = 0.21;
var ganttPhaseCount = 0;

// GANTT Layout - FEST in Rastereinheiten
var GANTT_LEFT_RE = 9;         // Links: 9 RE
var GANTT_TOP_RE = 17;         // Oben: 17 RE
var GANTT_MAX_WIDTH_RE = 118;  // Max. Gesamtbreite: 118 RE

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
  console.log("=== createGanttChart START v2.14 ===");
  
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

  // ═══════════════════════════════════════════════════════════
  // SPALTENBREITE BERECHNEN
  // ═══════════════════════════════════════════════════════════
  
  // Verfügbare Breite für den Chart-Bereich (ohne Label-Spalte)
  var chartAreaWidthRE = GANTT_MAX_WIDTH_RE - labelWidthRE;
  
  var visibleColumns = timeUnits.length;
  var truncated = false;
  
  if (widthMode === "auto") {
    // ═══ AUTO-MODUS ═══
    // Berechne die größtmögliche GANZE Rastereinheit pro Spalte,
    // sodass alle Spalten in die verfügbare Breite passen
    
    colWidthRE = Math.floor(chartAreaWidthRE / timeUnits.length);
    
    // Minimum 1 RE, Maximum 10 RE
    if (colWidthRE < 1) {
      colWidthRE = 1;
      // Bei 1 RE: Berechne wie viele Spalten passen
      visibleColumns = Math.floor(chartAreaWidthRE / 1);
      truncated = visibleColumns < timeUnits.length;
    }
    if (colWidthRE > 10) {
      colWidthRE = 10;
    }
    
    console.log("Auto-Modus: " + timeUnits.length + " Zeiteinheiten");
    console.log("Verfügbare Breite: " + chartAreaWidthRE + " RE");
    console.log("Berechnete Spaltenbreite: " + colWidthRE + " RE");
    
  } else {
    // ═══ FESTER MODUS ═══
    // Prüfen ob alle Spalten mit der festen Breite passen
    var requiredWidth = timeUnits.length * colWidthRE;
    
    if (requiredWidth > chartAreaWidthRE) {
      // Abschneiden: Berechne wie viele Spalten passen
      visibleColumns = Math.floor(chartAreaWidthRE / colWidthRE);
      truncated = true;
    }
    
    console.log("Fester Modus: " + colWidthRE + " RE pro Spalte");
    console.log("Benötigte Breite: " + requiredWidth + " RE");
    console.log("Verfügbare Breite: " + chartAreaWidthRE + " RE");
  }
  
  // Tatsächlich genutzte Chart-Breite (in ganzen RE)
  var actualChartWidthRE = visibleColumns * colWidthRE;
  
  console.log("Sichtbare Spalten: " + visibleColumns + " von " + timeUnits.length);
  console.log("Tatsächliche Chart-Breite: " + actualChartWidthRE + " RE");

  // Info anzeigen
  var unitNames = {day:"Tage", week:"Wochen", month:"Monate", quarter:"Quartale"};
  var infoText = "<b>" + phases.length + "</b> Phasen, <b>" + visibleColumns + "</b> " + unitNames[unit];
  if (truncated) {
    infoText += " <span style='color:#e94560'>(von " + timeUnits.length + " – abgeschnitten)</span>";
  }
  infoText += "<br>Spaltenbreite: <b>" + colWidthRE + " RE</b>";
  infoText += " | Chart-Breite: <b>" + actualChartWidthRE + " RE</b>";
  infoText += " | Gesamt: <b>" + (labelWidthRE + actualChartWidthRE) + " RE</b>";
  ganttInfo(infoText, false);

  showStatus("Erstelle GANTT auf aktueller Folie...", "info");

  // Sichtbare Zeiteinheiten (abgeschnitten falls nötig)
  var visibleTimeUnits = timeUnits.slice(0, visibleColumns);

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
          if (allSlides.items.length > 0) {
            slide = allSlides.items[0];
          }
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
    var msg = "GANTT erstellt! " + visibleColumns + " Spalten à " + colWidthRE + " RE = " + actualChartWidthRE + " RE";
    if (truncated) msg += " (abgeschnitten)";
    showStatus(msg, "success");
  }).catch(function(err) {
    console.error("Fehler:", err);
    showStatus("Fehler: " + err.message, "error");
  });
}

function drawGantt(ctx, slide, projStart, projEnd, unit, phases, timeUnits, 
                   labelWidthRE, headerHeightRE, barHeightRE, rowHeightRE, 
                   colWidthRE, totalDays, showTodayLine, actualChartWidthRE) {
  
  // ═══ FESTE POSITIONIERUNG IN RE → POINTS ═══
  var GANTT_LEFT_PT = re2pt(GANTT_LEFT_RE);
  var GANTT_TOP_PT = re2pt(GANTT_TOP_RE);
  
  // Dimensionen in Points
  var labelWidthPt = re2pt(labelWidthRE);
  var headerHeightPt = re2pt(headerHeightRE);
  var barHeightPt = re2pt(barHeightRE);
  var rowHeightPt = re2pt(rowHeightRE);
  var colWidthPt = re2pt(colWidthRE);
  
  // Chart-Breite basiert auf tatsächlicher Breite in RE
  var chartWidthPt = re2pt(actualChartWidthRE);
  
  // Balken-Padding (zentriert in Zeile)
  var barPadding = Math.max(2, Math.round((rowHeightPt - barHeightPt) / 2));
  
  var chartLeft = GANTT_LEFT_PT + labelWidthPt;
  var totalWidth = labelWidthPt + chartWidthPt;
  
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
    GANTT_LEFT_PT: GANTT_LEFT_PT,
    GANTT_TOP_PT: GANTT_TOP_PT,
    labelWidthPt: labelWidthPt,
    colWidthPt: colWidthPt,
    colWidthRE: colWidthRE,
    chartWidthPt: chartWidthPt,
    actualChartWidthRE: actualChartWidthRE,
    visibleColumns: timeUnits.length
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
  
  // ═══ 2. MONATSZEILE (wenn benötigt) ═══
  if (needsMonthRow) {
    console.log("2a. Monatszeile");
    var monthGroups = computeMonthGroups(timeUnits, unit);
    var monthX = 0;
    
    for (var m = 0; m < monthGroups.length; m++) {
      var mg = monthGroups[m];
      var monthWidthPt = mg.count * colWidthPt;
      
      var monthCell = slide.shapes.addGeometricShape(
        PowerPoint.GeometricShapeType.rectangle,
        {
          left: Math.round(chartLeft + monthX),
          top: Math.round(GANTT_TOP_PT),
          width: Math.round(monthWidthPt),
          height: Math.round(monthRowHeightPt)
        }
      );
      monthCell.fill.setSolidColor("B0B0B0");
      
      try {
        monthCell.textFrame.textRange.text = mg.label;
        monthCell.textFrame.textRange.font.size = FONT_SIZE;
        monthCell.textFrame.textRange.font.bold = true;
        monthCell.textFrame.textRange.font.color = "000000";
        monthCell.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
        monthCell.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
      } catch(e) {}
      
      monthX += monthWidthPt;
    }
  }
  
  // ═══ 2b. HEADER-ZELLEN (Zeiteinheiten) ═══
  console.log("2b. Header-Zellen: " + timeUnits.length);
  var colX = 0;
  var linePositions = [];
  var headerTop = GANTT_TOP_PT + monthRowHeightPt;
  
  for (var c = 0; c < timeUnits.length; c++) {
    // Header-Zelle mit fester Breite in RE
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
  console.log("3. Trennlinien: " + linePositions.length);
  for (var li = 0; li < linePositions.length; li++) {
    var line = slide.shapes.addLine(
      PowerPoint.ConnectorType.straight,
      {
        left: linePositions[li],
        top: GANTT_TOP_PT + monthRowHeightPt + headerHeightPt,
        width: 0.01,
        height: lineHeight
      }
    );
    line.lineFormat.color = "AAAAAA";
    line.lineFormat.weight = 0.5;
  }

  // ═══ 4. PHASEN (Zeilen + Balken) - OHNE MEILENSTEIN ═══
  console.log("4. Phasen: " + phases.length);
  for (var p = 0; p < phases.length; p++) {
    var phase = phases[p];
    var rowTop = chartTop + (p * rowHeightPt);
    
    console.log("Phase " + p + ": " + phase.name);
    
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
      var barLeft = chartLeft + (phaseStartDay / totalDays) * chartWidthPt;
      var barWidth = ((phaseEndDay - phaseStartDay) / totalDays) * chartWidthPt;
      
      // Begrenzen auf sichtbaren Bereich
      var maxRight = chartLeft + chartWidthPt;
      if (barLeft + barWidth > maxRight) {
        barWidth = maxRight - barLeft;
      }
      if (barLeft > maxRight) continue;  // Balken außerhalb des sichtbaren Bereichs
      
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
      
      // MEILENSTEIN-DREIECK ENTFERNT (v2.14)
    }
  }

  // ═══ 5. HEUTE-LINIE (rot) - Label UNTEN ═══
  if (showTodayLine) {
    var today = new Date();
    var todayDay = daysBetween(projStart, today);
    
    console.log("5. Heute-Linie: Tag " + todayDay + " von " + totalDays);
    
    if (todayDay >= 0 && todayDay <= totalDays) {
      var todayX = chartLeft + (todayDay / totalDays) * chartWidthPt;
      
      // Prüfen ob im sichtbaren Bereich
      if (todayX <= chartLeft + chartWidthPt) {
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
              top: Math.round(GANTT_TOP_PT + totalHeight + 4),
              width: 40,
              height: 16
            }
          );
          todayLabel.fill.setSolidColor("FF0000");
          todayLabel.textFrame.textRange.text = "Heute";
          todayLabel.textFrame.textRange.font.size = 9;
          todayLabel.textFrame.textRange.font.color = "FFFFFF";
          todayLabel.textFrame.textRange.font.bold = true;
          todayLabel.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
          todayLabel.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
        } catch(e) {}
      }
    }
  }
  
  return ctx.sync();
}

// ═══════════════════════════════════════════
// COMPUTE FUNCTIONS
// ═══════════════════════════════════════════

// Berechnet die Monatsgruppen für die obere Zeile
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
  
  // Letzte Gruppe hinzufügen
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
