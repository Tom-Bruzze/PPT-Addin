/* 
 ═══════════════════════════════════════════════════════
 Droege GANTT Generator  –  taskpane.js  v2.0
 
 Standalone GANTT-Diagramm Add-in für PowerPoint
 Erstellt rasterbasierte GANTT-Charts direkt auf der Folie.
 
 DROEGE GROUP · 2025
 ═══════════════════════════════════════════════════════
 */

/* ═══ Konstanten ═══ */
var CM = 28.3465;                     /* 1 cm in PowerPoint-Punkten   */
var gridUnitCm = 0.21;               /* Standard-Rastereinheit       */
var apiOk = false;

var GANTT_MAX_W = 118;               /* max Breite in RE             */
var GANTT_MAX_H = 69;                /* max Höhe in RE               */
var GANTT_LEFT  = 8;                 /* Abstand links in RE          */
var GANTT_TOP   = 17;                /* Abstand oben in RE           */
var GANTT_TAG   = "DROEGE_GANTT";    /* Shape-Name-Prefix            */
var ganttPhaseCount = 0;

/* ═══════════════════════════════════════════
   Office Init
   ═══════════════════════════════════════════ */
Office.onReady(function (info) {
  if (info.host === Office.HostType.PowerPoint) {
    if (Office.context.requirements && Office.context.requirements.isSetSupported) {
      apiOk = Office.context.requirements.isSetSupported("PowerPointApi", "1.5");
    } else {
      apiOk = (typeof PowerPoint !== "undefined" && PowerPoint.run && typeof PowerPoint.run === "function");
    }
    initUI();
    if (!apiOk) showStatus("PowerPointApi 1.5 nicht verfügbar", "warning");
  }
});

/* ═══════════════════════════════════════════
   UI INIT
   ═══════════════════════════════════════════ */
function initUI() {

  /* Rastereinheit: Input */
  var gi = document.getElementById("gridUnit");
  gi.addEventListener("change", function () {
    var v = parseFloat(this.value);
    if (!isNaN(v) && v > 0) {
      gridUnitCm = v;
      hlPre(v);
      showStatus("RE: " + v.toFixed(2) + " cm", "info");
    }
  });

  /* Rastereinheit: Presets */
  document.querySelectorAll(".pre").forEach(function (b) {
    b.addEventListener("click", function () {
      var v = parseFloat(this.dataset.value);
      gridUnitCm = v;
      gi.value = v;
      hlPre(v);
      showStatus("RE: " + v.toFixed(2) + " cm", "info");
    });
  });

  /* Papierformat */
  bind("setSlide", function () { setSlideSize(); });

  /* GANTT Init */
  initGantt();
}

/* ═══════════════════════════════════════════
   HILFSFUNKTIONEN
   ═══════════════════════════════════════════ */
function bind(id, fn) {
  var el = document.getElementById(id);
  if (!el) return;
  el.addEventListener("click", fn);
}

function hlPre(val) {
  document.querySelectorAll(".pre").forEach(function (b) {
    b.classList.toggle("active", Math.abs(parseFloat(b.dataset.value) - val) < 0.001);
  });
}

function showStatus(msg, type) {
  var el = document.getElementById("status");
  el.textContent = msg;
  el.className = "sts " + (type || "info");
}

function c2p(cm) { return cm * CM; }
function p2c(pt) { return pt / CM; }

/* ═══════════════════════════════════════════
   PAPIERFORMAT 27,728 × 19,297 cm
   ═══════════════════════════════════════════ */
function setSlideSize() {
  if (!apiOk) { showStatus("PowerPointApi 1.5 nötig", "error"); return; }
  PowerPoint.run(function (ctx) {
    var ps = ctx.presentation.pageSetup;
    ps.load(["slideWidth", "slideHeight"]);
    return ctx.sync().then(function () {
      ps.slideWidth = 786;
      return ctx.sync();
    }).then(function () {
      ps.slideHeight = 547;
      return ctx.sync();
    }).then(function () {
      showStatus("Format: 27,728 × 19,297 cm ✓", "success");
    });
  }).catch(function (e) { showStatus("Fehler: " + e.message, "error"); });
}

/* ═══════════════════════════════════════════════════════
   ██████   ██████   ██   █   ████████  ████████
   █        █    █   ██   █      █         █
   █  ███   ██████   █ █  █      █         █
   █    █   █    █   █  █ █      █         █
   ██████   █    █   █   ██      █         █

   GANTT-DIAGRAMM – Erzeugt auf der aktuellen Folie
   
   Fläche:   max 118 RE breit × max 69 RE hoch
   Position: links 8 RE, oben 17 RE vom Rand
   ═══════════════════════════════════════════════════════ */

function initGantt() {

  /* Default-Datum: heute → +3 Monate */
  var today = new Date();
  var d3m = new Date(today);
  d3m.setMonth(d3m.getMonth() + 3);
  document.getElementById("ganttStart").value = isoDate(today);
  document.getElementById("ganttEnd").value = isoDate(d3m);

  /* 3 Beispiel-Phasen */
  addGanttPhase("Konzeption", today, offsetDays(today, 14), "#2e86c1");
  addGanttPhase("Umsetzung", offsetDays(today, 14), offsetDays(today, 56), "#27ae60");
  addGanttPhase("Abnahme", offsetDays(today, 56), d3m, "#e94560");

  /* Buttons */
  bind("ganttAddPhase", function () {
    var s = new Date(document.getElementById("ganttStart").value);
    var e = new Date(document.getElementById("ganttEnd").value);
    if (isNaN(s.getTime()) || isNaN(e.getTime())) { s = today; e = d3m; }
    addGanttPhase("Phase " + (ganttPhaseCount + 1), s, offsetDays(s, 14), randomColor());
  });

  bind("createGantt", function () { createGanttChart(); });
}

/* ─── Phase-UI hinzufügen ─── */
function addGanttPhase(name, start, end, color) {
  ganttPhaseCount++;
  var id = "gp_" + ganttPhaseCount;
  var div = document.createElement("div");
  div.className = "gantt-phase";
  div.id = id;
  div.innerHTML =
    '<input type="text" value="' + name + '" placeholder="Name" title="Phasenname">' +
    '<input type="date" value="' + isoDate(start) + '" title="Start">' +
    '<input type="date" value="' + isoDate(end) + '" title="Ende">' +
    '<input type="color" value="' + color + '" title="Farbe">' +
    '<button class="gantt-del" title="Entfernen">&times;</button>';
  document.getElementById("ganttPhases").appendChild(div);
  div.querySelector(".gantt-del").addEventListener("click", function () {
    div.remove();
  });
}

/* ─── Hilfsfunktionen ─── */
function isoDate(d) {
  var mm = ("0" + (d.getMonth() + 1)).slice(-2);
  var dd = ("0" + d.getDate()).slice(-2);
  return d.getFullYear() + "-" + mm + "-" + dd;
}

function offsetDays(d, n) {
  var r = new Date(d);
  r.setDate(r.getDate() + n);
  return r;
}

function randomColor() {
  var colors = ["#2e86c1","#27ae60","#e94560","#f39c12","#8e44ad","#1abc9c","#e67e22","#3498db","#d35400","#16a085"];
  return colors[Math.floor(Math.random() * colors.length)];
}

function daysBetween(a, b) {
  return Math.round((b - a) / (1000 * 60 * 60 * 24));
}

function weeksBetween(a, b) {
  return Math.ceil(daysBetween(a, b) / 7);
}

function monthsBetween(a, b) {
  return (b.getFullYear() - a.getFullYear()) * 12 + (b.getMonth() - a.getMonth()) + (b.getDate() > a.getDate() ? 1 : 0);
}

function quartersBetween(a, b) {
  return Math.ceil(monthsBetween(a, b) / 3);
}

function ganttInfo(msg, err) {
  var el = document.getElementById("ganttInfo");
  el.innerHTML = msg;
  el.className = "gantt-info" + (err ? " err" : "");
}

/* ─── Phasen aus UI lesen ─── */
function readPhases() {
  var phases = [];
  var items = document.querySelectorAll(".gantt-phase");
  items.forEach(function (div) {
    var inputs = div.querySelectorAll("input");
    var name  = inputs[0].value || "Phase";
    var start = new Date(inputs[1].value);
    var end   = new Date(inputs[2].value);
    var color = inputs[3].value || "#2e86c1";
    if (!isNaN(start.getTime()) && !isNaN(end.getTime()) && end > start) {
      phases.push({ name: name, start: start, end: end, color: color });
    }
  });
  return phases;
}

/* ═══════════════════════════════════════════
   GANTT ERZEUGEN – Hauptfunktion
   ═══════════════════════════════════════════ */
function createGanttChart() {
  if (!apiOk) { showStatus("PowerPointApi 1.5 nötig", "error"); return; }

  /* Eingaben lesen */
  var projStart = new Date(document.getElementById("ganttStart").value);
  var projEnd   = new Date(document.getElementById("ganttEnd").value);
  var unit      = document.getElementById("ganttUnit").value;
  var labelWRE  = parseInt(document.getElementById("ganttLabelW").value) || 20;
  var showHead  = document.getElementById("ganttHeader").checked;
  var showToday = document.getElementById("ganttToday").checked;
  var barColor  = document.getElementById("ganttBarColor").value;
  var headColor = document.getElementById("ganttHeadColor").value;
  var rowColor  = document.getElementById("ganttRowColor").value;

  if (isNaN(projStart.getTime()) || isNaN(projEnd.getTime())) {
    ganttInfo("❌ Ungültige Datumsangaben!", true); return;
  }
  if (projEnd <= projStart) {
    ganttInfo("❌ Ende muss nach Start liegen!", true); return;
  }

  var phases = readPhases();
  if (phases.length === 0) {
    ganttInfo("❌ Mindestens eine Phase hinzufügen!", true); return;
  }

  /* Zeiteinheiten berechnen */
  var numUnits;
  if (unit === "week")    numUnits = weeksBetween(projStart, projEnd);
  if (unit === "month")   numUnits = monthsBetween(projStart, projEnd);
  if (unit === "quarter") numUnits = quartersBetween(projStart, projEnd);
  if (numUnits < 1) numUnits = 1;

  /* Layout berechnen */
  var numRows   = phases.length + (showHead ? 1 : 0);
  var chartWRE  = GANTT_MAX_W - labelWRE;            /* Breite für Zeitachse */
  var colWRE    = Math.floor(chartWRE / numUnits);    /* Breite pro Zeiteinheit */
  if (colWRE < 1) colWRE = 1;
  var usedWRE   = colWRE * numUnits;                  /* tatsächlich genutzte Breite */
  var rowHRE    = Math.floor(GANTT_MAX_H / numRows);  /* Höhe pro Zeile */
  if (rowHRE < 2) rowHRE = 2;
  if (rowHRE > 6) rowHRE = 6;                         /* max 6 RE pro Zeile */

  var totalDays = daysBetween(projStart, projEnd);
  if (totalDays < 1) totalDays = 1;

  /* Info anzeigen */
  ganttInfo(
    "<b>" + numUnits + "</b> " + (unit === "week" ? "Wochen" : unit === "month" ? "Monate" : "Quartale") +
    " | <b>" + phases.length + "</b> Phasen | Spalte: <b>" + colWRE + " RE</b> | Zeile: <b>" + rowHRE + " RE</b>"
  );

  /* ─── PowerPoint: Shapes erzeugen ─── */
  PowerPoint.run(function (ctx) {
    var sel = ctx.presentation.getSelectedSlides();
    sel.load("items");
    return ctx.sync().then(function () {
      var slide;
      if (sel.items.length > 0) {
        slide = sel.items[0];
      } else {
        var slides = ctx.presentation.slides;
        slides.load("items");
        return ctx.sync().then(function () {
          if (!slides.items.length) { showStatus("Keine Folie!", "error"); return ctx.sync(); }
          return buildGantt(ctx, slides.items[0], projStart, projEnd, unit, numUnits,
            labelWRE, colWRE, rowHRE, usedWRE, phases, showHead, showToday,
            barColor, headColor, rowColor, totalDays);
        });
      }
      return buildGantt(ctx, slide, projStart, projEnd, unit, numUnits,
        labelWRE, colWRE, rowHRE, usedWRE, phases, showHead, showToday,
        barColor, headColor, rowColor, totalDays);
    });
  }).catch(function (e) { showStatus("Fehler: " + e.message, "error"); });
}

/* ═══════════════════════════════════════════
   GANTT SHAPES BAUEN
   ═══════════════════════════════════════════ */
function buildGantt(ctx, slide, projStart, projEnd, unit, numUnits,
                    labelWRE, colWRE, rowHRE, usedWRE, phases, showHead, showToday,
                    barColor, headColor, rowColor, totalDays) {

  var re   = gridUnitCm;
  var x0   = c2p(GANTT_LEFT * re);   /* 8 RE vom linken Rand  */
  var y0   = c2p(GANTT_TOP  * re);   /* 17 RE vom oberen Rand */
  var lbW  = c2p(labelWRE * re);     /* Label-Spaltenbreite   */
  var cW   = c2p(colWRE * re);       /* Spaltenbreite         */
  var rH   = c2p(rowHRE * re);       /* Zeilenhöhe            */
  var gap  = c2p(re);                /* 1 RE Abstand          */

  var curRow = 0;

  /* ─── KOPFZEILE (Zeitachse) ─── */
  if (showHead) {
    /* Leere Label-Zelle oben links */
    var hdrLabel = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    hdrLabel.left   = x0;
    hdrLabel.top    = y0;
    hdrLabel.width  = lbW;
    hdrLabel.height = rH;
    hdrLabel.fill.setSolidColor(headColor.replace("#", ""));
    hdrLabel.lineFormat.color = "FFFFFF";
    hdrLabel.lineFormat.weight = 0.3;
    hdrLabel.name = GANTT_TAG + "_hdr_label";

    /* Zeiteinheit-Zellen */
    for (var u = 0; u < numUnits; u++) {
      var hCell = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
      hCell.left   = x0 + lbW + gap + u * (cW + gap);
      hCell.top    = y0;
      hCell.width  = cW;
      hCell.height = rH;
      hCell.fill.setSolidColor(headColor.replace("#", ""));
      hCell.lineFormat.color = "FFFFFF";
      hCell.lineFormat.weight = 0.3;
      hCell.name = GANTT_TAG + "_hdr_" + u;

      /* Label-Text: Woche/Monat/Quartal */
      var label = getUnitLabel(projStart, u, unit);
      var tf = hCell.textFrame;
      tf.autoSizeSetting = PowerPoint.ShapeAutoSize.none;
      var tr = tf.textRange;
      tr.text = label;
      tr.font.size = 7;
      tr.font.color = "FFFFFF";
      tr.font.bold = true;
      tf.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
      tr.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
    }
    curRow = 1;
  }

  /* ─── PHASEN-ZEILEN ─── */
  for (var p = 0; p < phases.length; p++) {
    var phase = phases[p];
    var rowY = y0 + curRow * (rH + gap);

    /* Label-Zelle (Phasenname) */
    var lb = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    lb.left   = x0;
    lb.top    = rowY;
    lb.width  = lbW;
    lb.height = rH;
    lb.fill.setSolidColor(rowColor.replace("#", ""));
    lb.lineFormat.color = "CCCCCC";
    lb.lineFormat.weight = 0.3;
    lb.name = GANTT_TAG + "_label_" + p;

    var lbTf = lb.textFrame;
    lbTf.autoSizeSetting = PowerPoint.ShapeAutoSize.none;
    var lbTr = lbTf.textRange;
    lbTr.text = phase.name;
    lbTr.font.size = 7;
    lbTr.font.color = "333333";
    lbTr.font.bold = true;
    lbTf.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
    lbTr.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.left;

    /* Hintergrund-Zellen (Zeile) */
    for (var u = 0; u < numUnits; u++) {
      var bgCell = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
      bgCell.left   = x0 + lbW + gap + u * (cW + gap);
      bgCell.top    = rowY;
      bgCell.width  = cW;
      bgCell.height = rH;
      bgCell.fill.setSolidColor(u % 2 === 0 ? rowColor.replace("#", "") : "FFFFFF");
      bgCell.lineFormat.color = "E0E0E0";
      bgCell.lineFormat.weight = 0.2;
      bgCell.name = GANTT_TAG + "_bg_" + p + "_" + u;
    }

    /* ─── Balken (Phase) ─── */
    var chartStartX = x0 + lbW + gap;
    var totalChartW = numUnits * (cW + gap) - gap;

    /* Phase-Start relativ zum Projektstart */
    var pStartDay = daysBetween(projStart, phase.start);
    var pEndDay   = daysBetween(projStart, phase.end);
    if (pStartDay < 0) pStartDay = 0;
    if (pEndDay > totalDays) pEndDay = totalDays;
    if (pEndDay <= pStartDay) { curRow++; continue; }

    var barXStart = chartStartX + (pStartDay / totalDays) * totalChartW;
    var barXEnd   = chartStartX + (pEndDay / totalDays) * totalChartW;
    var barW      = barXEnd - barXStart;
    if (barW < c2p(re)) barW = c2p(re);  /* min 1 RE */

    /* Balken etwas kleiner als Zeile (Padding) */
    var barPad = rH * 0.15;
    var bar = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundedRectangle);
    bar.left   = barXStart;
    bar.top    = rowY + barPad;
    bar.width  = barW;
    bar.height = rH - barPad * 2;
    bar.fill.setSolidColor(phase.color.replace("#", ""));
    bar.lineFormat.visible = false;
    bar.name = GANTT_TAG + "_bar_" + p;

    curRow++;
  }

  /* ─── HEUTE-LINIE ─── */
  if (showToday) {
    var today = new Date();
    var todayDay = daysBetween(projStart, today);
    if (todayDay >= 0 && todayDay <= totalDays) {
      var chartStartX2 = x0 + lbW + gap;
      var totalChartW2 = numUnits * (cW + gap) - gap;
      var todayX = chartStartX2 + (todayDay / totalDays) * totalChartW2;
      var totalH = curRow * (rH + gap);

      var todayLine = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
      todayLine.left   = todayX;
      todayLine.top    = y0;
      todayLine.width  = c2p(0.05);
      todayLine.height = totalH;
      todayLine.fill.setSolidColor("E94560");
      todayLine.lineFormat.visible = false;
      todayLine.name = GANTT_TAG + "_today";
    }
  }

  return ctx.sync().then(function () {
    showStatus("Gantt: " + phases.length + " Phasen × " + numUnits + " Einheiten ✓", "success");
  });
}

/* ─── Zeiteinheit-Label erzeugen ─── */
function getUnitLabel(start, idx, unit) {
  var d = new Date(start);
  if (unit === "week") {
    d.setDate(d.getDate() + idx * 7);
    return "KW" + getWeekNumber(d);
  }
  if (unit === "month") {
    d.setMonth(d.getMonth() + idx);
    var months = ["Jan","Feb","Mrz","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez"];
    return months[d.getMonth()];
  }
  if (unit === "quarter") {
    d.setMonth(d.getMonth() + idx * 3);
    var q = Math.floor(d.getMonth() / 3) + 1;
    return "Q" + q + "/" + (d.getFullYear() % 100);
  }
  return "" + (idx + 1);
}

function getWeekNumber(d) {
  var oneJan = new Date(d.getFullYear(), 0, 1);
  var days = Math.floor((d - oneJan) / (24 * 60 * 60 * 1000));
  return Math.ceil((days + oneJan.getDay() + 1) / 7);
}
