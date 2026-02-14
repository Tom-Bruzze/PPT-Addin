/*
 ═══════════════════════════════════════════════════════
 Droege GANTT Generator  –  taskpane.js  v2.2
 
 PowerPoint Add-in – erstellt rasterbasierte GANTT-Charts.
 Nutzt PowerPointApi 1.10 (Requirement Set).
 
 WICHTIG:  Die PowerPoint JS API akzeptiert fuer Enums
           sowohl die Enum-Objekte als auch deren
           korrespondierende String-Werte.
           In v2.2 verwenden wir durchgaengig String-
           Literale – das ist offiziell unterstuetzt und
           umgeht jedes Enum-Aufloesung-Problem.
           Referenz: https://learn.microsoft.com/en-us/
           javascript/api/powerpoint
 
 DROEGE GROUP · 2025
 ═══════════════════════════════════════════════════════
 */

/* ═══ Konstanten ═══ */
var CM = 28.3465;
var gridUnitCm = 0.21;
var apiOk = false;

var GANTT_MAX_W = 118;
var GANTT_MAX_H = 69;
var GANTT_LEFT  = 8;
var GANTT_TOP   = 17;
var GANTT_TAG   = "DROEGE_GANTT";
var ganttPhaseCount = 0;

/* ═══════════════════════════════════════════
   ENUM STRING-KONSTANTEN (PowerPointApi 1.10)
   
   Die Office JS API akzeptiert diese String-
   Werte ueberall dort, wo ein Enum erwartet
   wird. Das ist sicherer als die Referenz auf
   PowerPoint.XyzEnum.value, weil der Namespace
   je nach Ladezeit undefiniert sein kann.
   ═══════════════════════════════════════════ */

var SHAPE_RECTANGLE       = "Rectangle";
var SHAPE_ROUNDED_RECT    = "RoundedRectangle";
var AUTOSIZE_NONE         = "None";
var VALIGN_MIDDLE         = "Middle";
var PALIGN_CENTER         = "Center";
var PALIGN_LEFT           = "Left";

/* ═══════════════════════════════════════════
   Office Init
   ═══════════════════════════════════════════ */
Office.onReady(function (info) {
  if (info.host === Office.HostType.PowerPoint) {

    /* Pruefe ob PowerPointApi 1.10 verfuegbar ist */
    if (Office.context.requirements && Office.context.requirements.isSetSupported) {
      apiOk = Office.context.requirements.isSetSupported("PowerPointApi", "1.10");
    } else {
      apiOk = (typeof PowerPoint !== "undefined" && typeof PowerPoint.run === "function");
    }

    initUI();

    if (!apiOk) {
      showStatus("PowerPointApi 1.10 nicht verfuegbar – bitte PowerPoint aktualisieren", "warning");
    } else {
      showStatus("Bereit  ·  API 1.10 ✓", "success");
    }
  }
});

/* ═══════════════════════════════════════════
   UI INIT
   ═══════════════════════════════════════════ */
function initUI() {

  var gi = document.getElementById("gridUnit");
  gi.addEventListener("change", function () {
    var v = parseFloat(this.value);
    if (!isNaN(v) && v > 0) {
      gridUnitCm = v;
      hlPre(v);
      showStatus("RE: " + v.toFixed(2) + " cm", "info");
    }
  });

  document.querySelectorAll(".pre").forEach(function (b) {
    b.addEventListener("click", function () {
      var v = parseFloat(this.dataset.value);
      gridUnitCm = v;
      gi.value = v;
      hlPre(v);
      showStatus("RE: " + v.toFixed(2) + " cm", "info");
    });
  });

  bind("setSlide", function () { setSlideSize(); });

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
   TEXT-FORMATIERUNG
   
   Setzt Text + Schriftart + Alignment auf
   einem Shape. Nutzt ausschliesslich String-
   Enum-Werte (API 1.10 kompatibel).
   ═══════════════════════════════════════════ */
function formatShapeText(shape, text, opts) {
  /* opts: { fontSize, fontColor, bold, vAlign, pAlign } */
  var tf = shape.textFrame;
  tf.autoSizeSetting = AUTOSIZE_NONE;
  tf.verticalAlignment = opts.vAlign || VALIGN_MIDDLE;

  var tr = tf.textRange;
  tr.text = text;
  tr.font.size  = opts.fontSize || 7;
  tr.font.color = opts.fontColor || "333333";
  tr.font.bold  = opts.bold || false;
  tr.paragraphFormat.alignment = opts.pAlign || PALIGN_LEFT;
}

/* ═══════════════════════════════════════════
   PAPIERFORMAT 27,728 x 19,297 cm
   ═══════════════════════════════════════════ */
function setSlideSize() {
  if (!apiOk) { showStatus("PowerPointApi 1.10 noetig", "error"); return; }
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
   GANTT-DIAGRAMM
   ═══════════════════════════════════════════════════════ */

function initGantt() {

  var today = new Date();
  var d3m = new Date(today);
  d3m.setMonth(d3m.getMonth() + 3);
  document.getElementById("ganttStart").value = isoDate(today);
  document.getElementById("ganttEnd").value = isoDate(d3m);

  addGanttPhase("Konzeption", today, offsetDays(today, 14), "#2e86c1");
  addGanttPhase("Umsetzung", offsetDays(today, 14), offsetDays(today, 56), "#27ae60");
  addGanttPhase("Abnahme", offsetDays(today, 56), d3m, "#e94560");

  bind("ganttAddPhase", function () {
    var s = new Date(document.getElementById("ganttStart").value);
    var e = new Date(document.getElementById("ganttEnd").value);
    if (isNaN(s.getTime()) || isNaN(e.getTime())) { s = new Date(); e = new Date(); e.setMonth(e.getMonth()+3); }
    addGanttPhase("Phase " + (ganttPhaseCount + 1), s, offsetDays(s, 14), randomColor());
  });

  bind("createGantt", function () { createGanttChart(); });
}

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
  if (!apiOk) { showStatus("PowerPointApi 1.10 noetig", "error"); return; }

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
    ganttInfo("Ungueltige Datumsangaben!", true); return;
  }
  if (projEnd <= projStart) {
    ganttInfo("Ende muss nach Start liegen!", true); return;
  }

  var phases = readPhases();
  if (phases.length === 0) {
    ganttInfo("Mindestens eine Phase hinzufuegen!", true); return;
  }

  var numUnits;
  if (unit === "week")    numUnits = weeksBetween(projStart, projEnd);
  if (unit === "month")   numUnits = monthsBetween(projStart, projEnd);
  if (unit === "quarter") numUnits = quartersBetween(projStart, projEnd);
  if (!numUnits || numUnits < 1) numUnits = 1;

  var numRows   = phases.length + (showHead ? 1 : 0);
  var chartWRE  = GANTT_MAX_W - labelWRE;
  var colWRE    = Math.floor(chartWRE / numUnits);
  if (colWRE < 1) colWRE = 1;
  var usedWRE   = colWRE * numUnits;
  var rowHRE    = Math.floor(GANTT_MAX_H / numRows);
  if (rowHRE < 2) rowHRE = 2;
  if (rowHRE > 6) rowHRE = 6;

  var totalDays = daysBetween(projStart, projEnd);
  if (totalDays < 1) totalDays = 1;

  ganttInfo(
    "<b>" + numUnits + "</b> " + (unit === "week" ? "Wochen" : unit === "month" ? "Monate" : "Quartale") +
    " | <b>" + phases.length + "</b> Phasen | Spalte: <b>" + colWRE + " RE</b> | Zeile: <b>" + rowHRE + " RE</b>"
  );

  showStatus("Erzeuge GANTT...", "info");

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
  var x0   = c2p(GANTT_LEFT * re);
  var y0   = c2p(GANTT_TOP  * re);
  var lbW  = c2p(labelWRE * re);
  var cW   = c2p(colWRE * re);
  var rH   = c2p(rowHRE * re);
  var gap  = c2p(re);

  var curRow = 0;

  /* ─── KOPFZEILE ─── */
  if (showHead) {
    var hdrLabel = slide.shapes.addGeometricShape(SHAPE_RECTANGLE);
    hdrLabel.left   = x0;
    hdrLabel.top    = y0;
    hdrLabel.width  = lbW;
    hdrLabel.height = rH;
    hdrLabel.fill.setSolidColor(headColor.replace("#", ""));
    hdrLabel.lineFormat.color = "FFFFFF";
    hdrLabel.lineFormat.weight = 0.3;
    hdrLabel.name = GANTT_TAG + "_hdr_label";

    for (var u = 0; u < numUnits; u++) {
      var hCell = slide.shapes.addGeometricShape(SHAPE_RECTANGLE);
      hCell.left   = x0 + lbW + gap + u * (cW + gap);
      hCell.top    = y0;
      hCell.width  = cW;
      hCell.height = rH;
      hCell.fill.setSolidColor(headColor.replace("#", ""));
      hCell.lineFormat.color = "FFFFFF";
      hCell.lineFormat.weight = 0.3;
      hCell.name = GANTT_TAG + "_hdr_" + u;

      var label = getUnitLabel(projStart, u, unit);
      formatShapeText(hCell, label, {
        fontSize: 7,
        fontColor: "FFFFFF",
        bold: true,
        vAlign: VALIGN_MIDDLE,
        pAlign: PALIGN_CENTER
      });
    }
    curRow = 1;
  }

  /* ─── PHASEN-ZEILEN ─── */
  for (var p = 0; p < phases.length; p++) {
    var phase = phases[p];
    var rowY = y0 + curRow * (rH + gap);

    /* Label-Zelle */
    var lb = slide.shapes.addGeometricShape(SHAPE_RECTANGLE);
    lb.left   = x0;
    lb.top    = rowY;
    lb.width  = lbW;
    lb.height = rH;
    lb.fill.setSolidColor(rowColor.replace("#", ""));
    lb.lineFormat.color = "CCCCCC";
    lb.lineFormat.weight = 0.3;
    lb.name = GANTT_TAG + "_label_" + p;

    formatShapeText(lb, phase.name, {
      fontSize: 7,
      fontColor: "333333",
      bold: true,
      vAlign: VALIGN_MIDDLE,
      pAlign: PALIGN_LEFT
    });

    /* Hintergrund-Zellen */
    for (var u = 0; u < numUnits; u++) {
      var bgCell = slide.shapes.addGeometricShape(SHAPE_RECTANGLE);
      bgCell.left   = x0 + lbW + gap + u * (cW + gap);
      bgCell.top    = rowY;
      bgCell.width  = cW;
      bgCell.height = rH;
      bgCell.fill.setSolidColor(u % 2 === 0 ? rowColor.replace("#", "") : "FFFFFF");
      bgCell.lineFormat.color = "E0E0E0";
      bgCell.lineFormat.weight = 0.2;
      bgCell.name = GANTT_TAG + "_bg_" + p + "_" + u;
    }

    /* ─── Balken ─── */
    var chartStartX = x0 + lbW + gap;
    var totalChartW = numUnits * (cW + gap) - gap;

    var pStartDay = daysBetween(projStart, phase.start);
    var pEndDay   = daysBetween(projStart, phase.end);
    if (pStartDay < 0) pStartDay = 0;
    if (pEndDay > totalDays) pEndDay = totalDays;
    if (pEndDay <= pStartDay) { curRow++; continue; }

    var barXStart = chartStartX + (pStartDay / totalDays) * totalChartW;
    var barXEnd   = chartStartX + (pEndDay / totalDays) * totalChartW;
    var barW      = barXEnd - barXStart;
    if (barW < c2p(re)) barW = c2p(re);

    var barPad = rH * 0.15;
    var bar = slide.shapes.addGeometricShape(SHAPE_ROUNDED_RECT);
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

      var todayLine = slide.shapes.addGeometricShape(SHAPE_RECTANGLE);
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

/* ─── Zeiteinheit-Label ─── */
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
