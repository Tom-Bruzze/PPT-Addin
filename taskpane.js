/*
 ═══════════════════════════════════════════════════════
 Droege GANTT Generator  –  taskpane.js  v2.5

 FIXES v2.5:
  - PowerPoint API 1.10 vollständig kompatibel
  - InvalidArgument behoben: addGeometricShape mit GeometricShapeType enum
  - lineFormat: weight=0 statt visible=false (visible nicht in allen Versionen)
  - textFrame.wordWrap entfernt (nicht unterstützt in API 1.10)
  - textFrame.autoSizeSetting entfernt (nur in neueren APIs)
  - Alle Shape-Koordinaten validiert (>= 0, width/height >= 1)
  - fill.setSolidColor mit korrektem Hex-Format
  - Robuste Fehlerbehandlung für alle Properties

 API: PowerPointApi 1.10 (Stand: Feb 2026)

 DROEGE GROUP · 2026
 ═══════════════════════════════════════════════════════
*/

/* ═══ Konstanten ═══ */
var VERSION   = "2.5";
var API_VER   = "1.10";
var CM        = 28.3465;          /* 1 cm in PowerPoint-Points */
var gridUnitCm = 0.21;           /* Standard-Rastereinheit    */
var apiOk     = false;

/* GANTT Layout in Rastereinheiten (RE) */
var G_LEFT    = 8;
var G_TOP     = 17;
var G_W       = 118;
var G_H       = 69;
var GANTT_TAG = "DG_GANTT";
var ganttPhaseCount = 0;

/* GeometricShapeType Enum-Werte für API 1.10 */
var GST_RECTANGLE = "Rectangle";
var GST_ROUNDED_RECTANGLE = "RoundedRectangle";

Office.onReady(function (info) {
  if (info.host === Office.HostType.PowerPoint) {
    // API-Verfügbarkeit prüfen
    try {
      if (Office.context.requirements && Office.context.requirements.isSetSupported) {
        apiOk = Office.context.requirements.isSetSupported("PowerPointApi", API_VER);
      } else {
        apiOk = (typeof PowerPoint !== "undefined" && typeof PowerPoint.run === "function");
      }
    } catch (e) {
      apiOk = false;
    }
    
    initUI();
    updateInfoBar();
    setInterval(updateInfoBar, 30000);

    if (!apiOk) {
      showStatus("PowerPointApi " + API_VER + " nicht verfügbar – bitte PowerPoint aktualisieren", "warning");
    } else {
      showStatus("Bereit · API " + API_VER + " ✓", "success");
    }
  }
});

/* ═══════════════════════════════════════════
   INFO-BAR
   ═══════════════════════════════════════════ */
function updateInfoBar() {
  var now = new Date();
  var d = pad2(now.getDate()) + "." + pad2(now.getMonth() + 1) + "." + now.getFullYear();
  var t = pad2(now.getHours()) + ":" + pad2(now.getMinutes());
  var elDT  = document.getElementById("infoDateTime");
  var elVer = document.getElementById("infoVersion");
  var elApi = document.getElementById("infoApi");
  if (elDT)  elDT.textContent  = d + "  " + t;
  if (elVer) elVer.textContent = "v" + VERSION;
  if (elApi) elApi.textContent = "API " + API_VER + (apiOk ? " ✓" : " ✗");
}

/* ═══════════════════════════════════════════
   UI
   ═══════════════════════════════════════════ */
function initUI() {
  var gi = document.getElementById("gridUnit");
  if (gi) {
    gi.addEventListener("change", function () {
      var v = parseFloat(this.value);
      if (!isNaN(v) && v > 0) {
        gridUnitCm = v;
        hlPre(v);
        showStatus("RE: " + v.toFixed(2) + " cm", "info");
      }
    });
  }
  
  document.querySelectorAll(".pre").forEach(function (b) {
    b.addEventListener("click", function () {
      var v = parseFloat(this.dataset.value);
      gridUnitCm = v;
      if (gi) gi.value = v;
      hlPre(v);
      showStatus("RE: " + v.toFixed(2) + " cm", "info");
    });
  });

  bind("setSlide", setSlideSize);
  bind("createGantt", createGanttChart);
  bind("ganttAddPhase", function () {
    if (document.querySelectorAll(".gantt-phase").length >= 10) {
      ganttInfo("Maximal 10 Phasen möglich!", true);
      return;
    }
    var s = new Date(document.getElementById("ganttStart").value);
    if (isNaN(s.getTime())) s = new Date();
    addGanttPhase("Phase " + (ganttPhaseCount + 1), s, offsetDays(s, 14), randomColor());
  });

  initGanttDefaults();
}

function initGanttDefaults() {
  var today = new Date();
  var d3m = new Date(today);
  d3m.setMonth(d3m.getMonth() + 3);
  document.getElementById("ganttStart").value = isoDate(today);
  document.getElementById("ganttEnd").value   = isoDate(d3m);
  addGanttPhase("Konzeption", today, offsetDays(today, 14), "#2e86c1");
  addGanttPhase("Umsetzung", offsetDays(today, 14), offsetDays(today, 56), "#27ae60");
  addGanttPhase("Abnahme", offsetDays(today, 56), d3m, "#e94560");
}

function addGanttPhase(name, start, end, color) {
  ganttPhaseCount++;
  var div = document.createElement("div");
  div.className = "gantt-phase";
  div.innerHTML =
    '<input type="text" value="' + esc(name) + '" placeholder="Name">' +
    '<input type="date" value="' + isoDate(start) + '">' +
    '<input type="date" value="' + isoDate(end) + '">' +
    '<input type="color" value="' + color + '">' +
    '<button class="gantt-del">&times;</button>';
  document.getElementById("ganttPhases").appendChild(div);
  div.querySelector(".gantt-del").addEventListener("click", function () { div.remove(); });
}

/* ═══════════════════════════════════════════
   Helpers
   ═══════════════════════════════════════════ */
function esc(s) { 
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); 
}

function bind(id, fn) { 
  var el = document.getElementById(id); 
  if (el) el.addEventListener("click", fn); 
}

function hlPre(v) { 
  document.querySelectorAll(".pre").forEach(function (b) { 
    b.classList.toggle("active", Math.abs(parseFloat(b.dataset.value) - v) < 0.001); 
  }); 
}

function showStatus(m, t) { 
  var el = document.getElementById("status"); 
  if (!el) return; 
  el.textContent = m; 
  el.className = "sts " + (t || "info"); 
}

function pad2(n) { return n < 10 ? "0" + n : "" + n; }

function isoDate(d) { 
  return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate()); 
}

function offsetDays(d, n) { 
  var r = new Date(d); 
  r.setDate(r.getDate() + n); 
  return r; 
}

function randomColor() { 
  var c = ["#2e86c1", "#27ae60", "#e94560", "#f39c12", "#8e44ad", "#1abc9c", "#e67e22", "#3498db", "#d35400", "#16a085"]; 
  return c[Math.floor(Math.random() * c.length)]; 
}

function daysBetween(a, b) { return Math.round((b - a) / 864e5); }

function ganttInfo(m, err) { 
  var el = document.getElementById("ganttInfo"); 
  if (!el) return; 
  el.innerHTML = m; 
  el.className = "gantt-info" + (err ? " err" : ""); 
}

/**
 * Konvertiert Rastereinheiten (RE) in PowerPoint-Points
 * Validiert und rundet Werte
 */
function re2pt(re) {
  var v = Math.max(0, re) * gridUnitCm * CM;
  return Math.max(0, Math.round(v));
}

/**
 * Stellt sicher, dass ein Wert für Shape-Dimensionen gültig ist
 * width/height müssen > 0 sein
 */
function clampDimension(v) {
  if (!isFinite(v) || v < 1) return 1;
  return Math.round(v);
}

/**
 * Stellt sicher, dass ein Wert für Shape-Position gültig ist
 * left/top können 0 sein
 */
function clampPosition(v) {
  if (!isFinite(v)) return 0;
  return Math.max(0, Math.round(v));
}

function hexNoHash(c) { 
  return String(c || "").replace("#", ""); 
}

function readPhases() {
  var phases = [];
  document.querySelectorAll(".gantt-phase").forEach(function (div) {
    var inp = div.querySelectorAll("input");
    var nm  = (inp[0].value || "Phase").trim();
    var s   = new Date(inp[1].value);
    var e   = new Date(inp[2].value);
    var col = inp[3].value || "#2e86c1";
    if (!isNaN(s.getTime()) && !isNaN(e.getTime()) && e > s) {
      phases.push({ name: nm, start: s, end: e, color: col });
    }
  });
  return phases;
}

/* ═══════════════════════════════════════════
   Shape Helper Functions (API 1.10 kompatibel)
   ═══════════════════════════════════════════ */

/**
 * Erzeugt eine geometrische Form mit validierten Parametern
 * Verwendet ShapeAddOptions für API 1.10 Kompatibilität
 */
function addShape(slide, shapeType, left, top, width, height, shapeName) {
  // Alle Werte validieren
  var opts = {
    left: clampPosition(left),
    top: clampPosition(top),
    width: clampDimension(width),
    height: clampDimension(height)
  };
  
  // Shape erstellen mit GeometricShapeType String
  var shape = slide.shapes.addGeometricShape(shapeType, opts);
  
  // Name setzen (optional, in try/catch für Sicherheit)
  if (shapeName) {
    try {
      shape.name = shapeName;
    } catch (e) {
      // Name konnte nicht gesetzt werden - ignorieren
    }
  }
  
  return shape;
}

/**
 * Setzt die Füllfarbe einer Form
 * Akzeptiert Hex-Farbe ohne #
 */
function setFill(shape, colorHex) {
  try {
    shape.fill.setSolidColor(colorHex);
  } catch (e) {
    console.warn("Fill konnte nicht gesetzt werden:", e);
  }
}

/**
 * Versteckt die Linie einer Form
 * API 1.10: lineFormat.visible ist nicht zuverlässig verfügbar
 * Stattdessen: weight auf 0 setzen
 */
function hideLine(shape) {
  try {
    // Primäre Methode: Linienfarbe transparent machen
    shape.lineFormat.color = "FFFFFF";
    shape.lineFormat.weight = 0;
  } catch (e1) {
    // Fallback: visible auf false (funktioniert nicht in allen Versionen)
    try {
      shape.lineFormat.visible = false;
    } catch (e2) {
      // Beide Methoden fehlgeschlagen - ignorieren
    }
  }
}

/**
 * Setzt Text und Formatierung einer Form
 * Best-effort Ansatz mit try/catch für jede Property
 */
function setShapeText(shape, text, options) {
  var opts = options || {};
  
  try {
    var tf = shape.textFrame;
    
    // verticalAlignment (Middle, Top, Bottom)
    try {
      tf.verticalAlignment = opts.vAlign || "Middle";
    } catch (e) {}
    
    // Text setzen
    var tr = tf.textRange;
    tr.text = text;
    
    // Font-Eigenschaften
    try { tr.font.size = opts.fontSize || 7; } catch (e) {}
    try { tr.font.color = opts.fontColor || "000000"; } catch (e) {}
    try { tr.font.bold = !!opts.bold; } catch (e) {}
    
    // Font-Name (kann fehlschlagen)
    try { tr.font.name = "Segoe UI"; } catch (e) {
      try { tr.font.name = "Arial"; } catch (e2) {}
    }
    
    // Paragraph-Ausrichtung
    try {
      tr.paragraphFormat.alignment = opts.pAlign || "Left";
    } catch (e) {}
    
  } catch (e) {
    console.warn("Text konnte nicht gesetzt werden:", e);
  }
}

/* ═══════════════════════════════════════════
   Slide Size
   ═══════════════════════════════════════════ */
function setSlideSize() {
  if (!apiOk) { 
    showStatus("API " + API_VER + " erforderlich", "error"); 
    return; 
  }
  
  showStatus("Setze Folienformat...", "info");
  
  PowerPoint.run(function (ctx) {
    var ps = ctx.presentation.pageSetup;
    ps.load(["slideWidth", "slideHeight"]);
    
    return ctx.sync().then(function () {
      // 27,728 × 19,297 cm in Points
      ps.slideWidth = 786;
      ps.slideHeight = 547;
      return ctx.sync();
    }).then(function () {
      showStatus("Format: 27,728 × 19,297 cm ✓", "success");
    });
  }).catch(function (e) { 
    showStatus("Fehler: " + (e.message || e), "error"); 
  });
}

/* ═══════════════════════════════════════════
   Time Units Calculation
   ═══════════════════════════════════════════ */
function computeUnits(projStart, projEnd, unit) {
  var totalDays = daysBetween(projStart, projEnd);
  if (totalDays < 1) return [];
  var units = [];

  if (unit === "day") {
    for (var d = 0; d < totalDays; d++) {
      var dt = offsetDays(projStart, d);
      units.push({ 
        label: pad2(dt.getDate()) + "." + pad2(dt.getMonth() + 1), 
        startDay: d, 
        endDay: d + 1 
      });
    }
  } else if (unit === "week") {
    var cur = new Date(projStart);
    while (cur < projEnd) {
      var wEnd = new Date(cur);
      wEnd.setDate(wEnd.getDate() + 7);
      if (wEnd > projEnd) wEnd = new Date(projEnd);
      var sd = daysBetween(projStart, cur);
      var ed = daysBetween(projStart, wEnd);
      if (ed > sd) {
        units.push({ 
          label: "KW" + getISOWeek(cur), 
          startDay: sd, 
          endDay: ed 
        });
      }
      cur = wEnd;
    }
  } else if (unit === "month") {
    var months = ["Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"];
    var curM = new Date(projStart.getFullYear(), projStart.getMonth(), 1);
    while (curM < projEnd) {
      var mStart = new Date(curM);
      var mEnd = new Date(curM.getFullYear(), curM.getMonth() + 1, 1);
      if (mStart < projStart) mStart = new Date(projStart);
      if (mEnd > projEnd) mEnd = new Date(projEnd);
      var sd2 = daysBetween(projStart, mStart);
      var ed2 = daysBetween(projStart, mEnd);
      if (ed2 > sd2) {
        units.push({ 
          label: months[curM.getMonth()] + " " + String(curM.getFullYear()).slice(-2), 
          startDay: sd2, 
          endDay: ed2 
        });
      }
      curM = new Date(curM.getFullYear(), curM.getMonth() + 1, 1);
    }
  } else if (unit === "quarter") {
    var curQ = new Date(projStart.getFullYear(), Math.floor(projStart.getMonth() / 3) * 3, 1);
    while (curQ < projEnd) {
      var qStart = new Date(curQ);
      var qEnd = new Date(curQ.getFullYear(), Math.floor(curQ.getMonth() / 3) * 3 + 3, 1);
      if (qStart < projStart) qStart = new Date(projStart);
      if (qEnd > projEnd) qEnd = new Date(projEnd);
      var sd3 = daysBetween(projStart, qStart);
      var ed3 = daysBetween(projStart, qEnd);
      var q = Math.floor(curQ.getMonth() / 3) + 1;
      if (ed3 > sd3) {
        units.push({ 
          label: "Q" + q + "/" + String(curQ.getFullYear()).slice(-2), 
          startDay: sd3, 
          endDay: ed3 
        });
      }
      curQ = new Date(curQ.getFullYear(), Math.floor(curQ.getMonth() / 3) * 3 + 3, 1);
    }
  }
  
  return units;
}

function getISOWeek(d) {
  var tmp = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  tmp.setDate(tmp.getDate() + 3 - ((tmp.getDay() + 6) % 7));
  var week1 = new Date(tmp.getFullYear(), 0, 4);
  return 1 + Math.round(((tmp - week1) / 864e5 - 3 + ((week1.getDay() + 6) % 7)) / 7);
}

/* ═══════════════════════════════════════════
   GANTT Chart Creation
   ═══════════════════════════════════════════ */
function createGanttChart() {
  if (!apiOk) { 
    showStatus("API " + API_VER + " erforderlich", "error"); 
    return; 
  }

  var projStart = new Date(document.getElementById("ganttStart").value);
  var projEnd = new Date(document.getElementById("ganttEnd").value);
  var unit = document.getElementById("ganttUnit").value;
  var labelWRE = parseInt(document.getElementById("ganttLabelW").value, 10) || 20;
  var headerHRE = parseInt(document.getElementById("ganttHeaderH").value, 10) || 3;

  // Validierungen
  if (isNaN(projStart.getTime()) || isNaN(projEnd.getTime())) { 
    ganttInfo("Ungültige Datumsangaben!", true); 
    return; 
  }
  if (projEnd <= projStart) { 
    ganttInfo("Ende muss nach Start liegen!", true); 
    return; 
  }
  if (labelWRE >= (G_W - 5)) { 
    ganttInfo("Label-Breite zu groß (max. " + (G_W - 6) + " RE).", true); 
    return; 
  }
  if (headerHRE >= (G_H - 5)) { 
    ganttInfo("Kopfzeile zu hoch (max. " + (G_H - 6) + " RE).", true); 
    return; 
  }

  var phases = readPhases();
  if (phases.length === 0) { 
    ganttInfo("Mindestens eine Phase hinzufügen!", true); 
    return; 
  }
  if (phases.length > 10) { 
    ganttInfo("Maximal 10 Phasen!", true); 
    return; 
  }

  var timeUnits = computeUnits(projStart, projEnd, unit);
  if (timeUnits.length === 0) { 
    ganttInfo("Keine Zeiteinheiten berechenbar!", true); 
    return; 
  }
  if (timeUnits.length > 120) { 
    ganttInfo("Zu viele Einheiten (" + timeUnits.length + ") – größere Einheit wählen!", true); 
    return; 
  }

  var totalDays = daysBetween(projStart, projEnd);
  if (totalDays < 1) totalDays = 1;

  var n = phases.length;
  var chartWRE = G_W - labelWRE;
  var availHRE = G_H - headerHRE;

  // Zeilenhöhe berechnen
  var firstRowH = Math.floor((availHRE + (n * (n - 1)) / 2) / n);
  var lastRowH = firstRowH - n + 1;
  if (lastRowH < 2) {
    ganttInfo("Zu viele Phasen für die verfügbare Höhe.", true);
    return;
  }

  // Spaltenbreiten proportional berechnen
  var colWidths = [];
  var used = 0;
  for (var i = 0; i < timeUnits.length; i++) {
    var frac = (timeUnits[i].endDay - timeUnits[i].startDay) / totalDays;
    var w = Math.max(1, Math.round(frac * chartWRE));
    colWidths.push(w);
    used += w;
  }
  // Restkorrektur auf letzte Spalte
  colWidths[colWidths.length - 1] += (chartWRE - used);
  if (colWidths[colWidths.length - 1] < 1) colWidths[colWidths.length - 1] = 1;

  var unitName = { day: "Tage", week: "Kalenderwochen", month: "Monate", quarter: "Quartale" }[unit] || unit;
  ganttInfo(
    "<b>" + timeUnits.length + "</b> " + unitName +
    " | <b>" + n + "</b> Phasen" +
    " | Zeilen: <b>" + firstRowH + "→" + lastRowH + " RE</b>" +
    " | RE=" + gridUnitCm.toFixed(2) + "cm",
    false
  );

  showStatus("Erzeuge GANTT-Diagramm...", "info");

  PowerPoint.run(function (ctx) {
    var sel = ctx.presentation.getSelectedSlides();
    sel.load("items");
    
    return ctx.sync().then(function () {
      var slide;
      if (sel.items && sel.items.length > 0) {
        slide = sel.items[0];
      } else {
        var slides = ctx.presentation.slides;
        slides.load("items");
        return ctx.sync().then(function () {
          if (!slides.items || !slides.items.length) {
            throw new Error("Keine Folie verfügbar");
          }
          return buildGantt(ctx, slides.items[0], projStart, projEnd, totalDays, timeUnits, colWidths, phases, labelWRE, headerHRE, firstRowH);
        });
      }
      return buildGantt(ctx, slide, projStart, projEnd, totalDays, timeUnits, colWidths, phases, labelWRE, headerHRE, firstRowH);
    });
  }).catch(function (e) {
    showStatus("Fehler: " + (e && e.message ? e.message : e), "error");
    console.error("GANTT Fehler:", e);
  });
}

/**
 * Baut das GANTT-Diagramm auf der Folie
 */
function buildGantt(ctx, slide, projStart, projEnd, totalDays, timeUnits, colWidths, phases, labelWRE, headerHRE, firstRowH) {
  // Positionen in Points berechnen
  var x0 = re2pt(G_LEFT);
  var y0 = re2pt(G_TOP);
  var wAll = re2pt(G_W);
  var hAll = re2pt(G_H);
  var labelWPt = re2pt(labelWRE);
  var headerHPt = re2pt(headerHRE);
  var chartX = x0 + labelWPt;

  // 1. Hintergrund (weiß)
  var bg = addShape(slide, GST_RECTANGLE, x0, y0, wAll, hAll, GANTT_TAG + "_bg");
  setFill(bg, "FFFFFF");
  hideLine(bg);

  // 2. Header-Zellen (grau) mit Text
  var colX = 0;
  for (var c = 0; c < timeUnits.length; c++) {
    var cW = re2pt(colWidths[c]);
    var hdr = addShape(slide, GST_RECTANGLE, chartX + colX, y0, cW, headerHPt, GANTT_TAG + "_hdr_" + c);
    setFill(hdr, "D0D0D0");
    hideLine(hdr);
    
    var fontSize = timeUnits.length > 30 ? 5 : (timeUnits.length > 15 ? 6 : 7);
    setShapeText(hdr, timeUnits[c].label, { 
      fontSize: fontSize, 
      fontColor: "000000", 
      bold: true, 
      pAlign: "Center" 
    });
    colX += cW;
  }

  // 3. Vertikale Linien
  var lineW = Math.max(1, re2pt(0.1));

  // Linke Begrenzung
  var vl0 = addShape(slide, GST_RECTANGLE, chartX, y0, lineW, hAll, GANTT_TAG + "_vl_left");
  setFill(vl0, "B0B0B0");
  hideLine(vl0);

  // Spaltengrenzen
  colX = 0;
  for (var c2 = 0; c2 < timeUnits.length - 1; c2++) {
    colX += re2pt(colWidths[c2]);
    var vl = addShape(slide, GST_RECTANGLE, chartX + colX - lineW, y0, lineW, hAll, GANTT_TAG + "_vl_" + c2);
    setFill(vl, "B0B0B0");
    hideLine(vl);
  }

  // 4. Phasen-Zeilen mit Balken
  var chartWRE = G_W - labelWRE;
  var rowY = y0 + headerHPt;

  for (var p = 0; p < phases.length; p++) {
    var phase = phases[p];
    var rowHRE = firstRowH - p;
    var rowHPt = re2pt(rowHRE);

    // Label-Zelle (links)
    var lb = addShape(slide, GST_RECTANGLE, x0, rowY, labelWPt, rowHPt, GANTT_TAG + "_lb_" + p);
    setFill(lb, "F5F5F5");
    hideLine(lb);
    setShapeText(lb, " " + phase.name, { 
      fontSize: 7, 
      fontColor: "333333", 
      bold: true, 
      pAlign: "Left" 
    });

    // Balken berechnen
    var pStartDay = daysBetween(projStart, phase.start);
    var pEndDay = daysBetween(projStart, phase.end);
    if (pStartDay < 0) pStartDay = 0;
    if (pEndDay > totalDays) pEndDay = totalDays;

    if (pEndDay > pStartDay) {
      var barStartRE = Math.round((pStartDay / totalDays) * chartWRE);
      var barEndRE = Math.round((pEndDay / totalDays) * chartWRE);
      if (barEndRE <= barStartRE) barEndRE = barStartRE + 1;

      // Padding in RE
      var barPadRE = (rowHRE >= 3) ? 1 : 0;
      var barHRE = rowHRE - 2 * barPadRE;
      if (barHRE < 1) { barPadRE = 0; barHRE = rowHRE; }

      var bar = addShape(
        slide,
        GST_ROUNDED_RECTANGLE,
        chartX + re2pt(barStartRE),
        rowY + re2pt(barPadRE),
        re2pt(barEndRE - barStartRE),
        re2pt(barHRE),
        GANTT_TAG + "_bar_" + p
      );
      setFill(bar, hexNoHash(phase.color));
      hideLine(bar);

      // Balken-Text (wenn genug Platz)
      if (barHRE >= 2 && (barEndRE - barStartRE) >= 6) {
        setShapeText(bar, phase.name, { 
          fontSize: 6, 
          fontColor: "FFFFFF", 
          bold: true, 
          pAlign: "Center" 
        });
      }
    }

    rowY += rowHPt;
  }

  // 5. Heute-Linie (rot)
  var today = new Date();
  var todayDay = daysBetween(projStart, today);
  if (todayDay >= 0 && todayDay <= totalDays) {
    var todayRE = Math.round((todayDay / totalDays) * chartWRE);
    var tl = addShape(slide, GST_RECTANGLE, chartX + re2pt(todayRE), y0, Math.max(2, re2pt(0.2)), hAll, GANTT_TAG + "_today");
    setFill(tl, "E94560");
    hideLine(tl);

    // "HEUTE" Label (innerhalb der GANTT-Fläche)
    var lblW = re2pt(6);
    var lblH = re2pt(2);
    // Position: leicht links von der Linie, unten im Chart
    var lblX = chartX + re2pt(Math.max(0, Math.min(todayRE - 3, chartWRE - 6)));
    var lblY = y0 + hAll - lblH - 2;
    
    var todayLbl = addShape(slide, GST_RECTANGLE, lblX, lblY, lblW, lblH, GANTT_TAG + "_today_lbl");
    setFill(todayLbl, "E94560");
    hideLine(todayLbl);
    setShapeText(todayLbl, "HEUTE", { 
      fontSize: 5, 
      fontColor: "FFFFFF", 
      bold: true, 
      pAlign: "Center" 
    });
  }

  return ctx.sync().then(function () {
    showStatus("GANTT-Diagramm erfolgreich erstellt ✓", "success");
    updateInfoBar();
  });
}
