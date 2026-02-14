/*
 ═══════════════════════════════════════════════════════
 Droege GANTT Generator  –  taskpane.js  v2.3

 Rasterbasiertes GANTT-Diagramm fuer PowerPoint.
 API: PowerPointApi 1.10
 Alle Enum-Werte als String-Literale (offiziell unterstuetzt).

 Layout-Regeln:
   - Position:  8 RE von links, 17 RE von oben
   - Max-Masse: 118 RE breit, 69 RE hoch
   - Alle Abstände in ganzen RE
   - Nur vertikale Trennlinien
   - Zeilenhoehe nimmt je Phase um 1 RE ab
   - Kopfzeile: grau, schwarze Schrift
   - Hintergrund: weiss
   - Rote Heute-Linie

 DROEGE GROUP · 2025
 ═══════════════════════════════════════════════════════
*/

/* ═══ Konstanten ═══ */
var VERSION   = "2.3";
var API_VER   = "1.10";
var CM        = 28.3465;          /* 1 cm in PowerPoint-Points */
var gridUnitCm = 0.21;           /* Standard-Rastereinheit    */
var apiOk     = false;

/* GANTT Layout in Rastereinheiten (RE) */
var G_LEFT    = 8;               /* Abstand links             */
var G_TOP     = 17;              /* Abstand oben              */
var G_W       = 118;             /* max. Breite               */
var G_H       = 69;              /* max. Hoehe                */
var GANTT_TAG = "DG_GANTT";      /* Shape-Name-Prefix         */
var ganttPhaseCount = 0;

/* Enum String-Literale (PowerPointApi 1.10) */
var SH_RECT   = "Rectangle";
var SH_RRECT  = "RoundedRectangle";
var AS_NONE   = "None";
var VA_MID    = "Middle";
var PA_CENTER = "Center";
var PA_LEFT   = "Left";

/* ═══════════════════════════════════════════
   Office.onReady
   ═══════════════════════════════════════════ */
Office.onReady(function (info) {
  if (info.host === Office.HostType.PowerPoint) {
    if (Office.context.requirements && Office.context.requirements.isSetSupported) {
      apiOk = Office.context.requirements.isSetSupported("PowerPointApi", API_VER);
    } else {
      apiOk = (typeof PowerPoint !== "undefined" && typeof PowerPoint.run === "function");
    }
    initUI();
    updateInfoBar();
    setInterval(updateInfoBar, 30000);

    if (!apiOk) {
      showStatus("PowerPointApi " + API_VER + " nicht verfuegbar", "warning");
    } else {
      showStatus("Bereit  ·  API " + API_VER + " ✓", "success");
    }
  }
});

/* ═══════════════════════════════════════════
   INFO-BAR (Datum / Uhrzeit / Version / API)
   ═══════════════════════════════════════════ */
function updateInfoBar() {
  var now = new Date();
  var d = pad2(now.getDate()) + "." + pad2(now.getMonth()+1) + "." + now.getFullYear();
  var t = pad2(now.getHours()) + ":" + pad2(now.getMinutes());
  var elDT  = document.getElementById("infoDateTime");
  var elVer = document.getElementById("infoVersion");
  var elApi = document.getElementById("infoApi");
  if (elDT)  elDT.textContent  = d + "  " + t;
  if (elVer) elVer.textContent = "v" + VERSION;
  if (elApi) elApi.textContent = "API " + API_VER + (apiOk ? " ✓" : " ✗");
}

/* ═══════════════════════════════════════════
   UI INIT
   ═══════════════════════════════════════════ */
function initUI() {
  var gi = document.getElementById("gridUnit");
  gi.addEventListener("change", function () {
    var v = parseFloat(this.value);
    if (!isNaN(v) && v > 0) { gridUnitCm = v; hlPre(v); showStatus("RE: " + v.toFixed(2) + " cm", "info"); }
  });
  document.querySelectorAll(".pre").forEach(function (b) {
    b.addEventListener("click", function () {
      var v = parseFloat(this.dataset.value);
      gridUnitCm = v; gi.value = v; hlPre(v);
      showStatus("RE: " + v.toFixed(2) + " cm", "info");
    });
  });
  bind("setSlide",    setSlideSize);
  bind("createGantt", createGanttChart);
  bind("ganttAddPhase", function () {
    if (document.querySelectorAll(".gantt-phase").length >= 10) {
      ganttInfo("Maximal 10 Phasen moeglich!", true); return;
    }
    var s = new Date(document.getElementById("ganttStart").value);
    if (isNaN(s.getTime())) s = new Date();
    addGanttPhase("Phase " + (ganttPhaseCount+1), s, offsetDays(s,14), randomColor());
  });
  initGanttDefaults();
}

function initGanttDefaults() {
  var today = new Date();
  var d3m   = new Date(today); d3m.setMonth(d3m.getMonth()+3);
  document.getElementById("ganttStart").value = isoDate(today);
  document.getElementById("ganttEnd").value   = isoDate(d3m);
  addGanttPhase("Konzeption", today, offsetDays(today,14), "#2e86c1");
  addGanttPhase("Umsetzung",  offsetDays(today,14), offsetDays(today,56), "#27ae60");
  addGanttPhase("Abnahme",    offsetDays(today,56), d3m, "#e94560");
}

/* ═══════════════════════════════════════════
   PHASE UI
   ═══════════════════════════════════════════ */
function addGanttPhase(name, start, end, color) {
  ganttPhaseCount++;
  var div = document.createElement("div");
  div.className = "gantt-phase";
  div.innerHTML =
    '<input type="text" value="' + name + '" placeholder="Name">' +
    '<input type="date" value="' + isoDate(start) + '">' +
    '<input type="date" value="' + isoDate(end) + '">' +
    '<input type="color" value="' + color + '">' +
    '<button class="gantt-del">&times;</button>';
  document.getElementById("ganttPhases").appendChild(div);
  div.querySelector(".gantt-del").addEventListener("click", function(){ div.remove(); });
}

/* ═══════════════════════════════════════════
   HILFSFUNKTIONEN
   ═══════════════════════════════════════════ */
function bind(id, fn) { var el = document.getElementById(id); if(el) el.addEventListener("click", fn); }
function hlPre(v) { document.querySelectorAll(".pre").forEach(function(b){ b.classList.toggle("active", Math.abs(parseFloat(b.dataset.value)-v)<0.001); }); }
function showStatus(m,t) { var el=document.getElementById("status"); el.textContent=m; el.className="sts "+(t||"info"); }
function c2p(cm) { return cm * CM; }
function re2p(re) { return re * gridUnitCm * CM; }
function pad2(n) { return n < 10 ? "0"+n : ""+n; }
function isoDate(d) { return d.getFullYear()+"-"+pad2(d.getMonth()+1)+"-"+pad2(d.getDate()); }
function offsetDays(d,n) { var r=new Date(d); r.setDate(r.getDate()+n); return r; }
function randomColor() { var c=["#2e86c1","#27ae60","#e94560","#f39c12","#8e44ad","#1abc9c","#e67e22","#3498db","#d35400","#16a085"]; return c[Math.floor(Math.random()*c.length)]; }
function daysBetween(a,b) { return Math.round((b-a)/(864e5)); }
function ganttInfo(m,err) { var el=document.getElementById("ganttInfo"); el.innerHTML=m; el.className="gantt-info"+(err?" err":""); }

function readPhases() {
  var phases = [];
  document.querySelectorAll(".gantt-phase").forEach(function(div) {
    var inp = div.querySelectorAll("input");
    var nm = inp[0].value||"Phase", s = new Date(inp[1].value), e = new Date(inp[2].value), col = inp[3].value||"#2e86c1";
    if (!isNaN(s.getTime()) && !isNaN(e.getTime()) && e > s) phases.push({name:nm,start:s,end:e,color:col});
  });
  return phases;
}

/* ═══════════════════════════════════════════
   TEXT auf Shape setzen
   ═══════════════════════════════════════════ */
function setShapeText(shape, text, opts) {
  var tf = shape.textFrame;
  tf.autoSizeSetting   = AS_NONE;
  tf.verticalAlignment = opts.vAlign || VA_MID;
  tf.wordWrap = false;
  var tr = tf.textRange;
  tr.text = text;
  tr.font.size  = opts.fontSize || 7;
  tr.font.color = opts.fontColor || "000000";
  tr.font.bold  = !!opts.bold;
  tr.font.name  = "Segoe UI";
  tr.paragraphFormat.alignment = opts.pAlign || PA_LEFT;
}

/* ═══════════════════════════════════════════
   PAPIERFORMAT
   ═══════════════════════════════════════════ */
function setSlideSize() {
  if (!apiOk) { showStatus("API " + API_VER + " noetig","error"); return; }
  PowerPoint.run(function(ctx) {
    var ps = ctx.presentation.pageSetup;
    ps.load(["slideWidth","slideHeight"]);
    return ctx.sync().then(function(){
      ps.slideWidth  = 786;   /* 27,728 cm */
      ps.slideHeight = 547;   /* 19,297 cm */
      return ctx.sync();
    }).then(function(){ showStatus("Format: 27,728 × 19,297 cm ✓","success"); });
  }).catch(function(e){ showStatus("Fehler: "+e.message,"error"); });
}

/* ═══════════════════════════════════════════
   ZEITEINHEITEN BERECHNEN
   ═══════════════════════════════════════════ */

/* Gibt Array zurueck: [{label, startDay, endDay}, ...] */
function computeUnits(projStart, projEnd, unit) {
  var totalDays = daysBetween(projStart, projEnd);
  if (totalDays < 1) return [];
  var units = [];

  if (unit === "day") {
    for (var d = 0; d < totalDays; d++) {
      var dt = offsetDays(projStart, d);
      units.push({ label: pad2(dt.getDate())+"."+pad2(dt.getMonth()+1), startDay: d, endDay: d+1 });
    }
  }
  else if (unit === "week") {
    var cur = new Date(projStart);
    while (cur < projEnd) {
      var wEnd = new Date(cur); wEnd.setDate(wEnd.getDate() + 7);
      if (wEnd > projEnd) wEnd = new Date(projEnd);
      var sd = daysBetween(projStart, cur);
      var ed = daysBetween(projStart, wEnd);
      units.push({ label: "KW"+getISOWeek(cur), startDay: sd, endDay: ed });
      cur = wEnd;
    }
  }
  else if (unit === "month") {
    var cur = new Date(projStart.getFullYear(), projStart.getMonth(), 1);
    if (cur < projStart) cur = new Date(projStart);
    var months = ["Jan","Feb","Mrz","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez"];
    while (cur < projEnd) {
      var mStart = new Date(cur);
      var mEnd = new Date(cur.getFullYear(), cur.getMonth()+1, 1);
      if (mEnd > projEnd) mEnd = new Date(projEnd);
      if (mStart < projStart) mStart = new Date(projStart);
      var sd = daysBetween(projStart, mStart);
      var ed = daysBetween(projStart, mEnd);
      if (ed > sd) units.push({ label: months[cur.getMonth()]+" "+String(cur.getFullYear()).slice(-2), startDay: sd, endDay: ed });
      cur = new Date(cur.getFullYear(), cur.getMonth()+1, 1);
    }
  }
  else if (unit === "quarter") {
    var cur = new Date(projStart.getFullYear(), Math.floor(projStart.getMonth()/3)*3, 1);
    if (cur < projStart) cur = new Date(projStart);
    while (cur < projEnd) {
      var qStart = new Date(cur);
      var qEnd = new Date(cur.getFullYear(), Math.floor(cur.getMonth()/3)*3+3, 1);
      if (qEnd > projEnd) qEnd = new Date(projEnd);
      if (qStart < projStart) qStart = new Date(projStart);
      var sd = daysBetween(projStart, qStart);
      var ed = daysBetween(projStart, qEnd);
      var q  = Math.floor(cur.getMonth()/3)+1;
      if (ed > sd) units.push({ label: "Q"+q+"/"+String(cur.getFullYear()).slice(-2), startDay: sd, endDay: ed });
      cur = new Date(cur.getFullYear(), Math.floor(cur.getMonth()/3)*3+3, 1);
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
   GANTT ERZEUGEN – Hauptfunktion
   ═══════════════════════════════════════════ */
function createGanttChart() {
  if (!apiOk) { showStatus("API "+API_VER+" noetig","error"); return; }

  var projStart = new Date(document.getElementById("ganttStart").value);
  var projEnd   = new Date(document.getElementById("ganttEnd").value);
  var unit      = document.getElementById("ganttUnit").value;
  var labelWRE  = parseInt(document.getElementById("ganttLabelW").value) || 20;
  var headerHRE = parseInt(document.getElementById("ganttHeaderH").value) || 3;

  if (isNaN(projStart.getTime()) || isNaN(projEnd.getTime())) { ganttInfo("Ungueltige Datumsangaben!",true); return; }
  if (projEnd <= projStart) { ganttInfo("Ende muss nach Start liegen!",true); return; }

  var phases = readPhases();
  if (phases.length === 0)  { ganttInfo("Mindestens eine Phase hinzufuegen!",true); return; }
  if (phases.length > 10)   { ganttInfo("Maximal 10 Phasen!",true); return; }

  var timeUnits = computeUnits(projStart, projEnd, unit);
  if (timeUnits.length === 0) { ganttInfo("Keine Zeiteinheiten berechenbar!",true); return; }
  if (timeUnits.length > 120) { ganttInfo("Zu viele Einheiten ("+timeUnits.length+"). Groessere Einheit waehlen!",true); return; }

  var totalDays = daysBetween(projStart, projEnd);
  var n = phases.length;

  /* Chartflaeche in RE */
  var chartWRE = G_W - labelWRE;              /* Breite fuer Zeitachse     */
  var availHRE = G_H - headerHRE;             /* Hoehe fuer Phasenzeilen   */

  /* Zeilenhoehen berechnen: erste Zeile = h RE, dann h-1, h-2, ... */
  var firstRowH = Math.floor((availHRE + n*(n-1)/2) / n);
  var lastRowH  = firstRowH - n + 1;

  if (lastRowH < 2) {
    ganttInfo("Zu viele Phasen fuer die verfuegbare Hoehe! Max " + Math.floor(Math.sqrt(2*availHRE)) + " Phasen.",true);
    return;
  }

  /* Spaltenbreite: proportional zu Tagen, gerundet auf ganze RE */
  var colWidths = [];
  var usedW = 0;
  for (var i = 0; i < timeUnits.length; i++) {
    var frac = (timeUnits[i].endDay - timeUnits[i].startDay) / totalDays;
    var w = Math.max(1, Math.round(frac * chartWRE));
    colWidths.push(w);
    usedW += w;
  }
  /* Korrektur: ueberschuss/defizit auf letzte Spalte */
  if (usedW !== chartWRE) colWidths[colWidths.length-1] += (chartWRE - usedW);
  if (colWidths[colWidths.length-1] < 1) colWidths[colWidths.length-1] = 1;

  /* Info-Anzeige */
  var unitName = {day:"Tage",week:"KW",month:"Monate",quarter:"Quartale"}[unit];
  ganttInfo(
    "<b>" + timeUnits.length + "</b> " + unitName +
    " | <b>" + n + "</b> Phasen" +
    " | Zeilen: <b>" + firstRowH + "→" + lastRowH + " RE</b>" +
    " | RE=" + gridUnitCm.toFixed(2) + "cm"
  );

  showStatus("Erzeuge GANTT...", "info");

  PowerPoint.run(function(ctx) {
    var sel = ctx.presentation.getSelectedSlides();
    sel.load("items");
    return ctx.sync().then(function() {
      var slide;
      if (sel.items.length > 0) {
        slide = sel.items[0];
      } else {
        var slides = ctx.presentation.slides;
        slides.load("items");
        return ctx.sync().then(function() {
          if (!slides.items.length) { showStatus("Keine Folie!","error"); return ctx.sync(); }
          return buildGantt(ctx, slides.items[0], projStart, projEnd, totalDays,
            timeUnits, colWidths, phases, labelWRE, headerHRE, firstRowH, unit);
        });
      }
      return buildGantt(ctx, slide, projStart, projEnd, totalDays,
        timeUnits, colWidths, phases, labelWRE, headerHRE, firstRowH, unit);
    });
  }).catch(function(e){ showStatus("Fehler: "+e.message,"error"); });
}

/* ═══════════════════════════════════════════
   GANTT SHAPES BAUEN
   ═══════════════════════════════════════════ */
function buildGantt(ctx, slide, projStart, projEnd, totalDays,
                    timeUnits, colWidths, phases, labelWRE, headerHRE, firstRowH, unit) {

  var x0 = re2p(G_LEFT);
  var y0 = re2p(G_TOP);
  var totalWPt = re2p(G_W);
  var totalHPt = re2p(G_H);
  var labelWPt = re2p(labelWRE);
  var headerHPt = re2p(headerHRE);
  var chartX   = x0 + labelWPt;

  /* ── 1. WEISSER HINTERGRUND (gesamter GANTT-Bereich) ── */
  var bg = slide.shapes.addGeometricShape(SH_RECT);
  bg.left = x0;  bg.top = y0;  bg.width = totalWPt;  bg.height = totalHPt;
  bg.fill.setSolidColor("FFFFFF");
  bg.lineFormat.visible = false;
  bg.name = GANTT_TAG + "_bg";

  /* ── 2. KOPFZEILE: Graue Zellen pro Zeiteinheit ── */
  var colX = 0;
  for (var c = 0; c < timeUnits.length; c++) {
    var cW = re2p(colWidths[c]);
    var hdr = slide.shapes.addGeometricShape(SH_RECT);
    hdr.left   = chartX + colX;
    hdr.top    = y0;
    hdr.width  = cW;
    hdr.height = headerHPt;
    hdr.fill.setSolidColor("D0D0D0");
    hdr.lineFormat.visible = false;
    hdr.name = GANTT_TAG + "_hdr_" + c;

    /* Beschriftung */
    var fontSize = timeUnits.length > 30 ? 5 : (timeUnits.length > 15 ? 6 : 7);
    setShapeText(hdr, timeUnits[c].label, {
      fontSize: fontSize,
      fontColor: "000000",
      bold: true,
      pAlign: PA_CENTER
    });

    colX += cW;
  }

  /* ── 3. VERTIKALE TRENNLINIEN ── */
  colX = 0;
  for (var c = 0; c < timeUnits.length; c++) {
    colX += re2p(colWidths[c]);
    /* Linie am rechten Rand jeder Spalte (ausser letzte) */
    if (c < timeUnits.length - 1) {
      var vl = slide.shapes.addGeometricShape(SH_RECT);
      vl.left   = chartX + colX - re2p(0.1);
      vl.top    = y0;
      vl.width  = re2p(0.1);
      vl.height = totalHPt;
      vl.fill.setSolidColor("B0B0B0");
      vl.lineFormat.visible = false;
      vl.name = GANTT_TAG + "_vl_" + c;
    }
  }

  /* Linke Begrenzung der Chartflaeche */
  var vlLeft = slide.shapes.addGeometricShape(SH_RECT);
  vlLeft.left   = chartX;
  vlLeft.top    = y0;
  vlLeft.width  = re2p(0.1);
  vlLeft.height = totalHPt;
  vlLeft.fill.setSolidColor("B0B0B0");
  vlLeft.lineFormat.visible = false;
  vlLeft.name = GANTT_TAG + "_vl_left";

  /* ── 4. PHASEN-ZEILEN (abnehmende Hoehe) ── */
  var rowY = y0 + headerHPt;
  var chartTotalW = re2p(G_W - labelWRE);     /* Breite der Zeitachse in pt */

  for (var p = 0; p < phases.length; p++) {
    var phase = phases[p];
    var rowH  = firstRowH - p;                 /* RE: nimmt um 1 ab        */
    var rowHPt = re2p(rowH);

    /* Label-Zelle (links) */
    var lb = slide.shapes.addGeometricShape(SH_RECT);
    lb.left   = x0;
    lb.top    = rowY;
    lb.width  = labelWPt;
    lb.height = rowHPt;
    lb.fill.setSolidColor("F5F5F5");
    lb.lineFormat.visible = false;
    lb.name = GANTT_TAG + "_lb_" + p;
    setShapeText(lb, " " + phase.name, {
      fontSize: 7,
      fontColor: "333333",
      bold: true,
      pAlign: PA_LEFT
    });

    /* Balken: Position aus Tagen berechnen, auf RE runden */
    var pStartDay = daysBetween(projStart, phase.start);
    var pEndDay   = daysBetween(projStart, phase.end);
    if (pStartDay < 0)          pStartDay = 0;
    if (pEndDay   > totalDays)  pEndDay   = totalDays;

    if (pEndDay > pStartDay) {
      var barStartRE = Math.round((pStartDay / totalDays) * (G_W - labelWRE));
      var barEndRE   = Math.round((pEndDay   / totalDays) * (G_W - labelWRE));
      if (barEndRE <= barStartRE) barEndRE = barStartRE + 1;

      var barPadRE = 1;   /* 1 RE Padding oben/unten */
      var barHRE   = rowH - 2 * barPadRE;
      if (barHRE < 1) { barPadRE = 0; barHRE = rowH; }

      var bar = slide.shapes.addGeometricShape(SH_RRECT);
      bar.left   = chartX + re2p(barStartRE);
      bar.top    = rowY + re2p(barPadRE);
      bar.width  = re2p(barEndRE - barStartRE);
      bar.height = re2p(barHRE);
      bar.fill.setSolidColor(phase.color.replace("#",""));
      bar.lineFormat.visible = false;
      bar.name = GANTT_TAG + "_bar_" + p;

      /* Balken-Text (Phasenname im Balken) */
      if (barHRE >= 2 && (barEndRE - barStartRE) >= 4) {
        setShapeText(bar, phase.name, {
          fontSize: 6,
          fontColor: "FFFFFF",
          bold: true,
          pAlign: PA_CENTER
        });
      }
    }

    rowY += rowHPt;
  }

  /* ── 5. HEUTE-LINIE (rot) ── */
  var today    = new Date();
  var todayDay = daysBetween(projStart, today);
  if (todayDay >= 0 && todayDay <= totalDays) {
    var todayRE = Math.round((todayDay / totalDays) * (G_W - labelWRE));
    var tl = slide.shapes.addGeometricShape(SH_RECT);
    tl.left   = chartX + re2p(todayRE);
    tl.top    = y0;
    tl.width  = re2p(0.2);
    tl.height = totalHPt;
    tl.fill.setSolidColor("E94560");
    tl.lineFormat.visible = false;
    tl.name = GANTT_TAG + "_today";

    /* Kleines "Heute"-Label */
    var todayLbl = slide.shapes.addGeometricShape(SH_RECT);
    todayLbl.left   = chartX + re2p(todayRE) - re2p(2);
    todayLbl.top    = y0 + totalHPt;
    todayLbl.width  = re2p(5);
    todayLbl.height = re2p(2);
    todayLbl.fill.setSolidColor("E94560");
    todayLbl.lineFormat.visible = false;
    todayLbl.name = GANTT_TAG + "_today_lbl";
    setShapeText(todayLbl, "HEUTE", {
      fontSize: 5,
      fontColor: "FFFFFF",
      bold: true,
      pAlign: PA_CENTER
    });
  }

  return ctx.sync().then(function() {
    showStatus("GANTT: " + phases.length + " Phasen × " + timeUnits.length + " Einheiten ✓", "success");
    updateInfoBar();
  });
}
