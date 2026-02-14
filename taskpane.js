/* 
   DROEGE GANTT Generator – taskpane.js
   Build 14.02.2026
   
   Erzeugt GANTT-Diagramme direkt in PowerPoint als Shapes.
   Alle Maße in Rastereinheiten (RE), Position immer ganzzahlig.
 */

/* ══════════════════════════════════════════════════════════════
   GLOBALS
   ══════════════════════════════════════════════════════════════ */
var gridUnitCm = 0.21;   /* FIX: Standard auf 0,21 cm */
var apiOk = false;

/* Feste Diagramm-Position & Größe in RE */
var GANTT_LEFT  = 8;
var GANTT_TOP   = 17;
var GANTT_MAX_W = 118;
var GANTT_MAX_H = 69;

/* Zeilenhöhe & Balkenhöhe in RE */
var ROW_HEIGHT_RE = 4;
var BAR_HEIGHT_RE = 3;

/* Farb-Palette für Phasen */
var PHASE_COLORS = [
    "#2471A3", "#27AE60", "#8E44AD", "#E67E22",
    "#2980B9", "#1ABC9C", "#C0392B", "#D4AC0D",
    "#16A085", "#E74C3C", "#3498DB", "#9B59B6"
];

/* ══════════════════════════════════════════════════════════════
   KONVERTIERUNGEN
   ══════════════════════════════════════════════════════════════ */
function c2p(cm)  { return cm * 72 / 2.54; }
function p2c(pt)  { return pt * 2.54 / 72; }
function re2cm(re){ return re * gridUnitCm; }
function re2pt(re){ return c2p(re * gridUnitCm); }

/* ══════════════════════════════════════════════════════════════
   OFFICE READY
   ══════════════════════════════════════════════════════════════ */
Office.onReady(function (info) {
    if (info.host === Office.HostType.PowerPoint) {
        try {
            apiOk = Office.context.requirements.isSetSupported("PowerPointApi", "1.5");
        } catch (e) {
            apiOk = Office.context.requirements.isSetSupported("PowerPointApi", "1.1");
        }
    }
    initUI();
    setDefaultDates();
    addPhaseRow();
    addPhaseRow();
    addPhaseRow();
});

/* ══════════════════════════════════════════════════════════════
   UI INIT
   ══════════════════════════════════════════════════════════════ */
function initUI() {
    /* Rastereinheit Presets */
    document.querySelectorAll(".pre").forEach(function (b) {
        b.addEventListener("click", function () {
            gridUnitCm = parseFloat(this.dataset.v);
            document.querySelectorAll(".pre").forEach(function (p) { p.classList.remove("active"); });
            this.classList.add("active");
            showStatus("RE = " + gridUnitCm.toFixed(2) + " cm", "info");
        });
    });

    /* Position & Größe Inputs */
    var leftEl = document.getElementById("ganttLeft");
    var topEl  = document.getElementById("ganttTop");
    var maxWEl = document.getElementById("ganttMaxW");
    var maxHEl = document.getElementById("ganttMaxH");

    if (leftEl) leftEl.addEventListener("change", function () {
        var v = parseInt(this.value); if (!isNaN(v) && v >= 0) GANTT_LEFT = v;
    });
    if (topEl) topEl.addEventListener("change", function () {
        var v = parseInt(this.value); if (!isNaN(v) && v >= 0) GANTT_TOP = v;
    });
    if (maxWEl) maxWEl.addEventListener("change", function () {
        var v = parseInt(this.value); if (!isNaN(v) && v > 0) GANTT_MAX_W = v;
    });
    if (maxHEl) maxHEl.addEventListener("change", function () {
        var v = parseInt(this.value); if (!isNaN(v) && v > 0) GANTT_MAX_H = v;
    });

    /* Zeilenhöhe & Balkenhöhe Inputs */
    var rowHEl = document.getElementById("rowHeightRE");
    var barHEl = document.getElementById("barHeightRE");

    if (rowHEl) rowHEl.addEventListener("change", function () {
        var v = parseInt(this.value);
        if (!isNaN(v) && v >= 2) {
            ROW_HEIGHT_RE = v;
            if (BAR_HEIGHT_RE >= ROW_HEIGHT_RE) {
                BAR_HEIGHT_RE = ROW_HEIGHT_RE - 1;
                var be = document.getElementById("barHeightRE");
                if (be) be.value = BAR_HEIGHT_RE;
            }
        }
    });
    if (barHEl) barHEl.addEventListener("change", function () {
        var v = parseInt(this.value);
        if (!isNaN(v) && v >= 1) {
            if (v >= ROW_HEIGHT_RE) {
                showStatus("Balkenhöhe muss kleiner als Zeilenhöhe sein!", "warning");
                this.value = ROW_HEIGHT_RE - 1;
                BAR_HEIGHT_RE = ROW_HEIGHT_RE - 1;
            } else {
                BAR_HEIGHT_RE = v;
            }
        }
    });

    /* Phase Buttons */
    document.getElementById("addPhase").addEventListener("click", function () { addPhaseRow(); });
    document.getElementById("removePhase").addEventListener("click", function () { removeLastPhase(); });

    /* Generate Button */
    document.getElementById("generateGantt").addEventListener("click", function () { generateGantt(); });
}

/* ══════════════════════════════════════════════════════════════
   STATUS
   ══════════════════════════════════════════════════════════════ */
function showStatus(msg, type) {
    var el = document.getElementById("status");
    el.textContent = msg;
    el.className = "sts " + (type || "info");
}

/* ══════════════════════════════════════════════════════════════
   DEFAULT DATES
   ══════════════════════════════════════════════════════════════ */
function setDefaultDates() {
    var now   = new Date();
    var start = new Date(now.getFullYear(), now.getMonth(), 1);
    var end   = new Date(now.getFullYear(), now.getMonth() + 6, 0);
    document.getElementById("startDate").value = formatDate(start);
    document.getElementById("endDate").value   = formatDate(end);
}

function formatDate(d) {
    var m   = (d.getMonth() + 1).toString();
    if (m.length < 2) m = "0" + m;
    var day = d.getDate().toString();
    if (day.length < 2) day = "0" + day;
    return d.getFullYear() + "-" + m + "-" + day;
}

/* ══════════════════════════════════════════════════════════════
   PHASE MANAGEMENT
   ══════════════════════════════════════════════════════════════ */
var phaseCount = 0;

function addPhaseRow() {
    phaseCount++;
    var idx       = phaseCount;
    var container = document.getElementById("phaseContainer");
    var color     = PHASE_COLORS[(idx - 1) % PHASE_COLORS.length];

    var startD = document.getElementById("startDate").value;
    var endD   = document.getElementById("endDate").value;

    var div = document.createElement("div");
    div.className = "phase-row";
    div.id = "phase_" + idx;
    div.innerHTML =
        '<span class="phase-num">' + idx + '</span>' +
        '<input type="text" class="phase-name" placeholder="Phase ' + idx + '" value="Phase ' + idx + '"/>' +
        '<input type="date" class="phase-date phase-start" value="' + startD + '"/>' +
        '<input type="date" class="phase-date phase-end" value="' + endD + '"/>' +
        '<input type="color" class="phase-color" value="' + color + '"/>';
    container.appendChild(div);
}

function removeLastPhase() {
    if (phaseCount <= 1) { showStatus("Mind. 1 Phase nötig!", "warning"); return; }
    var el = document.getElementById("phase_" + phaseCount);
    if (el) el.remove();
    phaseCount--;
}

function getPhases() {
    var phases = [];
    for (var i = 1; i <= phaseCount; i++) {
        var row = document.getElementById("phase_" + i);
        if (!row) continue;
        var name  = row.querySelector(".phase-name").value  || ("Phase " + i);
        var start = row.querySelector(".phase-start").value;
        var end   = row.querySelector(".phase-end").value;
        var color = row.querySelector(".phase-color").value;
        if (!start || !end) continue;
        phases.push({
            name:  name,
            start: new Date(start),
            end:   new Date(end),
            color: color.replace("#", "")  /* FIX: # entfernen für setSolidColor */
        });
    }
    return phases;
}

/* ══════════════════════════════════════════════════════════════
   ZEITEINHEITEN-BERECHNUNG
   ══════════════════════════════════════════════════════════════ */
function getTimeSlots(startDate, endDate, unit) {
    var slots = [];
    var d;

    if (unit === "days") {
        d = new Date(startDate);
        while (d <= endDate) {
            var next = new Date(d);
            next.setDate(next.getDate() + 1);
            slots.push({
                label: d.getDate() + "." + (d.getMonth() + 1) + ".",
                start: new Date(d),
                end:   new Date(next)
            });
            d = next;
        }
    }
    else if (unit === "weeks") {
        d = new Date(startDate);
        var dow  = d.getDay();
        var diff = (dow === 0) ? -6 : 1 - dow;
        d.setDate(d.getDate() + diff);
        while (d <= endDate) {
            var wEnd = new Date(d);
            wEnd.setDate(wEnd.getDate() + 7);
            var kw = getISOWeek(d);
            slots.push({
                label: "KW" + kw,
                start: new Date(d),
                end:   new Date(wEnd)
            });
            d = wEnd;
        }
    }
    else if (unit === "months") {
        d = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
        while (d <= endDate) {
            var mEnd   = new Date(d.getFullYear(), d.getMonth() + 1, 1);
            var mNames = ["Jan","Feb","Mär","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez"];
            slots.push({
                label: mNames[d.getMonth()] + " " + d.getFullYear().toString().substr(2),
                start: new Date(d),
                end:   new Date(mEnd)
            });
            d = mEnd;
        }
    }
    else if (unit === "quarters") {
        var qStart = new Date(startDate.getFullYear(), Math.floor(startDate.getMonth() / 3) * 3, 1);
        d = qStart;
        while (d <= endDate) {
            var qEnd = new Date(d.getFullYear(), d.getMonth() + 3, 1);
            var qNum = Math.floor(d.getMonth() / 3) + 1;
            slots.push({
                label: "Q" + qNum + "/" + d.getFullYear().toString().substr(2),
                start: new Date(d),
                end:   new Date(qEnd)
            });
            d = qEnd;
        }
    }

    return slots;
}

function getISOWeek(date) {
    var d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
    var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    var weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    return weekNo;
}

/* ══════════════════════════════════════════════════════════════
   GENERATE GANTT
   ══════════════════════════════════════════════════════════════ */
function generateGantt() {
    var startDate = new Date(document.getElementById("startDate").value);
    var endDate   = new Date(document.getElementById("endDate").value);
    var unit      = document.getElementById("timeUnit").value;

    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
        showStatus("Bitte Start- und Enddatum eingeben!", "error");
        return;
    }
    if (endDate <= startDate) {
        showStatus("Enddatum muss nach Startdatum liegen!", "error");
        return;
    }

    var phases = getPhases();
    if (phases.length === 0) {
        showStatus("Bitte mindestens eine Phase anlegen!", "error");
        return;
    }

    var timeSlots = getTimeSlots(startDate, endDate, unit);
    if (timeSlots.length === 0) {
        showStatus("Keine Zeiteinheiten berechenbar!", "error");
        return;
    }

    var showToday     = document.getElementById("showToday").checked;
    var showHeader    = document.getElementById("showHeader").checked;
    var showLabels    = document.getElementById("showLabels").checked;
    var showGridLines = document.getElementById("showGridLines").checked;

    /* Position & Größe aus Inputs */
    var leftEl = document.getElementById("ganttLeft");
    var topEl  = document.getElementById("ganttTop");
    var maxWEl = document.getElementById("ganttMaxW");
    var maxHEl = document.getElementById("ganttMaxH");
    if (leftEl) GANTT_LEFT = parseInt(leftEl.value) || GANTT_LEFT;
    if (topEl)  GANTT_TOP  = parseInt(topEl.value)  || GANTT_TOP;
    if (maxWEl) GANTT_MAX_W = parseInt(maxWEl.value) || GANTT_MAX_W;
    if (maxHEl) GANTT_MAX_H = parseInt(maxHEl.value) || GANTT_MAX_H;

    /* Zeilen-/Balkenhöhe aus Inputs */
    ROW_HEIGHT_RE = parseInt(document.getElementById("rowHeightRE").value) || 4;
    BAR_HEIGHT_RE = parseInt(document.getElementById("barHeightRE").value) || 3;
    if (BAR_HEIGHT_RE >= ROW_HEIGHT_RE) BAR_HEIGHT_RE = ROW_HEIGHT_RE - 1;
    if (BAR_HEIGHT_RE < 1) BAR_HEIGHT_RE = 1;

    /* Layout-Berechnung in RE (ganzzahlig) */
    var labelWidthRE = showLabels ? Math.floor(GANTT_MAX_W * 0.15) : 0;
    if (labelWidthRE < 8 && showLabels) labelWidthRE = 8;

    var chartWidthRE   = GANTT_MAX_W - labelWidthRE;
    var headerHeightRE = showHeader ? 3 : 0;

    /* Spaltenbreite in RE – mindestens 1 RE pro Slot */
    var colWidthRE = Math.max(1, Math.floor(chartWidthRE / timeSlots.length));
    var actualChartWidthRE = colWidthRE * timeSlots.length;
    if (actualChartWidthRE > chartWidthRE) {
        colWidthRE = Math.floor(chartWidthRE / timeSlots.length);
        actualChartWidthRE = colWidthRE * timeSlots.length;
    }

    var rowHeightRE = ROW_HEIGHT_RE;
    var barHeightRE = BAR_HEIGHT_RE;

    var totalHeightRE = headerHeightRE + (rowHeightRE * phases.length);
    if (totalHeightRE > GANTT_MAX_H) {
        showStatus("Warnung: GANTT (" + totalHeightRE + " RE) größer als Max (" + GANTT_MAX_H + " RE)!", "warning");
    }

    showStatus("Erstelle GANTT... " + timeSlots.length + " Spalten, " + phases.length + " Phasen", "info");

    /* PowerPoint Shapes erzeugen */
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
                    if (slides.items.length === 0) {
                        showStatus("Keine Folie gefunden!", "error");
                        return ctx.sync();
                    }
                    slide = slides.items[0];
                    return buildGantt(ctx, slide, timeSlots, phases, {
                        startDate: startDate, endDate: endDate, unit: unit,
                        labelWidthRE: labelWidthRE, chartWidthRE: actualChartWidthRE,
                        headerHeightRE: headerHeightRE, rowHeightRE: rowHeightRE,
                        barHeightRE: barHeightRE, colWidthRE: colWidthRE,
                        showToday: showToday, showHeader: showHeader,
                        showLabels: showLabels, showGridLines: showGridLines
                    });
                });
            }
            return buildGantt(ctx, slide, timeSlots, phases, {
                startDate: startDate, endDate: endDate, unit: unit,
                labelWidthRE: labelWidthRE, chartWidthRE: actualChartWidthRE,
                headerHeightRE: headerHeightRE, rowHeightRE: rowHeightRE,
                barHeightRE: barHeightRE, colWidthRE: colWidthRE,
                showToday: showToday, showHeader: showHeader,
                showLabels: showLabels, showGridLines: showGridLines
            });
        });
    }).catch(function (e) {
        showStatus("Fehler: " + e.message, "error");
    });
}

/* ══════════════════════════════════════════════════════════════
   BUILD GANTT – Shapes auf Folie erzeugen
   ══════════════════════════════════════════════════════════════ */
function buildGantt(ctx, slide, timeSlots, phases, cfg) {
    var x0 = GANTT_LEFT;
    var y0 = GANTT_TOP;

    var chartX0 = x0 + cfg.labelWidthRE;
    var chartY0 = y0 + cfg.headerHeightRE;

    var totalMs = cfg.endDate.getTime() - cfg.startDate.getTime();

    /* ────────────────────────────────────────────
       1) HINTERGRUND
       ──────────────────────────────────────────── */
    var bgW = cfg.labelWidthRE + cfg.chartWidthRE;
    var bgH = cfg.headerHeightRE + (cfg.rowHeightRE * phases.length);
    var bg  = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    bg.left   = re2pt(x0);
    bg.top    = re2pt(y0);
    bg.width  = re2pt(bgW);
    bg.height = re2pt(bgH);
    bg.fill.setSolidColor("F5F5F5");
    bg.lineFormat.visible = false;
    bg.name = "GANTT_BG";

    /* ────────────────────────────────────────────
       2) HEADER-ZEILE (Zeiteinheiten)
       ──────────────────────────────────────────── */
    if (cfg.showHeader) {
        /* Header-Hintergrund */
        var hdr = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
        hdr.left   = re2pt(chartX0);
        hdr.top    = re2pt(y0);
        hdr.width  = re2pt(cfg.chartWidthRE);
        hdr.height = re2pt(cfg.headerHeightRE);
        hdr.fill.setSolidColor("1A1A2E");
        hdr.lineFormat.visible = false;
        hdr.name = "GANTT_HDR_BG";

        /* Header Labels */
        for (var h = 0; h < timeSlots.length; h++) {
            var hx  = chartX0 + (h * cfg.colWidthRE);
            var htb = slide.shapes.addTextBox(timeSlots[h].label, {
                left:   re2pt(hx),
                top:    re2pt(y0),
                width:  re2pt(cfg.colWidthRE),
                height: re2pt(cfg.headerHeightRE)
            });
            htb.name = "GANTT_HDR_" + h;
            htb.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
            htb.textFrame.textRange.font.size  = 7;
            htb.textFrame.textRange.font.color = "FFFFFF";
            htb.textFrame.textRange.font.bold  = true;
            htb.textFrame.textRange.paragraphFormat.horizontalAlignment =
                PowerPoint.ParagraphHorizontalAlignment.center;
            htb.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;
            htb.fill.clear();
            htb.lineFormat.visible = false;
        }
    }

    /* ────────────────────────────────────────────
       3) LABELS links (Phasennamen)
       ──────────────────────────────────────────── */
    if (cfg.showLabels) {
        for (var l = 0; l < phases.length; l++) {
            var ly  = chartY0 + (l * cfg.rowHeightRE);
            var ltb = slide.shapes.addTextBox(phases[l].name, {
                left:   re2pt(x0),
                top:    re2pt(ly),
                width:  re2pt(cfg.labelWidthRE),
                height: re2pt(cfg.rowHeightRE)
            });
            ltb.name = "GANTT_LBL_" + l;
            ltb.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
            ltb.textFrame.textRange.font.size  = 8;
            ltb.textFrame.textRange.font.color = "1A1A2E";
            ltb.textFrame.textRange.font.bold  = true;
            ltb.textFrame.textRange.paragraphFormat.horizontalAlignment =
                PowerPoint.ParagraphHorizontalAlignment.left;
            ltb.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;
            ltb.fill.clear();
            ltb.lineFormat.visible = false;
        }
    }

    /* ────────────────────────────────────────────
       4) ZEILEN-HINTERGRUND (abwechselnd) → ENTFERNT
       ──────────────────────────────────────────── */

    /* ────────────────────────────────────────────
       5) VERTIKALE RASTERLINIEN
       ──────────────────────────────────────────── */
    if (cfg.showGridLines) {
        var gridH = cfg.headerHeightRE + (cfg.rowHeightRE * phases.length);
        for (var g = 0; g <= timeSlots.length; g++) {
            var gx = chartX0 + (g * cfg.colWidthRE);
            var gl = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
            gl.left   = re2pt(gx);
            gl.top    = re2pt(y0);
            gl.width  = c2p(0.02);
            gl.height = re2pt(gridH);
            gl.fill.setSolidColor("B0B0B0");
            gl.lineFormat.visible = false;
            gl.name = "GANTT_GRID_" + g;
        }
    }

    /* ────────────────────────────────────────────
       6) GANTT-BALKEN (Phasen)
       ──────────────────────────────────────────── */
    for (var p = 0; p < phases.length; p++) {
        var phase = phases[p];

        /* Zeitbereich clippen */
        var pStart = Math.max(phase.start.getTime(), cfg.startDate.getTime());
        var pEnd   = Math.min(phase.end.getTime(),   cfg.endDate.getTime());
        if (pEnd <= pStart) continue;

        /* Position in RE berechnen (ganzzahlig) */
        var startRatio = (pStart - cfg.startDate.getTime()) / totalMs;
        var endRatio   = (pEnd   - cfg.startDate.getTime()) / totalMs;

        var barLeftRE  = Math.round(startRatio * cfg.chartWidthRE);
        var barRightRE = Math.round(endRatio   * cfg.chartWidthRE);
        var barWidthRE = Math.max(1, barRightRE - barLeftRE);

        /* Vertikal: zentriert in der Zeile */
        var rowY       = chartY0 + (p * cfg.rowHeightRE);
        var barYOffset = Math.floor((cfg.rowHeightRE - cfg.barHeightRE) / 2);
        var barY       = rowY + barYOffset;

        /* Balken-Shape mit Phasenfarbe
           FIX: phase.color ist bereits ohne # (durch getPhases) */
        var bar = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundedRectangle);
        bar.left   = re2pt(chartX0 + barLeftRE);
        bar.top    = re2pt(barY);
        bar.width  = re2pt(barWidthRE);
        bar.height = re2pt(cfg.barHeightRE);
        bar.fill.setSolidColor(phase.color);
        bar.lineFormat.visible = false;
        bar.name = "GANTT_BAR_" + p;

        /* Balken-Text (Phasenname auf dem Balken) */
        if (barWidthRE >= 4) {
            var barTb = slide.shapes.addTextBox(phase.name, {
                left:   re2pt(chartX0 + barLeftRE),
                top:    re2pt(barY),
                width:  re2pt(barWidthRE),
                height: re2pt(cfg.barHeightRE)
            });
            barTb.name = "GANTT_BARTXT_" + p;
            barTb.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
            barTb.textFrame.textRange.font.size  = 7;
            barTb.textFrame.textRange.font.color = "FFFFFF";
            barTb.textFrame.textRange.font.bold  = true;
            barTb.textFrame.textRange.paragraphFormat.horizontalAlignment =
                PowerPoint.ParagraphHorizontalAlignment.center;
            barTb.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;
            barTb.fill.clear();
            barTb.lineFormat.visible = false;
        }
    }

    /* ────────────────────────────────────────────
       7) HEUTE-LINIE (rot, vertikal)
       ──────────────────────────────────────────── */
    if (cfg.showToday) {
        var today = new Date();
        today.setHours(0, 0, 0, 0);
        if (today >= cfg.startDate && today <= cfg.endDate) {
            var todayRatio = (today.getTime() - cfg.startDate.getTime()) / totalMs;
            var todayRE    = Math.round(todayRatio * cfg.chartWidthRE);
            var todayX     = chartX0 + todayRE;
            var totalH     = cfg.headerHeightRE + (cfg.rowHeightRE * phases.length);

            /* Rote Linie */
            var tLine = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
            tLine.left   = re2pt(todayX);
            tLine.top    = re2pt(y0);
            tLine.width  = c2p(0.06);
            tLine.height = re2pt(totalH);
            tLine.fill.setSolidColor("FF0000");
            tLine.lineFormat.visible = false;
            tLine.name = "GANTT_TODAY";

            /* "Heute" Label oben */
            var tLbl = slide.shapes.addTextBox("\u25BC Heute", {
                left:   re2pt(todayX) - c2p(0.5),
                top:    re2pt(y0) - c2p(0.5),
                width:  c2p(1.5),
                height: c2p(0.5)
            });
            tLbl.name = "GANTT_TODAY_LBL";
            tLbl.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
            tLbl.textFrame.textRange.font.size  = 6;
            tLbl.textFrame.textRange.font.color = "FF0000";
            tLbl.textFrame.textRange.font.bold  = true;
            tLbl.textFrame.textRange.paragraphFormat.horizontalAlignment =
                PowerPoint.ParagraphHorizontalAlignment.center;
            tLbl.fill.clear();
            tLbl.lineFormat.visible = false;
        }
    }

    showStatus("GANTT erstellt! (" + phases.length + " Phasen, " + timeSlots.length + " Spalten)", "success");
    return ctx.sync();
}
