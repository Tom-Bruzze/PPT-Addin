/* ═══════════════════════════════════════════════════════════════
   DROEGE GANTT Generator – taskpane.js
   ═══════════════════════════════════════════════════════════════
   Erzeugt GANTT-Diagramme direkt in PowerPoint als Shapes.
   Alle Maße in Rastereinheiten (RE), Position immer ganzzahlig.
   ═══════════════════════════════════════════════════════════════ */

/* ── Globals ── */
var gridUnitCm = 0.63;
var apiOk = false;

/* Feste Diagramm-Position & Größe in RE */
var GANTT_LEFT   = 8;
var GANTT_TOP    = 17;
var GANTT_MAX_W  = 118;
var GANTT_MAX_H  = 69;

/* Farb-Palette für Phasen */
var PHASE_COLORS = [
    "#2471A3", "#27AE60", "#8E44AD", "#E67E22",
    "#2980B9", "#1ABC9C", "#C0392B", "#D4AC0D",
    "#16A085", "#E74C3C", "#3498DB", "#9B59B6"
];

/* ── Konvertierungen ── */
function c2p(cm) { return cm * 72 / 2.54; }
function p2c(pt) { return pt * 2.54 / 72; }
function re2cm(re) { return re * gridUnitCm; }
function re2pt(re) { return c2p(re * gridUnitCm); }

/* ── Office Ready ── */
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
            document.getElementById("reInput").value = gridUnitCm;
            document.querySelectorAll(".pre").forEach(function (p) { p.classList.remove("active"); });
            this.classList.add("active");
            showStatus("RE = " + gridUnitCm.toFixed(2) + " cm", "info");
        });
    });

    /* RE manuell */
    var rei = document.getElementById("reInput");
    rei.addEventListener("change", function () {
        var v = parseFloat(this.value);
        if (!isNaN(v) && v > 0) {
            gridUnitCm = v;
            document.querySelectorAll(".pre").forEach(function (p) {
                p.classList.toggle("active", parseFloat(p.dataset.v) === v);
            });
            showStatus("RE = " + gridUnitCm.toFixed(2) + " cm", "info");
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
    var now = new Date();
    var start = new Date(now.getFullYear(), now.getMonth(), 1);
    var end = new Date(now.getFullYear(), now.getMonth() + 6, 0);
    document.getElementById("startDate").value = formatDate(start);
    document.getElementById("endDate").value = formatDate(end);
}

function formatDate(d) {
    var m = (d.getMonth() + 1).toString();
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
    var idx = phaseCount;
    var container = document.getElementById("phaseContainer");
    var color = PHASE_COLORS[(idx - 1) % PHASE_COLORS.length];

    var startD = document.getElementById("startDate").value;
    var endD = document.getElementById("endDate").value;

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
        var name  = row.querySelector(".phase-name").value || ("Phase " + i);
        var start = row.querySelector(".phase-start").value;
        var end   = row.querySelector(".phase-end").value;
        var color = row.querySelector(".phase-color").value;
        if (!start || !end) continue;
        phases.push({
            name:  name,
            start: new Date(start),
            end:   new Date(end),
            color: color.replace("#", "")
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
        /* Start auf Montag der Startwoche setzen */
        d = new Date(startDate);
        var dow = d.getDay();
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
            var mEnd = new Date(d.getFullYear(), d.getMonth() + 1, 1);
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

function getISOWeek(d) {
    var date = new Date(d.getTime());
    date.setHours(0, 0, 0, 0);
    date.setDate(date.getDate() + 3 - (date.getDay() + 6) % 7);
    var week1 = new Date(date.getFullYear(), 0, 4);
    return 1 + Math.round(((date - week1) / 86400000 - 3 + (week1.getDay() + 6) % 7) / 7);
}

/* ══════════════════════════════════════════════════════════════
   GANTT GENERATOR – Hauptfunktion
   ══════════════════════════════════════════════════════════════
   Alles in ganzen Rastereinheiten (RE).
   Position: Links=8 RE, Oben=17 RE
   Max: Breite=118 RE, Höhe=69 RE
   ══════════════════════════════════════════════════════════════ */
function generateGantt() {
    /* ── Eingaben lesen ── */
    var startDate = new Date(document.getElementById("startDate").value);
    var endDate   = new Date(document.getElementById("endDate").value);
    var unit      = document.getElementById("timeUnit").value;

    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
        showStatus("Bitte Start- und End-Datum angeben!", "error"); return;
    }
    if (endDate <= startDate) {
        showStatus("End-Datum muss nach Start-Datum liegen!", "error"); return;
    }

    var phases    = getPhases();
    if (phases.length === 0) { showStatus("Mind. 1 Phase mit Daten nötig!", "error"); return; }

    var showToday     = document.getElementById("showToday").checked;
    var showHeader    = document.getElementById("showHeader").checked;
    var showLabels    = document.getElementById("showLabels").checked;
    var showGridLines = document.getElementById("showGridLines").checked;

    var timeSlots = getTimeSlots(startDate, endDate, unit);
    if (timeSlots.length === 0) { showStatus("Keine Zeiteinheiten im Bereich!", "error"); return; }
    if (timeSlots.length > 120) { showStatus("Zu viele Zeiteinheiten (max 120)!", "error"); return; }

    /* ── Layout-Berechnung in RE (ganzzahlig) ── */
    var labelWidthRE = showLabels ? Math.floor(GANTT_MAX_W * 0.15) : 0; /* ~15% für Labels */
    if (labelWidthRE < 8 && showLabels) labelWidthRE = 8;

    var chartWidthRE  = GANTT_MAX_W - labelWidthRE;
    var headerHeightRE = showHeader ? 3 : 0;

    /* Spaltenbreite in RE – mindestens 1 RE pro Slot */
    var colWidthRE = Math.max(1, Math.floor(chartWidthRE / timeSlots.length));
    /* Tatsächliche Chart-Breite anpassen */
    var actualChartWidthRE = colWidthRE * timeSlots.length;
    if (actualChartWidthRE > chartWidthRE) {
        colWidthRE = Math.floor(chartWidthRE / timeSlots.length);
        actualChartWidthRE = colWidthRE * timeSlots.length;
    }

    /* Zeilenhöhe in RE – jede Phase 1 RE kleiner Abstand zur vorherigen */
    var availHeightRE = GANTT_MAX_H - headerHeightRE;
    var rowHeightRE = Math.max(2, Math.floor(availHeightRE / (phases.length + 1)));
    if (rowHeightRE > 6) rowHeightRE = 6;
    var barHeightRE = Math.max(1, rowHeightRE - 1); /* 1 RE kleiner */

    var totalHeightRE = headerHeightRE + (rowHeightRE * phases.length);
    if (totalHeightRE > GANTT_MAX_H) {
        rowHeightRE = Math.floor((GANTT_MAX_H - headerHeightRE) / phases.length);
        barHeightRE = Math.max(1, rowHeightRE - 1);
        totalHeightRE = headerHeightRE + (rowHeightRE * phases.length);
    }

    showStatus("Erstelle GANTT... " + timeSlots.length + " Spalten, " + phases.length + " Phasen", "info");

    /* ── PowerPoint Shapes erzeugen ── */
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
    var x0 = GANTT_LEFT;  /* Start links in RE */
    var y0 = GANTT_TOP;   /* Start oben in RE */

    var chartX0 = x0 + cfg.labelWidthRE;     /* Chart-Bereich Start X */
    var chartY0 = y0 + cfg.headerHeightRE;   /* Chart-Bereich Start Y */

    var totalMs = cfg.endDate.getTime() - cfg.startDate.getTime();

    /* ────────────────────────────────────────────
       1) HINTERGRUND
       ──────────────────────────────────────────── */
    var bgW = cfg.labelWidthRE + cfg.chartWidthRE;
    var bgH = cfg.headerHeightRE + (cfg.rowHeightRE * phases.length);
    var bg = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
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
            var hx = chartX0 + (h * cfg.colWidthRE);
            var htb = slide.shapes.addTextBox();
            htb.left   = re2pt(hx);
            htb.top    = re2pt(y0);
            htb.width  = re2pt(cfg.colWidthRE);
            htb.height = re2pt(cfg.headerHeightRE);
            htb.name = "GANTT_HDR_" + h;
            htb.textFrame.autoSizeSetting = "autoSizeNone";

            var htf = htb.textFrame.getRange();
            htf.text = timeSlots[h].label;
            htf.font.size = 7;
            htf.font.color = "FFFFFF";
            htf.font.bold = true;
            htf.paragraphFormat.alignment = "Center";
        }
    }

    /* ────────────────────────────────────────────
       3) LABELS links (Phasennamen)
       ──────────────────────────────────────────── */
    if (cfg.showLabels) {
        for (var l = 0; l < phases.length; l++) {
            var ly = chartY0 + (l * cfg.rowHeightRE);
            var ltb = slide.shapes.addTextBox();
            ltb.left   = re2pt(x0);
            ltb.top    = re2pt(ly);
            ltb.width  = re2pt(cfg.labelWidthRE);
            ltb.height = re2pt(cfg.rowHeightRE);
            ltb.name = "GANTT_LBL_" + l;
            ltb.textFrame.autoSizeSetting = "autoSizeNone";

            var ltf = ltb.textFrame.getRange();
            ltf.text = phases[l].name;
            ltf.font.size = 8;
            ltf.font.color = "1A1A2E";
            ltf.font.bold = true;
            ltf.paragraphFormat.alignment = "Left";
        }
    }

    /* ────────────────────────────────────────────
       4) ZEILEN-HINTERGRUND (abwechselnd)
       ──────────────────────────────────────────── */
    for (var r = 0; r < phases.length; r++) {
        var ry = chartY0 + (r * cfg.rowHeightRE);
        var rowBg = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
        rowBg.left   = re2pt(chartX0);
        rowBg.top    = re2pt(ry);
        rowBg.width  = re2pt(cfg.chartWidthRE);
        rowBg.height = re2pt(cfg.rowHeightRE);
        rowBg.fill.setSolidColor(r % 2 === 0 ? "FFFFFF" : "F0F0F0");
        rowBg.lineFormat.visible = false;
        rowBg.name = "GANTT_ROWBG_" + r;
    }

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
            gl.width  = c2p(0.02);   /* Haarline */
            gl.height = re2pt(gridH);
            gl.fill.setSolidColor("B0B0B0");
            gl.lineFormat.visible = false;
            gl.name = "GANTT_GRID_" + g;
        }
    }

    /* ────────────────────────────────────────────
       6) GANTT-BALKEN (Phasen)
       ──────────────────────────────────────────── 
       Balkenhöhe = rowHeightRE - 1 RE (1 RE kleiner)
       Balken vertikal zentriert in der Zeile.
       Horizontale Position: proportional zur Zeit,
       aber auf ganze RE gerundet. */
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

        /* Vertikal: zentriert in der Zeile, 1 RE kleiner */
        var rowY = chartY0 + (p * cfg.rowHeightRE);
        var barYOffset = Math.floor((cfg.rowHeightRE - cfg.barHeightRE) / 2);
        var barY = rowY + barYOffset;

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
            var barTb = slide.shapes.addTextBox();
            barTb.left   = re2pt(chartX0 + barLeftRE);
            barTb.top    = re2pt(barY);
            barTb.width  = re2pt(barWidthRE);
            barTb.height = re2pt(cfg.barHeightRE);
            barTb.name = "GANTT_BARTXT_" + p;
            barTb.textFrame.autoSizeSetting = "autoSizeNone";

            var btf = barTb.textFrame.getRange();
            btf.text = phase.name;
            btf.font.size = 7;
            btf.font.color = "FFFFFF";
            btf.font.bold = true;
            btf.paragraphFormat.alignment = "Center";
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
            var todayRE = Math.round(todayRatio * cfg.chartWidthRE);
            var todayX = chartX0 + todayRE;
            var totalH = cfg.headerHeightRE + (cfg.rowHeightRE * phases.length);

            /* Rote Linie */
            var tLine = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
            tLine.left   = re2pt(todayX);
            tLine.top    = re2pt(y0);
            tLine.width  = c2p(0.06);   /* Etwas dicker */
            tLine.height = re2pt(totalH);
            tLine.fill.setSolidColor("FF0000");
            tLine.lineFormat.visible = false;
            tLine.name = "GANTT_TODAY";

            /* "Heute" Label oben */
            var tLbl = slide.shapes.addTextBox();
            tLbl.left   = re2pt(todayX) - c2p(0.5);
            tLbl.top    = re2pt(y0) - c2p(0.5);
            tLbl.width  = c2p(1.5);
            tLbl.height = c2p(0.5);
            tLbl.name = "GANTT_TODAY_LBL";
            tLbl.textFrame.autoSizeSetting = "autoSizeNone";

            var tlf = tLbl.textFrame.getRange();
            tlf.text = "▼ Heute";
            tlf.font.size = 6;
            tlf.font.color = "FF0000";
            tlf.font.bold = true;
            tlf.paragraphFormat.alignment = "Center";
        }
    }

    /* ────────────────────────────────────────────
       8) SYNC & STATUS
       ──────────────────────────────────────────── */
    return ctx.sync().then(function () {
        var info = phases.length + " Phasen · " + timeSlots.length + " " +
            (cfg.unit === "days" ? "Tage" :
             cfg.unit === "weeks" ? "KW" :
             cfg.unit === "months" ? "Monate" : "Quartale");
        showStatus("GANTT erstellt ✓ · " + info, "success");

        /* Info-Box anzeigen */
        var el = document.getElementById("infoBox");
        el.innerHTML =
            '<div class="info-item"><span class="info-label">Position:</span>' +
            '<span class="info-value">' + GANTT_LEFT + ' × ' + GANTT_TOP + ' RE</span></div>' +
            '<div class="info-item"><span class="info-label">Größe:</span>' +
            '<span class="info-value">' + (cfg.labelWidthRE + cfg.chartWidthRE) + ' × ' +
            (cfg.headerHeightRE + cfg.rowHeightRE * phases.length) + ' RE</span></div>' +
            '<div class="info-item"><span class="info-label">Spaltenbreite:</span>' +
            '<span class="info-value">' + cfg.colWidthRE + ' RE</span></div>' +
            '<div class="info-item"><span class="info-label">Zeilenhöhe:</span>' +
            '<span class="info-value">' + cfg.rowHeightRE + ' RE (Balken: ' + cfg.barHeightRE + ' RE)</span></div>';
        el.classList.add("visible");
    });
}
