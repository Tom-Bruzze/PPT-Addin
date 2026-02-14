/*
  DROEGE GANTT Generator – taskpane.js
  Build 14.02.2026 v2
  Farb-Zuweisung pro Phase via Swatch + Hex-Eingabe
*/

/* ── GLOBALS ── */
var gridUnitCm = 0.21;
var apiOk = false;

var GANTT_LEFT   = 8;
var GANTT_TOP    = 17;
var GANTT_MAX_W  = 118;
var GANTT_MAX_H  = 69;
var ROW_HEIGHT_RE = 4;
var BAR_HEIGHT_RE = 3;

/* Farb-Palette (12 unterscheidbare Farben) */
var PHASE_COLORS = [
    "#2471A3", "#27AE60", "#8E44AD", "#E67E22",
    "#2980B9", "#1ABC9C", "#C0392B", "#D4AC0D",
    "#16A085", "#E74C3C", "#3498DB", "#F39C12"
];

var phaseCount = 0;

/* ── INIT ── */
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

/* ── UI INIT ── */
function initUI() {
    /* RE Preset Buttons */
    var preButtons = document.querySelectorAll(".pre");
    for (var i = 0; i < preButtons.length; i++) {
        preButtons[i].addEventListener("click", function () {
            for (var j = 0; j < preButtons.length; j++) preButtons[j].classList.remove("active");
            this.classList.add("active");
            gridUnitCm = parseFloat(this.getAttribute("data-v"));
            showStatus("RE = " + gridUnitCm + " cm", "info");
        });
    }

    /* Position & Größe Inputs */
    var leftEl = document.getElementById("ganttLeft");
    var topEl  = document.getElementById("ganttTop");
    var maxWEl = document.getElementById("ganttMaxW");
    var maxHEl = document.getElementById("ganttMaxH");
    var rowHEl = document.getElementById("rowHeightRE");
    var barHEl = document.getElementById("barHeightRE");

    if (leftEl) leftEl.addEventListener("change", function () {
        var v = parseInt(this.value); if (!isNaN(v) && v >= 0) GANTT_LEFT = v;
    });
    if (topEl) topEl.addEventListener("change", function () {
        var v = parseInt(this.value); if (!isNaN(v) && v >= 0) GANTT_TOP = v;
    });
    if (maxWEl) maxWEl.addEventListener("change", function () {
        var v = parseInt(this.value); if (!isNaN(v) && v >= 10) GANTT_MAX_W = v;
    });
    if (maxHEl) maxHEl.addEventListener("change", function () {
        var v = parseInt(this.value); if (!isNaN(v) && v >= 10) GANTT_MAX_H = v;
    });
    if (rowHEl) rowHEl.addEventListener("change", function () {
        var v = parseInt(this.value);
        if (!isNaN(v) && v >= 2) {
            ROW_HEIGHT_RE = v;
            if (BAR_HEIGHT_RE >= v) {
                BAR_HEIGHT_RE = v - 1;
                var be = document.getElementById("barHeightRE");
                if (be) be.value = BAR_HEIGHT_RE;
            }
        }
    });
    if (barHEl) barHEl.addEventListener("change", function () {
        var v = parseInt(this.value);
        if (!isNaN(v) && v >= 1) {
            if (v >= ROW_HEIGHT_RE) {
                showStatus("Balkenhöhe muss kleiner als Zeilenhöhe sein", "warning");
                this.value = BAR_HEIGHT_RE;
            } else {
                BAR_HEIGHT_RE = v;
            }
        }
    });

    /* Phase Buttons */
    document.getElementById("addPhase").addEventListener("click", addPhaseRow);
    document.getElementById("removePhase").addEventListener("click", removeLastPhase);
    document.getElementById("generateGantt").addEventListener("click", generateGantt);
}

/* ── STATUS ── */
function showStatus(msg, type) {
    var el = document.getElementById("status");
    el.textContent = msg;
    el.className = "sts " + (type || "info");
}

/* ── DATES ── */
function setDefaultDates() {
    var now = new Date();
    var start = new Date(now.getFullYear(), now.getMonth(), 1);
    var end   = new Date(now.getFullYear(), now.getMonth() + 6, 0);
    document.getElementById("startDate").value = fmtDate(start);
    document.getElementById("endDate").value   = fmtDate(end);
}

function fmtDate(d) {
    var m = (d.getMonth() + 1).toString(); if (m.length < 2) m = "0" + m;
    var day = d.getDate().toString();       if (day.length < 2) day = "0" + day;
    return d.getFullYear() + "-" + m + "-" + day;
}

/* ═══════════════════════════════════════════════════════════
   PHASE ROWS – mit Swatch + Hex-Input + native Color-Picker
   ═══════════════════════════════════════════════════════════ */

function normalizeHex(raw) {
    /* Eingabe: "#abc", "abc", "#AABBCC", "aabbcc", etc.
       Ausgabe: 6-stellig UPPERCASE ohne # */
    var hex = raw.replace(/^#/, "").toUpperCase();
    if (hex.length === 3) hex = hex[0]+hex[0] + hex[1]+hex[1] + hex[2]+hex[2];
    if (!/^[0-9A-F]{6}$/.test(hex)) return null;
    return hex;
}

function addPhaseRow() {
    phaseCount++;
    var idx = phaseCount;
    var container = document.getElementById("phaseContainer");
    var defaultColor = PHASE_COLORS[(idx - 1) % PHASE_COLORS.length];
    var startD = document.getElementById("startDate").value;
    var endD   = document.getElementById("endDate").value;

    /* Zeile */
    var div = document.createElement("div");
    div.className = "phase-row";
    div.id = "phase_" + idx;

    /* Nummer-Badge */
    var num = document.createElement("span");
    num.className = "phase-num";
    num.textContent = idx;
    div.appendChild(num);

    /* Felder-Container */
    var fields = document.createElement("div");
    fields.className = "phase-fields";

    /* Zeile 1: Name + Farbe */
    var row1 = document.createElement("div");
    row1.className = "phase-field-row";

    var nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.className = "phase-name";
    nameInput.placeholder = "Phase " + idx;
    nameInput.value = "Phase " + idx;
    row1.appendChild(nameInput);

    /* Farb-Wrapper: Swatch (klickbar) + versteckter Color-Picker + Hex-Input */
    var colorWrap = document.createElement("div");
    colorWrap.className = "phase-color-wrap";

    /* Swatch (visuelles Farbquadrat) */
    var swatch = document.createElement("div");
    swatch.className = "phase-swatch";
    swatch.style.backgroundColor = defaultColor;
    swatch.title = "Klicken für Farbwahl";
    colorWrap.appendChild(swatch);

    /* Nativer Color-Picker (versteckt, wird via Swatch-Klick geöffnet) */
    var colorPicker = document.createElement("input");
    colorPicker.type = "color";
    colorPicker.className = "phase-color";
    colorPicker.value = defaultColor;
    colorWrap.appendChild(colorPicker);

    /* Hex-Eingabe */
    var hexInput = document.createElement("input");
    hexInput.type = "text";
    hexInput.className = "phase-hex";
    hexInput.value = defaultColor.replace("#", "");
    hexInput.maxLength = 6;
    hexInput.placeholder = "HEX";
    colorWrap.appendChild(hexInput);

    row1.appendChild(colorWrap);
    fields.appendChild(row1);

    /* Zeile 2: Start + End Date */
    var row2 = document.createElement("div");
    row2.className = "phase-field-row";

    var startInput = document.createElement("input");
    startInput.type = "date";
    startInput.className = "phase-date phase-start";
    startInput.value = startD;
    row2.appendChild(startInput);

    var endInput = document.createElement("input");
    endInput.type = "date";
    endInput.className = "phase-date phase-end";
    endInput.value = endD;
    row2.appendChild(endInput);

    fields.appendChild(row2);
    div.appendChild(fields);
    container.appendChild(div);

    /* ── EVENT: Swatch klicken → Color-Picker öffnen ── */
    swatch.addEventListener("click", function () {
        colorPicker.click();
    });

    /* ── EVENT: Color-Picker ändert → Swatch + Hex synchronisieren ── */
    colorPicker.addEventListener("input", function () {
        var c = this.value;
        swatch.style.backgroundColor = c;
        hexInput.value = c.replace("#", "").toUpperCase();
    });
    colorPicker.addEventListener("change", function () {
        var c = this.value;
        swatch.style.backgroundColor = c;
        hexInput.value = c.replace("#", "").toUpperCase();
    });

    /* ── EVENT: Hex-Input ändert → Swatch + Color-Picker synchronisieren ── */
    hexInput.addEventListener("input", function () {
        var norm = normalizeHex(this.value);
        if (norm) {
            swatch.style.backgroundColor = "#" + norm;
            colorPicker.value = "#" + norm.toLowerCase();
        }
    });
    hexInput.addEventListener("change", function () {
        var norm = normalizeHex(this.value);
        if (norm) {
            this.value = norm;
            swatch.style.backgroundColor = "#" + norm;
            colorPicker.value = "#" + norm.toLowerCase();
        } else {
            /* Ungültig → Reset auf aktuelle Picker-Farbe */
            var cur = colorPicker.value.replace("#", "").toUpperCase();
            this.value = cur;
            showStatus("Ungültiger Hex-Wert, zurückgesetzt", "warning");
        }
    });
}

function removeLastPhase() {
    if (phaseCount <= 1) { showStatus("Mindestens 1 Phase nötig", "warning"); return; }
    var row = document.getElementById("phase_" + phaseCount);
    if (row) row.remove();
    phaseCount--;
}

/* ═══════════════════════════════════════════════
   PHASEN AUSLESEN – Farbe robust aus Hex-Input
   ═══════════════════════════════════════════════ */
function getPhases() {
    var phases = [];
    for (var i = 1; i <= phaseCount; i++) {
        var row = document.getElementById("phase_" + i);
        if (!row) continue;

        var name  = row.querySelector(".phase-name").value || ("Phase " + i);
        var start = row.querySelector(".phase-start").value;
        var end   = row.querySelector(".phase-end").value;
        if (!start || !end) continue;

        /* Farbe: Zuerst Hex-Input lesen (ist immer synchron zum Picker) */
        var hexEl    = row.querySelector(".phase-hex");
        var pickerEl = row.querySelector(".phase-color");

        var hex;
        if (hexEl && hexEl.value) {
            hex = normalizeHex(hexEl.value);
        }
        if (!hex && pickerEl && pickerEl.value) {
            hex = normalizeHex(pickerEl.value);
        }
        if (!hex) {
            hex = normalizeHex(PHASE_COLORS[(i - 1) % PHASE_COLORS.length]);
        }

        phases.push({
            name:  name,
            start: new Date(start),
            end:   new Date(end),
            color: hex
        });
    }
    return phases;
}

/* ═══════════════════════════════════════════════
   TIME SLOTS
   ═══════════════════════════════════════════════ */
function getTimeSlots(startDate, endDate, unit) {
    var slots = [];
    var d = new Date(startDate.getTime());
    while (d < endDate) {
        var label, slotEnd;
        if (unit === "days") {
            label = d.getDate() + "." + (d.getMonth() + 1) + ".";
            slotEnd = new Date(d); slotEnd.setDate(slotEnd.getDate() + 1);
        } else if (unit === "weeks") {
            label = "KW " + getISOWeek(d);
            slotEnd = new Date(d); slotEnd.setDate(slotEnd.getDate() + 7);
        } else if (unit === "months") {
            var mNames = ["Jan","Feb","Mär","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez"];
            label = mNames[d.getMonth()] + " " + d.getFullYear().toString().substr(2);
            slotEnd = new Date(d.getFullYear(), d.getMonth() + 1, 1);
        } else {
            var q = Math.floor(d.getMonth() / 3) + 1;
            label = "Q" + q + "/" + d.getFullYear().toString().substr(2);
            slotEnd = new Date(d.getFullYear(), d.getMonth() + 3, 1);
        }
        slots.push({ label: label, start: new Date(d), end: slotEnd });
        d = slotEnd;
    }
    return slots;
}

function getISOWeek(date) {
    var d = new Date(date.getTime());
    d.setHours(0, 0, 0, 0);
    d.setDate(d.getDate() + 3 - (d.getDay() + 6) % 7);
    var week1 = new Date(d.getFullYear(), 0, 4);
    return 1 + Math.round(((d - week1) / 86400000 - 3 + (week1.getDay() + 6) % 7) / 7);
}

/* ═══════════════════════════════════════════════
   RE → Points Conversion
   ═══════════════════════════════════════════════ */
function re2pt(re) {
    return Math.round(re * gridUnitCm * 72 / 2.54);
}

/* ═══════════════════════════════════════════════
   GENERATE GANTT
   ═══════════════════════════════════════════════ */
function generateGantt() {
    if (!apiOk) {
        showStatus("PowerPoint API nicht verfügbar", "error");
        return;
    }

    var phases = getPhases();
    if (phases.length === 0) {
        showStatus("Keine gültigen Phasen definiert", "error");
        return;
    }

    var startDate = new Date(document.getElementById("startDate").value);
    var endDate   = new Date(document.getElementById("endDate").value);
    if (isNaN(startDate) || isNaN(endDate) || endDate <= startDate) {
        showStatus("Ungültiger Zeitraum", "error");
        return;
    }

    var unit = document.getElementById("timeUnit").value;
    var timeSlots = getTimeSlots(startDate, endDate, unit);
    if (timeSlots.length === 0 || timeSlots.length > 200) {
        showStatus("Zu viele / zu wenige Zeitabschnitte (" + timeSlots.length + ")", "error");
        return;
    }

    var showHeader    = document.getElementById("showHeader").checked;
    var showLabels    = document.getElementById("showLabels").checked;
    var showGridLines = document.getElementById("showGridLines").checked;
    var showToday     = document.getElementById("showToday").checked;

    var labelWidthRE    = showLabels ? 15 : 0;
    var headerHeightRE  = showHeader ? 3 : 0;
    var chartWidthRE    = GANTT_MAX_W - labelWidthRE;
    var slotWidthRE     = Math.max(1, Math.round(chartWidthRE / timeSlots.length));
    chartWidthRE        = slotWidthRE * timeSlots.length;

    var rowHeightRE = ROW_HEIGHT_RE;
    var barHeightRE = BAR_HEIGHT_RE;

    var totalHeightRE = headerHeightRE + (rowHeightRE * phases.length);
    if (totalHeightRE > GANTT_MAX_H) {
        showStatus("Warnung: GANTT (" + totalHeightRE + " RE) größer als Max (" + GANTT_MAX_H + " RE)!", "warning");
    }

    showStatus("Erstelle GANTT... " + timeSlots.length + " Spalten, " + phases.length + " Phasen", "info");

    PowerPoint.run(function (context) {
        var slide = context.presentation.getSelectedSlides().getItemAt(0);
        slide.load("id");
        return context.sync().then(function () {
            buildGantt(slide, phases, timeSlots, {
                x0: GANTT_LEFT, y0: GANTT_TOP,
                labelWidthRE: labelWidthRE, headerHeightRE: headerHeightRE,
                chartWidthRE: chartWidthRE, slotWidthRE: slotWidthRE,
                rowHeightRE: rowHeightRE, barHeightRE: barHeightRE,
                startDate: startDate, endDate: endDate,
                showHeader: showHeader, showLabels: showLabels,
                showGridLines: showGridLines, showToday: showToday
            });
            return context.sync();
        }).then(function () {
            showStatus("GANTT erfolgreich erstellt!", "success");
        });
    }).catch(function (e) {
        showStatus("Fehler: " + e.message, "error");
    });
}

/* ═══════════════════════════════════════════════
   BUILD GANTT – Shapes auf Folie erzeugen
   ═══════════════════════════════════════════════ */
function buildGantt(slide, phases, timeSlots, cfg) {
    var x0 = cfg.x0;
    var y0 = cfg.y0;
    var chartX0 = x0 + cfg.labelWidthRE;
    var chartY0 = y0 + cfg.headerHeightRE;

    /*
      1) HINTERGRUND (WEISS)
    */
    var totalW = cfg.labelWidthRE + cfg.chartWidthRE;
    var totalH = cfg.headerHeightRE + (cfg.rowHeightRE * phases.length);
    var bg = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    bg.left   = re2pt(x0);
    bg.top    = re2pt(y0);
    bg.width  = re2pt(totalW);
    bg.height = re2pt(totalH);
    bg.fill.setSolidColor("FFFFFF");
    bg.lineFormat.visible = false;
    bg.name = "GANTT_BG";

    /*
      2) HEADER (Zeitslot-Beschriftungen)
    */
    if (cfg.showHeader) {
        for (var h = 0; h < timeSlots.length; h++) {
            var hx = chartX0 + (h * cfg.slotWidthRE);
            var htb = slide.shapes.addTextBox(timeSlots[h].label, {
                left:   re2pt(hx),
                top:    re2pt(y0),
                width:  re2pt(cfg.slotWidthRE),
                height: re2pt(cfg.headerHeightRE)
            });
            htb.name = "GANTT_HDR_" + h;
            htb.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
            htb.textFrame.textRange.font.size = 6;
            htb.textFrame.textRange.font.color = "1A1A2E";
            htb.textFrame.textRange.paragraphFormat.horizontalAlignment =
                PowerPoint.ParagraphHorizontalAlignment.center;
            htb.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;
            htb.fill.clear();
            htb.lineFormat.visible = false;
        }
    }

    /*
      3) LABELS links (Phasennamen)
    */
    if (cfg.showLabels) {
        for (var l = 0; l < phases.length; l++) {
            var ly = chartY0 + (l * cfg.rowHeightRE);
            var ltb = slide.shapes.addTextBox(phases[l].name, {
                left:   re2pt(x0),
                top:    re2pt(ly),
                width:  re2pt(cfg.labelWidthRE),
                height: re2pt(cfg.rowHeightRE)
            });
            ltb.name = "GANTT_LBL_" + l;
            ltb.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
            ltb.textFrame.textRange.font.size = 8;
            ltb.textFrame.textRange.font.color = "1A1A2E";
            ltb.textFrame.textRange.font.bold = true;
            ltb.textFrame.textRange.paragraphFormat.horizontalAlignment =
                PowerPoint.ParagraphHorizontalAlignment.left;
            ltb.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;
            ltb.fill.clear();
            ltb.lineFormat.visible = false;
        }
    }

    /*
      4) (reserviert)
    */

    /*
      5) VERTIKALE RASTERLINIEN
    */
    if (cfg.showGridLines) {
        var gridH = cfg.headerHeightRE + (cfg.rowHeightRE * phases.length);
        for (var g = 0; g <= timeSlots.length; g++) {
            var gx = chartX0 + (g * cfg.slotWidthRE);
            var gl = slide.shapes.addLine(
                PowerPoint.ConnectorType.straight,
                { left: re2pt(gx), top: re2pt(y0), width: 0, height: re2pt(gridH) }
            );
            gl.name = "GANTT_GRID_" + g;
            gl.lineFormat.color = "D5D8DC";
            gl.lineFormat.weight = 0.5;
        }
    }

    /*
      6) GANTT-BALKEN (Phasen) – JEDE PHASE MIT EIGENER FARBE
    */
    var totalMs = cfg.endDate.getTime() - cfg.startDate.getTime();
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

        /* Balken-Shape erzeugen */
        var bar = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundedRectangle);
        bar.left   = re2pt(chartX0 + barLeftRE);
        bar.top    = re2pt(barY);
        bar.width  = re2pt(barWidthRE);
        bar.height = re2pt(cfg.barHeightRE);

        /* ── FARBE SETZEN: phase.color ist bereits 6-stellig UPPERCASE ── */
        try {
            bar.fill.setSolidColor(phase.color);
        } catch (colorErr) {
            bar.fill.setSolidColor("2471A3");
        }

        bar.lineFormat.visible = false;
        bar.name = "GANTT_BAR_" + p;

        /* Balken-Text (Phasenname auf dem Balken) */
        if (barWidthRE >= 5) {
            var barTb = slide.shapes.addTextBox(phase.name, {
                left:   re2pt(chartX0 + barLeftRE),
                top:    re2pt(barY),
                width:  re2pt(barWidthRE),
                height: re2pt(cfg.barHeightRE)
            });
            barTb.name = "GANTT_BARTXT_" + p;
            barTb.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
            barTb.textFrame.textRange.font.size = 7;
            barTb.textFrame.textRange.font.color = "FFFFFF";
            barTb.textFrame.textRange.font.bold = true;
            barTb.textFrame.textRange.paragraphFormat.horizontalAlignment =
                PowerPoint.ParagraphHorizontalAlignment.center;
            barTb.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;
            barTb.fill.clear();
            barTb.lineFormat.visible = false;
        }
    }

    /*
      7) HEUTE-LINIE
    */
    if (cfg.showToday) {
        var now = new Date();
        if (now >= cfg.startDate && now <= cfg.endDate) {
            var todayRatio = (now.getTime() - cfg.startDate.getTime()) / totalMs;
            var todayRE = Math.round(todayRatio * cfg.chartWidthRE);
            var todayH = cfg.headerHeightRE + (cfg.rowHeightRE * phases.length);
            var tl = slide.shapes.addLine(
                PowerPoint.ConnectorType.straight,
                { left: re2pt(chartX0 + todayRE), top: re2pt(y0), width: 0, height: re2pt(todayH) }
            );
            tl.name = "GANTT_TODAY";
            tl.lineFormat.color = "E74C3C";
            tl.lineFormat.weight = 1.5;
            tl.lineFormat.dashStyle = PowerPoint.ShapeLineDashStyle.dash;
        }
    }
}
