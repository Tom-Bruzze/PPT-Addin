/*
  DROEGE GANTT Generator – taskpane.js
  Build 14.02.2026 v3
  FIX: Farb-Palette statt input[type=color]
  FIX: Rasterlinien als dünne Rechtecke (nicht addLine)
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

/* 16 gut unterscheidbare Farben für die Palette */
var PALETTE = [
    "#2471A3", "#27AE60", "#8E44AD", "#E67E22",
    "#2980B9", "#1ABC9C", "#C0392B", "#D4AC0D",
    "#16A085", "#E74C3C", "#3498DB", "#F39C12",
    "#1F618D", "#239B56", "#7D3C98", "#AF601A"
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
    var preButtons = document.querySelectorAll(".pre");
    for (var i = 0; i < preButtons.length; i++) {
        preButtons[i].addEventListener("click", function () {
            for (var j = 0; j < preButtons.length; j++) preButtons[j].classList.remove("active");
            this.classList.add("active");
            gridUnitCm = parseFloat(this.getAttribute("data-v"));
            showStatus("RE = " + gridUnitCm + " cm", "info");
        });
    }

    var ids = [
        ["ganttLeft",    function(v){ if(v>=0)  GANTT_LEFT=v; }],
        ["ganttTop",     function(v){ if(v>=0)  GANTT_TOP=v; }],
        ["ganttMaxW",    function(v){ if(v>=10) GANTT_MAX_W=v; }],
        ["ganttMaxH",    function(v){ if(v>=10) GANTT_MAX_H=v; }]
    ];
    for (var k = 0; k < ids.length; k++) {
        (function(id, fn) {
            var el = document.getElementById(id);
            if (el) el.addEventListener("change", function(){ var v=parseInt(this.value); if(!isNaN(v)) fn(v); });
        })(ids[k][0], ids[k][1]);
    }

    document.getElementById("rowHeightRE").addEventListener("change", function () {
        var v = parseInt(this.value);
        if (!isNaN(v) && v >= 2) {
            ROW_HEIGHT_RE = v;
            if (BAR_HEIGHT_RE >= v) {
                BAR_HEIGHT_RE = v - 1;
                document.getElementById("barHeightRE").value = BAR_HEIGHT_RE;
            }
        }
    });
    document.getElementById("barHeightRE").addEventListener("change", function () {
        var v = parseInt(this.value);
        if (!isNaN(v) && v >= 1) {
            if (v >= ROW_HEIGHT_RE) {
                showStatus("Balkenhöhe < Zeilenhöhe!", "warning");
                this.value = BAR_HEIGHT_RE;
            } else { BAR_HEIGHT_RE = v; }
        }
    });

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
    var s = new Date(now.getFullYear(), now.getMonth(), 1);
    var e = new Date(now.getFullYear(), now.getMonth() + 6, 0);
    document.getElementById("startDate").value = fmtDate(s);
    document.getElementById("endDate").value   = fmtDate(e);
}
function fmtDate(d) {
    var m = (d.getMonth()+1).toString(); if(m.length<2) m="0"+m;
    var day = d.getDate().toString();    if(day.length<2) day="0"+day;
    return d.getFullYear()+"-"+m+"-"+day;
}

/* ═══════════════════════════════════════════════════════
   NORMALIZE HEX – robust
   ═══════════════════════════════════════════════════════ */
function normalizeHex(raw) {
    if (!raw) return null;
    var h = raw.replace(/^#/, "").toUpperCase();
    if (h.length === 3) h = h[0]+h[0]+h[1]+h[1]+h[2]+h[2];
    if (!/^[0-9A-F]{6}$/.test(h)) return null;
    return h;
}

/* ═══════════════════════════════════════════════════════
   PHASE ROWS – Klickbare Farbpalette + Hex-Eingabe
   Kein <input type="color"> mehr!
   ═══════════════════════════════════════════════════════ */
function addPhaseRow() {
    phaseCount++;
    var idx = phaseCount;
    var container = document.getElementById("phaseContainer");
    var defaultColor = PALETTE[(idx - 1) % PALETTE.length];

    var startD = document.getElementById("startDate").value;
    var endD   = document.getElementById("endDate").value;

    /* ── Zeile ── */
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

    /* Zeile 1: Name */
    var row1 = document.createElement("div");
    row1.className = "phase-field-row";
    var nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.className = "phase-name";
    nameInput.placeholder = "Phase " + idx;
    nameInput.value = "Phase " + idx;
    row1.appendChild(nameInput);
    fields.appendChild(row1);

    /* Zeile 2: Start + End */
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

    /* Zeile 3: Farb-Auswahl  =  Swatch + Hex + Palette-Toggle */
    var row3 = document.createElement("div");
    row3.className = "phase-color-row";

    /* Aktuelles Farbquadrat (nur Anzeige) */
    var swatch = document.createElement("div");
    swatch.className = "phase-swatch-current ring";
    swatch.style.backgroundColor = defaultColor;
    swatch.title = "Aktuelle Farbe";
    row3.appendChild(swatch);

    /* Hex-Eingabe */
    var hexInput = document.createElement("input");
    hexInput.type = "text";
    hexInput.className = "phase-hex";
    hexInput.value = defaultColor.replace("#", "");
    hexInput.maxLength = 7;
    hexInput.placeholder = "HEX";
    hexInput.setAttribute("data-color", defaultColor.replace("#", ""));
    row3.appendChild(hexInput);

    /* Palette-Toggle-Button */
    var palBtn = document.createElement("button");
    palBtn.type = "button";
    palBtn.className = "phase-palette-btn";
    palBtn.textContent = "▼ Palette";
    row3.appendChild(palBtn);

    fields.appendChild(row3);

    /* Zeile 4: Farbpalette (versteckt, aufklappbar) */
    var palDiv = document.createElement("div");
    palDiv.className = "phase-palette";

    for (var c = 0; c < PALETTE.length; c++) {
        var sw = document.createElement("div");
        sw.className = "pal-swatch";
        sw.style.backgroundColor = PALETTE[c];
        sw.setAttribute("data-hex", PALETTE[c].replace("#", ""));
        if (PALETTE[c].toUpperCase() === defaultColor.toUpperCase()) {
            sw.classList.add("selected");
        }
        palDiv.appendChild(sw);
    }
    fields.appendChild(palDiv);

    div.appendChild(fields);
    container.appendChild(div);

    /* ── EVENTS ── */

    /* Palette öffnen/schließen */
    palBtn.addEventListener("click", function () {
        var open = palDiv.classList.toggle("open");
        palBtn.textContent = open ? "▲ Palette" : "▼ Palette";
    });

    /* Swatch auch öffnet Palette */
    swatch.addEventListener("click", function () {
        var open = palDiv.classList.toggle("open");
        palBtn.textContent = open ? "▲ Palette" : "▼ Palette";
    });

    /* Palette-Swatch klicken → Farbe setzen */
    palDiv.addEventListener("click", function (evt) {
        var target = evt.target;
        if (!target.classList.contains("pal-swatch")) return;
        var hex = target.getAttribute("data-hex");
        if (!hex) return;

        /* Alle swatches de-selecten */
        var all = palDiv.querySelectorAll(".pal-swatch");
        for (var s = 0; s < all.length; s++) all[s].classList.remove("selected");
        target.classList.add("selected");

        /* Swatch + Hex aktualisieren */
        swatch.style.backgroundColor = "#" + hex;
        hexInput.value = hex;
        hexInput.setAttribute("data-color", hex);

        /* Palette schließen */
        palDiv.classList.remove("open");
        palBtn.textContent = "▼ Palette";
    });

    /* Hex-Input → Swatch aktualisieren */
    hexInput.addEventListener("input", function () {
        var norm = normalizeHex(this.value);
        if (norm) {
            swatch.style.backgroundColor = "#" + norm;
            hexInput.setAttribute("data-color", norm);
            /* Palette-Swatch markieren falls passend */
            var all = palDiv.querySelectorAll(".pal-swatch");
            for (var s = 0; s < all.length; s++) {
                all[s].classList.toggle("selected", all[s].getAttribute("data-hex") === norm);
            }
        }
    });
    hexInput.addEventListener("change", function () {
        var norm = normalizeHex(this.value);
        if (norm) {
            this.value = norm;
            swatch.style.backgroundColor = "#" + norm;
            hexInput.setAttribute("data-color", norm);
        } else {
            /* Ungültig → Zurücksetzen */
            var prev = hexInput.getAttribute("data-color") || PALETTE[0].replace("#","");
            this.value = prev;
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

/* ═══════════════════════════════════════════════════════
   PHASEN AUSLESEN
   ═══════════════════════════════════════════════════════ */
function getPhases() {
    var phases = [];
    for (var i = 1; i <= phaseCount; i++) {
        var row = document.getElementById("phase_" + i);
        if (!row) continue;

        var name  = row.querySelector(".phase-name").value || ("Phase " + i);
        var start = row.querySelector(".phase-start").value;
        var end   = row.querySelector(".phase-end").value;
        if (!start || !end) continue;

        /* Farbe aus data-Attribut des Hex-Inputs lesen */
        var hexEl = row.querySelector(".phase-hex");
        var hex = null;
        if (hexEl) {
            /* Primär: data-color (immer sauber normalisiert) */
            hex = hexEl.getAttribute("data-color");
            /* Fallback: Input-Wert direkt */
            if (!hex) hex = normalizeHex(hexEl.value);
        }
        /* Letzter Fallback: Default-Palette */
        if (!hex) hex = PALETTE[(i-1) % PALETTE.length].replace("#","");

        phases.push({
            name:  name,
            start: new Date(start),
            end:   new Date(end),
            color: hex.toUpperCase()
        });
    }
    return phases;
}

/* ═══════════════════════════════════════════════════════
   TIME SLOTS
   ═══════════════════════════════════════════════════════ */
function getTimeSlots(startDate, endDate, unit) {
    var slots = [];
    var d = new Date(startDate.getTime());
    while (d < endDate) {
        var label, slotEnd;
        if (unit === "days") {
            label = d.getDate() + "." + (d.getMonth()+1) + ".";
            slotEnd = new Date(d); slotEnd.setDate(slotEnd.getDate()+1);
        } else if (unit === "weeks") {
            label = "KW " + getISOWeek(d);
            slotEnd = new Date(d); slotEnd.setDate(slotEnd.getDate()+7);
        } else if (unit === "months") {
            var mNames = ["Jan","Feb","Mär","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez"];
            label = mNames[d.getMonth()] + " " + d.getFullYear().toString().substr(2);
            slotEnd = new Date(d.getFullYear(), d.getMonth()+1, 1);
        } else {
            var q = Math.floor(d.getMonth()/3)+1;
            label = "Q" + q + "/" + d.getFullYear().toString().substr(2);
            slotEnd = new Date(d.getFullYear(), d.getMonth()+3, 1);
        }
        slots.push({ label: label, start: new Date(d), end: slotEnd });
        d = slotEnd;
    }
    return slots;
}

function getISOWeek(date) {
    var d = new Date(date.getTime());
    d.setHours(0,0,0,0);
    d.setDate(d.getDate()+3-(d.getDay()+6)%7);
    var w1 = new Date(d.getFullYear(),0,4);
    return 1+Math.round(((d-w1)/86400000-3+(w1.getDay()+6)%7)/7);
}

/* ═══════════════════════════════════════════════════════
   RE → Points
   ═══════════════════════════════════════════════════════ */
function re2pt(re) {
    return Math.round(re * gridUnitCm * 72 / 2.54);
}

/* ═══════════════════════════════════════════════════════
   GENERATE
   ═══════════════════════════════════════════════════════ */
function generateGantt() {
    if (!apiOk) { showStatus("PowerPoint API nicht verfügbar","error"); return; }

    var phases = getPhases();
    if (phases.length === 0) { showStatus("Keine gültigen Phasen","error"); return; }

    var startDate = new Date(document.getElementById("startDate").value);
    var endDate   = new Date(document.getElementById("endDate").value);
    if (isNaN(startDate) || isNaN(endDate) || endDate <= startDate) {
        showStatus("Ungültiger Zeitraum","error"); return;
    }

    var unit = document.getElementById("timeUnit").value;
    var timeSlots = getTimeSlots(startDate, endDate, unit);
    if (timeSlots.length === 0 || timeSlots.length > 200) {
        showStatus("Zeitabschnitte: " + timeSlots.length + " (0–200 erlaubt)","error"); return;
    }

    var showHeader    = document.getElementById("showHeader").checked;
    var showLabels    = document.getElementById("showLabels").checked;
    var showGridLines = document.getElementById("showGridLines").checked;
    var showToday     = document.getElementById("showToday").checked;

    var labelWidthRE   = showLabels ? 15 : 0;
    var headerHeightRE = showHeader ? 3  : 0;
    var chartWidthRE   = GANTT_MAX_W - labelWidthRE;
    var slotWidthRE    = Math.max(1, Math.round(chartWidthRE / timeSlots.length));
    chartWidthRE       = slotWidthRE * timeSlots.length;

    showStatus("Erstelle... " + timeSlots.length + " Spalten, " + phases.length + " Phasen","info");

    PowerPoint.run(function (context) {
        var slide = context.presentation.getSelectedSlides().getItemAt(0);
        slide.load("id");
        return context.sync().then(function () {
            buildGantt(slide, phases, timeSlots, {
                x0: GANTT_LEFT, y0: GANTT_TOP,
                labelWidthRE: labelWidthRE,
                headerHeightRE: headerHeightRE,
                chartWidthRE: chartWidthRE,
                slotWidthRE: slotWidthRE,
                rowHeightRE: ROW_HEIGHT_RE,
                barHeightRE: BAR_HEIGHT_RE,
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

/* ═══════════════════════════════════════════════════════
   BUILD GANTT
   ═══════════════════════════════════════════════════════ */
function buildGantt(slide, phases, timeSlots, cfg) {
    var x0      = cfg.x0;
    var y0      = cfg.y0;
    var chartX0 = x0 + cfg.labelWidthRE;
    var chartY0 = y0 + cfg.headerHeightRE;
    var totalW  = cfg.labelWidthRE + cfg.chartWidthRE;
    var totalH  = cfg.headerHeightRE + (cfg.rowHeightRE * phases.length);

    /* ── 1) HINTERGRUND (WEISS) ── */
    var bg = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    bg.left   = re2pt(x0);
    bg.top    = re2pt(y0);
    bg.width  = re2pt(totalW);
    bg.height = re2pt(totalH);
    bg.fill.setSolidColor("FFFFFF");
    bg.lineFormat.visible = false;
    bg.name = "GANTT_BG";

    /* ── 2) HEADER ── */
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

    /* ── 3) LABELS ── */
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

    /* ══════════════════════════════════════════════════════════
       4) VERTIKALE RASTERLINIEN
       FIX: Dünne Rechtecke statt addLine() → garantiert gerade!
       ══════════════════════════════════════════════════════════ */
    if (cfg.showGridLines) {
        var gridH = totalH;
        /* Linienbreite: 0.5pt (≈ 0.018 cm) */
        var lineWidthPt = 0.5;

        for (var g = 0; g <= timeSlots.length; g++) {
            var gx = chartX0 + (g * cfg.slotWidthRE);
            var gl = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
            gl.left   = re2pt(gx);
            gl.top    = re2pt(y0);
            gl.width  = lineWidthPt;
            gl.height = re2pt(gridH);
            gl.fill.setSolidColor("D5D8DC");
            gl.lineFormat.visible = false;
            gl.name = "GANTT_GRID_" + g;
        }
    }

    /* ══════════════════════════════════════════════════════════
       5) BALKEN – jede Phase mit ihrer individuellen Farbe
       ══════════════════════════════════════════════════════════ */
    var totalMs = cfg.endDate.getTime() - cfg.startDate.getTime();
    for (var p = 0; p < phases.length; p++) {
        var phase = phases[p];
        var pStart = Math.max(phase.start.getTime(), cfg.startDate.getTime());
        var pEnd   = Math.min(phase.end.getTime(),   cfg.endDate.getTime());
        if (pEnd <= pStart) continue;

        var startRatio = (pStart - cfg.startDate.getTime()) / totalMs;
        var endRatio   = (pEnd   - cfg.startDate.getTime()) / totalMs;
        var barLeftRE  = Math.round(startRatio * cfg.chartWidthRE);
        var barRightRE = Math.round(endRatio   * cfg.chartWidthRE);
        var barWidthRE = Math.max(1, barRightRE - barLeftRE);

        var rowY       = chartY0 + (p * cfg.rowHeightRE);
        var barYOffset = Math.floor((cfg.rowHeightRE - cfg.barHeightRE) / 2);
        var barY       = rowY + barYOffset;

        var bar = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundedRectangle);
        bar.left   = re2pt(chartX0 + barLeftRE);
        bar.top    = re2pt(barY);
        bar.width  = re2pt(barWidthRE);
        bar.height = re2pt(cfg.barHeightRE);

        /* ── FARBE: phase.color ist 6-stellig UPPERCASE (z.B. "2471A3") ── */
        try {
            bar.fill.setSolidColor(phase.color);
        } catch (err) {
            bar.fill.setSolidColor("2471A3");
        }
        bar.lineFormat.visible = false;
        bar.name = "GANTT_BAR_" + p;

        /* Balken-Text */
        if (barWidthRE >= 5) {
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

    /* ══════════════════════════════════════════════════════════
       6) HEUTE-LINIE (ebenfalls als dünnes Rechteck!)
       ══════════════════════════════════════════════════════════ */
    if (cfg.showToday) {
        var now = new Date();
        if (now >= cfg.startDate && now <= cfg.endDate) {
            var todayRatio = (now.getTime() - cfg.startDate.getTime()) / totalMs;
            var todayRE    = Math.round(todayRatio * cfg.chartWidthRE);
            var tl = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
            tl.left   = re2pt(chartX0 + todayRE);
            tl.top    = re2pt(y0);
            tl.width  = 1.5;
            tl.height = re2pt(totalH);
            tl.fill.setSolidColor("E74C3C");
            tl.lineFormat.visible = false;
            tl.name = "GANTT_TODAY";
        }
    }
}
