/*
  DROEGE GANTT Generator – taskpane.js
  Build 14.02.2026 v10
  FIX: Farbe wird DIREKT beim Shape-Erstellen gesetzt (kein 2-Pass mehr)
       - cleanHex() liefert 6 Zeichen OHNE # → exakt was setSolidColor() braucht
       - Farbauswahl komplett neu: simples data-hex Attribut als Single Source of Truth
       - Kein dataset.color, kein style.backgroundColor Parsing
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

/* 16 gut unterscheidbare Farben */
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
   CLEAN HEX – gibt IMMER 6 Zeichen UPPERCASE zurück, OHNE #
   Das ist EXAKT das Format das setSolidColor() erwartet
   ═══════════════════════════════════════════════════════ */
function cleanHex(raw) {
    if (!raw) return null;
    var h = String(raw).replace(/[^0-9A-Fa-f]/g, "").toUpperCase();
    if (h.length === 3) h = h[0]+h[0]+h[1]+h[1]+h[2]+h[2];
    if (h.length !== 6) return null;
    return "#" + h;
}

/* ═══════════════════════════════════════════════════════
   PHASE ROWS – Komplett neugeschrieben
   
   Single Source of Truth: Das versteckte input.phase-color-value
   speichert den Hex-Wert OHNE # (z.B. "2471A3").
   
   Kein Parsen von style.backgroundColor, kein data-color,
   keine Mehrdeutigkeit.
   ═══════════════════════════════════════════════════════ */
function addPhaseRow() {
    phaseCount++;
    var idx = phaseCount;
    var container = document.getElementById("phaseContainer");
    var defaultHex = PALETTE[(idx - 1) % PALETTE.length];
    console.log('[addPhaseRow] Phase ' + idx + ' defaultHex:', defaultHex); /* schon OHNE # */

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

    /* Zeile 3: Farb-Auswahl */
    var row3 = document.createElement("div");
    row3.className = "phase-color-row";

    /* ★ VERSTECKTES INPUT = Single Source of Truth für die Farbe ★ */
    var colorVal = document.createElement("input");
    colorVal.type = "hidden";
    colorVal.className = "phase-color-value";
    colorVal.value = defaultHex;
    console.log('[addPhaseRow] Phase ' + idx + ' colorVal.value set to:', colorVal.value);
    row3.appendChild(colorVal);

    /* Sichtbares Farbquadrat (zeigt aktuelle Farbe) */
    var swatch = document.createElement("div");
    swatch.className = "phase-swatch-current";
    swatch.style.backgroundColor = defaultHex;
    swatch.title = "Klicken für Palette";
    row3.appendChild(swatch);

    /* Hex-Eingabe */
    var hexInput = document.createElement("input");
    hexInput.type = "text";
    hexInput.className = "phase-hex";
    hexInput.value = defaultHex;
    hexInput.maxLength = 7;
    hexInput.placeholder = "HEX";
    row3.appendChild(hexInput);

    /* Palette-Toggle-Button */
    var palBtn = document.createElement("button");
    palBtn.type = "button";
    palBtn.className = "phase-palette-btn";
    palBtn.textContent = "\u25BC";
    row3.appendChild(palBtn);

    fields.appendChild(row3);

    /* Zeile 4: Farbpalette (versteckt, wird bei Klick geöffnet) */
    var palDiv = document.createElement("div");
    palDiv.className = "phase-palette";

    for (var c = 0; c < PALETTE.length; c++) {
        var palHex = PALETTE[c];
        var sw = document.createElement("div");
        sw.className = "pal-swatch";
        sw.style.backgroundColor = palHex;
        sw.setAttribute("data-hex", palHex);
        if (palHex === defaultHex) sw.classList.add("selected");
        palDiv.appendChild(sw);
    }
    fields.appendChild(palDiv);

    div.appendChild(fields);
    container.appendChild(div);

    /* ═════════════════════════════════════════════════
       EVENTS – Alle setzen colorVal.value als Quelle
       ═════════════════════════════════════════════════ */

    /* Palette öffnen/schließen */
    palBtn.addEventListener("click", function () {
        var isOpen = palDiv.classList.toggle("open");
        palBtn.textContent = isOpen ? "\u25B2" : "\u25BC";
    });

    swatch.addEventListener("click", function () {
        var isOpen = palDiv.classList.toggle("open");
        palBtn.textContent = isOpen ? "\u25B2" : "\u25BC";
    });

    /* Palette-Swatch klicken → Farbe in ALLE drei Stellen schreiben */
    palDiv.addEventListener("click", function (evt) {
        console.log('[PaletteClick] Event fired');
        var target = evt.target;
        if (!target.classList.contains("pal-swatch")) return;
        var hex = target.getAttribute("data-hex");
        if (!hex) return;

        /* 1. Hidden Input (Source of Truth) */
        colorVal.value = hex;
        /* 2. Sichtbares Swatch */
        swatch.style.backgroundColor = "#" + hex;
        /* 3. Hex-Eingabefeld */
        hexInput.value = hex;

        /* Selection-Marker */
        var all = palDiv.querySelectorAll(".pal-swatch");
        for (var s = 0; s < all.length; s++) all[s].classList.remove("selected");
        target.classList.add("selected");

        /* Palette schließen */
        palDiv.classList.remove("open");
        palBtn.textContent = "\u25BC";
    });

    /* Hex-Input → live aktualisieren */
    hexInput.addEventListener("input", function () {
        var norm = cleanHex(this.value);
        if (norm) {
            colorVal.value = norm;
            swatch.style.backgroundColor = "#" + norm;
            var all = palDiv.querySelectorAll(".pal-swatch");
            for (var s = 0; s < all.length; s++) {
                all[s].classList.toggle("selected", all[s].getAttribute("data-hex") === norm);
            }
        }
    });

    hexInput.addEventListener("change", function () {
        var norm = cleanHex(this.value);
        if (norm) {
            this.value = norm;  // Display without # 
            colorVal.value = norm;
            swatch.style.backgroundColor = "#" + norm;
        } else {
            this.value = defaultHex;
            colorVal.value = defaultHex;
    console.log('[addPhaseRow] Phase ' + idx + ' colorVal.value set to:', colorVal.value);
            swatch.style.backgroundColor = defaultHex;
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
   Liest die Farbe aus dem Hidden Input (Single Source of Truth)
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

        /* ★ Farbe aus dem Hidden Input lesen ★ */
        var colorEl = row.querySelector(".phase-color-value");
        console.log('[getPhases] Phase ' + i + ' colorEl:', colorEl);
        var hex = colorEl ? colorEl.value : null;
        console.log('[getPhases] Phase ' + i + ' hex from colorEl.value:', hex);

        /* Validierung + Fallback */
        hex = cleanHex(hex);
        console.log('[getPhases] Phase ' + i + ' after cleanHex:', hex);
        if (!hex) hex = PALETTE[(i-1) % PALETTE.length];
        console.log('[getPhases] Phase ' + i + ' final hex (after fallback if needed):', hex);

        phases.push({
            name:  name,
            start: new Date(start),
            end:   new Date(end),
            color: hex   /* z.B. "2471A3" – 6 Zeichen UPPERCASE OHNE # */
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
   GENERATE – KEIN Zwei-Pass mehr!
   Farbe wird DIREKT beim Erstellen des Shapes gesetzt.
   ═══════════════════════════════════════════════════════ */
function generateGantt() {
    if (!apiOk) { showStatus("PowerPoint API nicht verfügbar","error"); return; }

    var phases = getPhases();
    if (phases.length === 0) { showStatus("Keine gültigen Phasen","error"); return; }

    /* Debug: Farben in Status anzeigen */
    var colorDebug = phases.map(function(p){ return p.name + "=" + p.color; }).join(", ");
    showStatus("Erstelle... " + colorDebug, "info");

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
            showStatus("GANTT erfolgreich erstellt! (" + phases.length + " Phasen)", "success");
        });
    }).catch(function (e) {
        showStatus("Fehler: " + e.message, "error");
    });
}

/* ═══════════════════════════════════════════════════════
   BUILD GANTT – Ein einziger Pass
   Farbe wird SOFORT bei Erstellung auf den Shape gesetzt
   ═══════════════════════════════════════════════════════ */
function buildGantt(slide, phases, timeSlots, cfg) {
    var x0      = cfg.x0;
    var y0      = cfg.y0;
    var chartX0 = x0 + cfg.labelWidthRE;
    var chartY0 = y0 + cfg.headerHeightRE;
    var totalW  = cfg.labelWidthRE + cfg.chartWidthRE;
    var totalH  = cfg.headerHeightRE + (cfg.rowHeightRE * phases.length);

    /* ── 1) HINTERGRUND ── */
    var bg = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    bg.left   = re2pt(x0);
    bg.top    = re2pt(y0);
    bg.width  = re2pt(totalW);
    bg.height = re2pt(totalH);
    bg.fill.setSolidColor("#FFFFFF");
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
            htb.textFrame.textRange.font.color = "#000000";
            htb.textFrame.textRange.paragraphFormat.horizontalAlignment =
                PowerPoint.ParagraphHorizontalAlignment.center;
            htb.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;
            htb.fill.setSolidColor("#E6E6E6");
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
            ltb.textFrame.textRange.font.color = "#1A1A2E";
            ltb.textFrame.textRange.font.bold  = true;
            ltb.textFrame.textRange.paragraphFormat.horizontalAlignment =
                PowerPoint.ParagraphHorizontalAlignment.left;
            ltb.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;
            ltb.fill.clear();
            ltb.lineFormat.visible = false;
        }
    }

    /* ── 4) RASTERLINIEN ── */
    if (cfg.showGridLines) {
        for (var g = 0; g <= timeSlots.length; g++) {
            var gx = chartX0 + (g * cfg.slotWidthRE);
            var gl = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
            gl.left   = re2pt(gx);
            gl.top    = re2pt(y0);
            gl.width  = 0.5;
            gl.height = re2pt(totalH);
            gl.fill.setSolidColor("#D5D8DC");
            gl.lineFormat.visible = false;
            gl.name = "GANTT_GRID_" + g;
        }
    }

    /* ── 5) BALKEN – Farbe wird SOFORT gesetzt! ── */
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

        /* ★ Shape erstellen UND sofort einfärben ★ */
        var bar = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundedRectangle);
        bar.left   = re2pt(chartX0 + barLeftRE);
        bar.top    = re2pt(barY);
        bar.width  = re2pt(barWidthRE);
        bar.height = re2pt(cfg.barHeightRE);
        bar.name   = "GANTT_BAR_" + p;
        bar.lineFormat.visible = false;

        console.log('[buildGantt] Bar ' + p + ' setting color:', phase.color);
        /* ★★★ FARBE DIREKT SETZEN – phase.color ist bereits "2471A3" ohne # ★★★ */
        bar.fill.setSolidColor(phase.color);

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
            barTb.textFrame.textRange.font.color = "#FFFFFF";
            barTb.textFrame.textRange.font.bold  = true;
            barTb.textFrame.textRange.paragraphFormat.horizontalAlignment =
                PowerPoint.ParagraphHorizontalAlignment.center;
            barTb.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;
            barTb.fill.clear();
            barTb.lineFormat.visible = false;
        }
    }

    /* ★★★ KRITISCH: Alle Balken-Shapes mit Farben zu PowerPoint übertragen! ★★★ */
    await context.sync();
    console.log('[buildGantt] context.sync() nach Balken-Erstellung abgeschlossen - Farben sollten jetzt sichtbar sein');

    /* ── 6) HEUTE-LINIE ── */
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
            tl.fill.setSolidColor("#E74C3C");
            tl.lineFormat.visible = false;
            tl.name = "GANTT_TODAY";
        }
    }
}

/* ── SECTION TOGGLE ── */
function toggleSection(header) {
    var body = header.nextElementSibling;
    if (body) {
        body.classList.toggle("collapsed");
        header.classList.toggle("is-collapsed");
    }
}
