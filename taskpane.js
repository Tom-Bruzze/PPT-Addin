// GANTT Generator v2.21 - Synchronisierte Version
// ============================================================

Office.onReady(function(info) {
    console.log("Office.onReady - Host:", info.host);
    if (info.host === Office.HostType.PowerPoint) {
        document.getElementById("createGantt").onclick = generateGantt;
        debugLog("GANTT Generator v2.21 bereit");
        debugLog("Klicken Sie 'GANTT erstellen' zum Starten");
    }
});

// ============================================================
// KONSTANTEN
// ============================================================
const POINTS_PER_CM = 28.3464567;
const RE_CM = 0.21;
const RE_PT = RE_CM * POINTS_PER_CM; // 5.95275590551

// Feste Foliengröße für 16:9 (PowerPoint Standard)
const SLIDE_WIDTH_PT = 960;
const SLIDE_HEIGHT_PT = 540;

// GANTT Start-Position in RE (vom Grid-Ursprung)
const GANTT_LEFT_RE = 9;
const GANTT_TOP_RE = 17;

// GANTT Dimensionen
const GANTT_MAX_WIDTH_RE = 118;
const TASK_HEIGHT_RE = 3;
const TASK_GAP_RE = 1;
const HEADER_HEIGHT_RE = 3;

// Liniendicke
const LINE_WEIGHT = 0.5;

// Farben
const COLORS = {
    taskFill: "#1A3A5C",
    taskBorder: "#1A3A5C",
    headerFill: "#FFFFFF",
    headerBorder: "#808080",
    headerText: "#000000"
};

// Grid-Berechnungen
const FULL_UNITS_X = Math.floor(SLIDE_WIDTH_PT / RE_PT);
const MARGIN_LEFT_PT = (SLIDE_WIDTH_PT - (FULL_UNITS_X * RE_PT)) / 2;
const FULL_UNITS_Y = Math.floor(SLIDE_HEIGHT_PT / RE_PT);
const MARGIN_TOP_PT = (SLIDE_HEIGHT_PT - (FULL_UNITS_Y * RE_PT)) / 2;

// Absolute GANTT-Position in Points
const GANTT_LEFT_PT = MARGIN_LEFT_PT + (GANTT_LEFT_RE * RE_PT);
const GANTT_TOP_PT = MARGIN_TOP_PT + (GANTT_TOP_RE * RE_PT);

// ============================================================
// DEBUG FUNKTION
// ============================================================
function debugLog(msg) {
    var out = document.getElementById("debugOutput");
    if (out) {
        out.innerText += msg + "\n";
        out.scrollTop = out.scrollHeight;
    }
    console.log("[GANTT] " + msg);
}

// ============================================================
// HAUPTFUNKTION
// ============================================================
function generateGantt() {
    // Debug-Ausgabe zurücksetzen
    var debugOut = document.getElementById("debugOutput");
    debugOut.innerText = "";
    
    debugLog("=== GANTT v2.21 Start ===");
    
    // Eingaben lesen
    var phasenInput = document.getElementById("phasen");
    var zeitInput = document.getElementById("zeiteinheiten");
    var labelInput = document.getElementById("labelBreite");
    
    if (!phasenInput || !zeitInput) {
        debugLog("FEHLER: Input-Felder nicht gefunden!");
        return;
    }
    
    var phasen = parseInt(phasenInput.value) || 3;
    var zeiteinheiten = parseInt(zeitInput.value) || 12;
    var labelBreiteRE = parseInt(labelInput.value) || 20;
    
    debugLog("Phasen: " + phasen);
    debugLog("Zeiteinheiten: " + zeiteinheiten);
    debugLog("Label-Breite: " + labelBreiteRE + " RE");
    debugLog("---");
    debugLog("GANTT-Position: " + GANTT_LEFT_PT.toFixed(1) + " x " + GANTT_TOP_PT.toFixed(1) + " pt");
    
    // Validierung
    if (phasen < 1 || phasen > 20) {
        debugLog("FEHLER: Phasen muss zwischen 1 und 20 sein");
        return;
    }
    if (zeiteinheiten < 1 || zeiteinheiten > 24) {
        debugLog("FEHLER: Zeiteinheiten muss zwischen 1 und 24 sein");
        return;
    }
    
    // Spaltenbreite berechnen (verfügbare Breite nach Abzug Label)
    var verfuegbareBreiteRE = GANTT_MAX_WIDTH_RE - labelBreiteRE;
    var spaltenBreiteRE = Math.floor(verfuegbareBreiteRE / zeiteinheiten);
    
    if (spaltenBreiteRE < 1) {
        debugLog("FEHLER: Spaltenbreite zu klein! Weniger Zeiteinheiten verwenden.");
        return;
    }
    
    var gesamtBreiteRE = labelBreiteRE + (spaltenBreiteRE * zeiteinheiten);
    
    debugLog("Spaltenbreite: " + spaltenBreiteRE + " RE");
    debugLog("Gesamtbreite: " + gesamtBreiteRE + " RE");
    debugLog("---");
    
    // PowerPoint API aufrufen
    debugLog("Starte PowerPoint.run...");
    
    PowerPoint.run(function(context) {
        debugLog("PowerPoint.run aktiv");
        
        var shapes = context.presentation.slides.getItemAt(0).shapes;
        var shapesCreated = 0;
        
        // Berechnungen in Points
        var labelBreitePT = labelBreiteRE * RE_PT;
        var spaltenBreitePT = spaltenBreiteRE * RE_PT;
        var headerHoehePT = HEADER_HEIGHT_RE * RE_PT;
        var taskHoehePT = TASK_HEIGHT_RE * RE_PT;
        var taskAbstandPT = TASK_GAP_RE * RE_PT;
        
        // ===== LABEL-SPALTE (links) =====
        var labelShape = shapes.addGeometricShape(Office.GeometricShapeType.rectangle, {
            left: GANTT_LEFT_PT,
            top: GANTT_TOP_PT,
            width: labelBreitePT,
            height: headerHoehePT
        });
        labelShape.fill.setSolidColor("#F0F0F0");
        labelShape.lineFormat.color = "#808080";
        labelShape.lineFormat.weight = LINE_WEIGHT;
        labelShape.textFrame.textRange.text = "Phase";
        labelShape.textFrame.textRange.font.bold = true;
        labelShape.textFrame.textRange.font.size = 10;
        shapesCreated++;
        
        // ===== HEADER-ZEILE (Zeiteinheiten) =====
        var headerStartX = GANTT_LEFT_PT + labelBreitePT;
        
        for (var t = 0; t < zeiteinheiten; t++) {
            var headerX = headerStartX + (t * spaltenBreitePT);
            
            var headerShape = shapes.addGeometricShape(Office.GeometricShapeType.rectangle, {
                left: headerX,
                top: GANTT_TOP_PT,
                width: spaltenBreitePT,
                height: headerHoehePT
            });
            
            headerShape.fill.setSolidColor(COLORS.headerFill);
            headerShape.lineFormat.color = COLORS.headerBorder;
            headerShape.lineFormat.weight = LINE_WEIGHT;
            headerShape.textFrame.textRange.text = String(t + 1);
            headerShape.textFrame.textRange.font.color = COLORS.headerText;
            headerShape.textFrame.textRange.font.size = 9;
            headerShape.textFrame.textRange.font.bold = true;
            headerShape.textFrame.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;
            
            shapesCreated++;
        }
        
        debugLog("Header erstellt: " + zeiteinheiten + " Spalten");
        
        // ===== PHASEN-ZEILEN =====
        var taskStartY = GANTT_TOP_PT + headerHoehePT;
        var zeilenHoehePT = taskHoehePT + taskAbstandPT;
        
        for (var p = 0; p < phasen; p++) {
            var taskY = taskStartY + (p * zeilenHoehePT);
            
            // Label-Zelle
            var phaseLabelShape = shapes.addGeometricShape(Office.GeometricShapeType.rectangle, {
                left: GANTT_LEFT_PT,
                top: taskY,
                width: labelBreitePT,
                height: taskHoehePT
            });
            phaseLabelShape.fill.setSolidColor("#FFFFFF");
            phaseLabelShape.lineFormat.color = "#808080";
            phaseLabelShape.lineFormat.weight = LINE_WEIGHT;
            phaseLabelShape.textFrame.textRange.text = "Phase " + (p + 1);
            phaseLabelShape.textFrame.textRange.font.size = 9;
            phaseLabelShape.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
            shapesCreated++;
            
            // Task-Balken (erstreckt sich über alle Zeiteinheiten als Demo)
            var taskX = headerStartX;
            var taskBreitePT = spaltenBreitePT * zeiteinheiten;
            
            var taskShape = shapes.addGeometricShape(Office.GeometricShapeType.rectangle, {
                left: taskX,
                top: taskY,
                width: taskBreitePT,
                height: taskHoehePT
            });
            
            taskShape.fill.setSolidColor(COLORS.taskFill);
            taskShape.lineFormat.color = COLORS.taskBorder;
            taskShape.lineFormat.weight = LINE_WEIGHT;
            
            shapesCreated++;
        }
        
        debugLog("Phasen erstellt: " + phasen);
        debugLog("Shapes gesamt: " + shapesCreated);
        
        return context.sync();
    })
    .then(function() {
        debugLog("---");
        debugLog("=== ERFOLGREICH ===");
    })
    .catch(function(error) {
        debugLog("---");
        debugLog("FEHLER: " + error.message);
        if (error.debugInfo) {
            debugLog("Debug: " + JSON.stringify(error.debugInfo));
        }
    });
}
