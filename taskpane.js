// GANTT Generator v2.20 - Fixed slide dimensions
Office.onReady(function(info) {
    if (info.host === Office.HostType.PowerPoint) {
        document.getElementById("createGantt").onclick = generateGantt;
        document.getElementById("debugOutput").innerText = "GANTT v2.20 bereit (feste Folienmaße)";
    }
});

// ============================================================
// KONSTANTEN
// ============================================================
const POINTS_PER_CM = 28.3464567;
const RE_CM = 0.21;
const RE_PT = RE_CM * POINTS_PER_CM; // 5.95275590551

// Feste Foliengröße für 16:9 (PowerPoint Standard)
const SLIDE_WIDTH_PT = 960;   // 33.867 cm
const SLIDE_HEIGHT_PT = 540;  // 19.05 cm

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

// ============================================================
// BERECHNETE GRID-MARGINS (zentriertes Raster)
// ============================================================
const FULL_UNITS_X = Math.floor(SLIDE_WIDTH_PT / RE_PT);  // 161
const FULL_UNITS_Y = Math.floor(SLIDE_HEIGHT_PT / RE_PT); // 90
const MARGIN_LEFT_PT = (SLIDE_WIDTH_PT - (FULL_UNITS_X * RE_PT)) / 2;  // ~0.808
const MARGIN_TOP_PT = (SLIDE_HEIGHT_PT - (FULL_UNITS_Y * RE_PT)) / 2;  // ~2.126

// Absolute GANTT-Position in Points
const GANTT_LEFT_PT = MARGIN_LEFT_PT + (GANTT_LEFT_RE * RE_PT);  // ~54.38
const GANTT_TOP_PT = MARGIN_TOP_PT + (GANTT_TOP_RE * RE_PT);     // ~103.32

// ============================================================
// DEBUG OUTPUT
// ============================================================
function debugLog(msg) {
    var out = document.getElementById("debugOutput");
    out.innerText += "\n" + msg;
    console.log(msg);
}

// ============================================================
// MAIN FUNCTION
// ============================================================
function generateGantt() {
    var debugOut = document.getElementById("debugOutput");
    debugOut.innerText = "=== GANTT v2.20 Start ===";
    
    // Eingaben lesen
    var phasenInput = document.getElementById("phasen").value.trim();
    var zeitInput = document.getElementById("zeiteinheiten").value.trim();
    
    var phasen = parseInt(phasenInput) || 0;
    var zeiteinheiten = parseInt(zeitInput) || 0;
    
    debugLog("Phasen: " + phasen + ", Zeiteinheiten: " + zeiteinheiten);
    debugLog("Slide: " + SLIDE_WIDTH_PT + " x " + SLIDE_HEIGHT_PT + " pt");
    debugLog("Grid-Einheiten: " + FULL_UNITS_X + " x " + FULL_UNITS_Y);
    debugLog("Margin: " + MARGIN_LEFT_PT.toFixed(3) + " x " + MARGIN_TOP_PT.toFixed(3) + " pt");
    debugLog("GANTT-Position: " + GANTT_LEFT_PT.toFixed(2) + " x " + GANTT_TOP_PT.toFixed(2) + " pt");
    
    if (phasen < 1 || zeiteinheiten < 1) {
        debugLog("FEHLER: Ungültige Eingaben");
        return;
    }
    
    // Spaltenbreite berechnen
    var columnWidthRE = Math.floor(GANTT_MAX_WIDTH_RE / zeiteinheiten);
    var totalWidthRE = columnWidthRE * zeiteinheiten;
    var totalWidthPT = totalWidthRE * RE_PT;
    
    debugLog("Spaltenbreite: " + columnWidthRE + " RE");
    debugLog("Gesamtbreite: " + totalWidthRE + " RE = " + totalWidthPT.toFixed(2) + " pt");
    
    // PowerPoint API aufrufen
    PowerPoint.run(function(context) {
        debugLog("PowerPoint.run gestartet...");
        
        var shapes = context.presentation.slides.getItemAt(0).shapes;
        var shapesCreated = 0;
        
        // ----- HEADER-ZEILE (Zeiteinheiten) -----
        var headerY = GANTT_TOP_PT;
        var headerHeightPT = HEADER_HEIGHT_RE * RE_PT;
        
        for (var t = 0; t < zeiteinheiten; t++) {
            var headerX = GANTT_LEFT_PT + (t * columnWidthRE * RE_PT);
            var headerWidthPT = columnWidthRE * RE_PT;
            
            var headerShape = shapes.addGeometricShape(Office.GeometricShapeType.rectangle, {
                left: headerX,
                top: headerY,
                width: headerWidthPT,
                height: headerHeightPT
            });
            
            headerShape.fill.setSolidColor(COLORS.headerFill);
            headerShape.lineFormat.color = COLORS.headerBorder;
            headerShape.lineFormat.weight = LINE_WEIGHT;
            
            headerShape.textFrame.textRange.text = String(t + 1);
            headerShape.textFrame.textRange.font.color = COLORS.headerText;
            headerShape.textFrame.textRange.font.size = 10;
            headerShape.textFrame.textRange.font.bold = true;
            headerShape.textFrame.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;
            
            shapesCreated++;
        }
        
        debugLog("Header erstellt: " + zeiteinheiten + " Spalten");
        
        // ----- TASK-BALKEN (Phasen) -----
        var taskStartY = GANTT_TOP_PT + headerHeightPT + (TASK_GAP_RE * RE_PT);
        var taskHeightPT = TASK_HEIGHT_RE * RE_PT;
        var taskGapPT = TASK_GAP_RE * RE_PT;
        
        for (var p = 0; p < phasen; p++) {
            var taskY = taskStartY + (p * (taskHeightPT + taskGapPT));
            var taskX = GANTT_LEFT_PT;
            
            // Task erstreckt sich über alle Zeiteinheiten (Beispiel)
            var taskWidthPT = totalWidthPT;
            
            var taskShape = shapes.addGeometricShape(Office.GeometricShapeType.rectangle, {
                left: taskX,
                top: taskY,
                width: taskWidthPT,
                height: taskHeightPT
            });
            
            taskShape.fill.setSolidColor(COLORS.taskFill);
            taskShape.lineFormat.color = COLORS.taskBorder;
            taskShape.lineFormat.weight = LINE_WEIGHT;
            
            taskShape.textFrame.textRange.text = "Phase " + (p + 1);
            taskShape.textFrame.textRange.font.color = "#FFFFFF";
            taskShape.textFrame.textRange.font.size = 10;
            taskShape.textFrame.textRange.font.bold = true;
            taskShape.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
            
            shapesCreated++;
        }
        
        debugLog("Tasks erstellt: " + phasen + " Phasen");
        debugLog("Shapes gesamt: " + shapesCreated);
        
        return context.sync();
    })
    .then(function() {
        debugLog("=== ERFOLGREICH ===");
    })
    .catch(function(error) {
        debugLog("FEHLER: " + error.message);
        console.error(error);
    });
}
