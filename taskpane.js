// ========================================
// GANTT Generator v2.15
// DROEGE GROUP - Exakte Rasterpositionierung
// ========================================

// Globales Raster-System
const DROEGE_GRID = (function() {
    // Basis-Konstanten (NIEMALS runden!)
    const POINTS_PER_CM = 28.3464567;
    const RE_CM = 0.21;
    const RE_PT = RE_CM * POINTS_PER_CM;  // 5.95275590551...
    
    // Maximale Inhaltsbreite
    const MAX_WIDTH_RE = 118;
    
    // Dynamisch berechnete Werte
    let gridMarginLeft = 0;
    let gridMarginTop = 0;
    let slideWidth = 0;
    let slideHeight = 0;
    let isInitialized = false;
    
    return {
        // Konstanten
        RE: RE_PT,
        RE_CM: RE_CM,
        MAX_WIDTH_RE: MAX_WIDTH_RE,
        
        // Status
        isReady: function() {
            return isInitialized;
        },
        
        getSlideInfo: function() {
            return {
                width: slideWidth,
                height: slideHeight,
                marginLeft: gridMarginLeft,
                marginTop: gridMarginTop
            };
        },
        
        // Initialisierung - MUSS beim Start aufgerufen werden!
        initialize: async function() {
            try {
                await PowerPoint.run(async (context) => {
                    const presentation = context.presentation;
                    presentation.load("slideWidth,slideHeight");
                    await context.sync();
                    
                    slideWidth = presentation.slideWidth;
                    slideHeight = presentation.slideHeight;
                    
                    // Berechne wie viele GANZE RE passen
                    const fullUnitsX = Math.floor(slideWidth / RE_PT);
                    const fullUnitsY = Math.floor(slideHeight / RE_PT);
                    
                    // Berechne den Rest-Rand
                    const totalGridWidth = fullUnitsX * RE_PT;
                    const totalGridHeight = fullUnitsY * RE_PT;
                    
                    // Rest gleichmäßig auf beide Seiten verteilen
                    gridMarginLeft = (slideWidth - totalGridWidth) / 2;
                    gridMarginTop = (slideHeight - totalGridHeight) / 2;
                    
                    isInitialized = true;
                    
                    console.log("DROEGE_GRID initialisiert:");
                    console.log("  Foliengröße:", slideWidth.toFixed(2), "×", slideHeight.toFixed(2), "points");
                    console.log("  Raster-Margin:", gridMarginLeft.toFixed(4), "×", gridMarginTop.toFixed(4), "points");
                    console.log("  Ganze RE horizontal:", fullUnitsX);
                    console.log("  Ganze RE vertikal:", fullUnitsY);
                });
                
                return true;
            } catch (error) {
                console.error("Fehler bei Grid-Initialisierung:", error);
                isInitialized = false;
                return false;
            }
        },
        
        // RE → Points (für Größen, ohne Offset)
        toPoints: function(re) {
            return re * RE_PT;
        },
        
        // Absolute Position auf der Folie berechnen
        // reX, reY = Position im Raster (ab Offset)
        // offsetLeftRE, offsetTopRE = DROEGE-Inhaltsbereich-Offset
        position: function(reX, reY, offsetLeftRE, offsetTopRE) {
            offsetLeftRE = offsetLeftRE || 0;
            offsetTopRE = offsetTopRE || 0;
            
            return {
                left: gridMarginLeft + ((offsetLeftRE + reX) * RE_PT),
                top: gridMarginTop + ((offsetTopRE + reY) * RE_PT)
            };
        },
        
        // Größe in RE → Points
        size: function(widthRE, heightRE) {
            return {
                width: widthRE * RE_PT,
                height: heightRE * RE_PT
            };
        },
        
        // Komplette Shape-Optionen
        shapeOptions: function(reX, reY, widthRE, heightRE, offsetLeftRE, offsetTopRE) {
            const pos = this.position(reX, reY, offsetLeftRE, offsetTopRE);
            const sz = this.size(widthRE, heightRE);
            return {
                left: pos.left,
                top: pos.top,
                width: sz.width,
                height: sz.height
            };
        }
    };
})();

// DROEGE Farben
const COLORS = {
    DARK_BLUE: "#1A3A5C",
    LIGHT_BLUE: "#4A90C2",
    GRAY: "#808080",
    GOLD: "#C4A35A",
    WHITE: "#FFFFFF",
    LIGHT_GRAY: "#F0F0F0"
};

// ========================================
// UI-Funktionen
// ========================================

function updateStatus(message, type) {
    const statusBox = document.getElementById("gridStatus");
    const statusText = document.getElementById("statusText");
    const statusIcon = statusBox.querySelector(".status-icon");
    
    statusText.textContent = message;
    statusBox.classList.remove("ready", "error");
    
    if (type === "ready") {
        statusBox.classList.add("ready");
        statusIcon.textContent = "✅";
    } else if (type === "error") {
        statusBox.classList.add("error");
        statusIcon.textContent = "❌";
    } else {
        statusIcon.textContent = "⏳";
    }
}

function updateGridInfo() {
    const info = DROEGE_GRID.getSlideInfo();
    
    document.getElementById("slideInfo").textContent = 
        `Foliengröße: ${(info.width / 28.3464567).toFixed(2)} × ${(info.height / 28.3464567).toFixed(2)} cm`;
    
    document.getElementById("marginInfo").textContent = 
        `Raster-Margin: ${info.marginLeft.toFixed(4)} × ${info.marginTop.toFixed(4)} pt`;
}

function showError(message) {
    const errorDiv = document.getElementById("errorMessage");
    errorDiv.textContent = message;
    errorDiv.classList.remove("hidden");
    setTimeout(() => {
        errorDiv.classList.add("hidden");
    }, 5000);
}

// ========================================
// Initialisierung
// ========================================

Office.onReady(async (info) => {
    if (info.host === Office.HostType.PowerPoint) {
        console.log("Office.onReady - PowerPoint erkannt");
        
        // Event-Listener registrieren
        setupEventListeners();
        
        // Raster initialisieren
        await initializeGrid();
    }
});

async function initializeGrid() {
    updateStatus("Initialisiere Raster...", "loading");
    
    const success = await DROEGE_GRID.initialize();
    
    if (success) {
        updateStatus("Raster bereit!", "ready");
        updateGridInfo();
        document.getElementById("generateBtn").disabled = false;
    } else {
        updateStatus("Fehler bei Initialisierung", "error");
        document.getElementById("generateBtn").disabled = true;
    }
}

function setupEventListeners() {
    // Radio-Button für Spaltenbreite
    document.querySelectorAll('input[name="widthMode"]').forEach(radio => {
        radio.addEventListener("change", function() {
            const fixedSection = document.getElementById("fixedWidthSection");
            if (this.value === "fixed") {
                fixedSection.classList.remove("hidden");
            } else {
                fixedSection.classList.add("hidden");
            }
        });
    });
    
    // Generate Button
    document.getElementById("generateBtn").addEventListener("click", generateGantt);
    
    // Reinit Button
    document.getElementById("reinitBtn").addEventListener("click", initializeGrid);
}

// ========================================
// GANTT-Generierung
// ========================================

async function generateGantt() {
    if (!DROEGE_GRID.isReady()) {
        showError("Raster nicht initialisiert. Bitte neu initialisieren.");
        return;
    }
    
    // Parameter auslesen
    const columnCount = parseInt(document.getElementById("columnCount").value) || 12;
    const taskCount = parseInt(document.getElementById("taskCount").value) || 5;
    const taskHeight = parseInt(document.getElementById("taskHeight").value) || 3;
    const taskSpacing = parseInt(document.getElementById("taskSpacing").value) || 1;
    const offsetLeft = parseInt(document.getElementById("offsetLeft").value) || 9;
    const offsetTop = parseInt(document.getElementById("offsetTop").value) || 17;
    
    const widthMode = document.querySelector('input[name="widthMode"]:checked').value;
    const fixedWidth = parseInt(document.getElementById("fixedWidth").value) || 5;
    
    // Spaltenbreiten berechnen
    let columnWidths = [];
    if (widthMode === "auto") {
        // Gleichmäßige Verteilung auf 118 RE
        const widthPerColumn = Math.floor(DROEGE_GRID.MAX_WIDTH_RE / columnCount);
        columnWidths = Array(columnCount).fill(widthPerColumn);
    } else {
        // Feste Breite pro Spalte
        columnWidths = Array(columnCount).fill(fixedWidth);
    }
    
    // Gesamtbreite prüfen
    const totalWidth = columnWidths.reduce((a, b) => a + b, 0);
    if (totalWidth > DROEGE_GRID.MAX_WIDTH_RE) {
        showError(`Gesamtbreite (${totalWidth} RE) überschreitet Maximum (${DROEGE_GRID.MAX_WIDTH_RE} RE). Wird abgeschnitten.`);
    }
    
    try {
        await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            const shapes = slide.shapes;
            
            // Header-Zeile erstellen
            let currentX = 0;
            for (let col = 0; col < columnCount; col++) {
                // Prüfen ob noch Platz
                if (currentX + columnWidths[col] > DROEGE_GRID.MAX_WIDTH_RE) {
                    console.log(`Spalte ${col + 1} abgeschnitten (würde über ${DROEGE_GRID.MAX_WIDTH_RE} RE hinausgehen)`);
                    break;
                }
                
                const opts = DROEGE_GRID.shapeOptions(
                    currentX,           // reX
                    0,                  // reY (Header-Zeile)
                    columnWidths[col],  // width
                    2,                  // height (2 RE für Header)
                    offsetLeft,         // offsetLeftRE
                    offsetTop           // offsetTopRE
                );
                
                const headerShape = shapes.addGeometricShape(
                    PowerPoint.GeometricShapeType.rectangle,
                    opts
                );
                
                headerShape.name = `GANTT_Header_${col + 1}`;
                headerShape.fill.setSolidColor(COLORS.LIGHT_GRAY);
                headerShape.lineFormat.color = COLORS.GRAY;
                headerShape.lineFormat.weight = 0.5;
                
                // Text hinzufügen
                headerShape.textFrame.textRange.text = `KW ${col + 1}`;
                headerShape.textFrame.textRange.font.size = 10;
                headerShape.textFrame.textRange.font.color = COLORS.DARK_BLUE;
                headerShape.textFrame.textRange.font.bold = true;
                headerShape.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
                headerShape.textFrame.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;
                
                currentX += columnWidths[col];
            }
            
            // Tasks erstellen
            const headerHeight = 2;  // RE
            const startY = headerHeight + 1;  // 1 RE Abstand nach Header
            
            for (let task = 0; task < taskCount; task++) {
                // Task-Position berechnen
                const taskY = startY + (task * (taskHeight + taskSpacing));
                
                // Beispiel-Task: Startet bei Spalte task, Länge variiert
                const taskStartCol = task % columnCount;
                const taskLength = Math.min(3 + task, columnCount - taskStartCol);
                
                // Start-X berechnen (Summe der Spaltenbreiten bis taskStartCol)
                let taskStartX = 0;
                for (let i = 0; i < taskStartCol; i++) {
                    taskStartX += columnWidths[i];
                }
                
                // Breite berechnen
                let taskWidthRE = 0;
                for (let i = taskStartCol; i < taskStartCol + taskLength && i < columnCount; i++) {
                    taskWidthRE += columnWidths[i];
                }
                
                // Prüfen ob Task noch im Bereich
                if (taskStartX >= DROEGE_GRID.MAX_WIDTH_RE) continue;
                if (taskStartX + taskWidthRE > DROEGE_GRID.MAX_WIDTH_RE) {
                    taskWidthRE = DROEGE_GRID.MAX_WIDTH_RE - taskStartX;
                }
                
                const taskOpts = DROEGE_GRID.shapeOptions(
                    taskStartX,
                    taskY,
                    taskWidthRE,
                    taskHeight,
                    offsetLeft,
                    offsetTop
                );
                
                const taskShape = shapes.addGeometricShape(
                    PowerPoint.GeometricShapeType.rectangle,
                    taskOpts
                );
                
                taskShape.name = `GANTT_Task_${task + 1}`;
                taskShape.fill.setSolidColor(task % 2 === 0 ? COLORS.DARK_BLUE : COLORS.LIGHT_BLUE);
                taskShape.lineFormat.color = COLORS.DARK_BLUE;
                taskShape.lineFormat.weight = 0;
                
                // Task-Text
                taskShape.textFrame.textRange.text = `Task ${task + 1}`;
                taskShape.textFrame.textRange.font.size = 10;
                taskShape.textFrame.textRange.font.color = COLORS.WHITE;
                taskShape.textFrame.textRange.font.bold = false;
                taskShape.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
                taskShape.textFrame.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.left;
                taskShape.textFrame.leftMargin = 5;
            }
            
            await context.sync();
            console.log("GANTT-Diagramm erfolgreich erstellt!");
        });
        
        updateStatus("GANTT-Diagramm erstellt!", "ready");
        
    } catch (error) {
        console.error("Fehler beim Erstellen:", error);
        showError("Fehler: " + error.message);
        updateStatus("Fehler beim Erstellen", "error");
    }
}
