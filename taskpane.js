/* GANTT Generator v2.13 - DROEGE GROUP */

Office.onReady(function(info) {
    if (info.host === Office.HostType.PowerPoint) {
        document.getElementById("generateBtn").addEventListener("click", generateGantt);
        document.getElementById("addPhaseBtn").addEventListener("click", addPhaseRow);
        document.getElementById("widthMode").addEventListener("change", toggleWidthMode);
        initializeDefaults();
    }
});

function initializeDefaults() {
    const today = new Date();
    const startDate = new Date(today.getFullYear(), today.getMonth(), 1);
    const endDate = new Date(today.getFullYear(), today.getMonth() + 3, 0);
    
    document.getElementById("startDate").value = formatDate(startDate);
    document.getElementById("endDate").value = formatDate(endDate);
    
    addPhaseRow();
}

function toggleWidthMode() {
    const mode = document.getElementById("widthMode").value;
    const columnWidthSection = document.getElementById("columnWidthSection");
    columnWidthSection.style.display = mode === "fixed" ? "block" : "none";
}

function formatDate(date) {
    return date.toISOString().split("T")[0];
}

function addPhaseRow() {
    const container = document.getElementById("phaseContainer");
    const row = document.createElement("div");
    row.className = "phase-row";
    
    const colors = ["#0078d4", "#107c10", "#ff8c00", "#d13438", "#8764b8", "#00bcf2"];
    const colorIndex = container.children.length % colors.length;
    
    row.innerHTML = 
        '<input type="text" placeholder="Phasenname" class="phase-name">' +
        '<input type="date" class="phase-start">' +
        '<input type="date" class="phase-end">' +
        '<input type="color" class="phase-color" value="' + colors[colorIndex] + '">' +
        '<button class="remove-btn" onclick="this.parentElement.remove()">×</button>';
    
    container.appendChild(row);
}

// Constants
const CM = 28.3465; // Points per cm
const GRID_UNIT_CM = 0.21; // 1 RE = 0.21 cm

// Fixed positioning (in RE)
const LEFT_RE = 9;
const TOP_RE = 17;
const MAX_WIDTH_RE = 118;

// Convert RE to Points
function re2pt(re) {
    return Math.round(re * GRID_UNIT_CM * CM);
}

async function generateGantt() {
    const status = document.getElementById("status");
    status.className = "status info";
    status.textContent = "Erstelle GANTT-Diagramm...";
    
    try {
        const timeUnit = document.getElementById("timeUnit").value;
        const widthMode = document.getElementById("widthMode").value;
        const columnWidthRE = parseInt(document.getElementById("columnWidth").value);
        const startDate = new Date(document.getElementById("startDate").value);
        const endDate = new Date(document.getElementById("endDate").value);
        
        if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
            throw new Error("Bitte gültige Start- und Enddaten eingeben.");
        }
        
        if (endDate <= startDate) {
            throw new Error("Enddatum muss nach Startdatum liegen.");
        }
        
        // Collect phases
        const phases = [];
        const phaseRows = document.querySelectorAll(".phase-row");
        phaseRows.forEach(function(row) {
            const name = row.querySelector(".phase-name").value.trim();
            const pStart = row.querySelector(".phase-start").value;
            const pEnd = row.querySelector(".phase-end").value;
            const color = row.querySelector(".phase-color").value;
            
            if (name && pStart && pEnd) {
                phases.push({
                    name: name,
                    start: new Date(pStart),
                    end: new Date(pEnd),
                    color: color
                });
            }
        });
        
        // Generate time periods
        const periods = generatePeriods(startDate, endDate, timeUnit);
        
        // Calculate column width
        let colWidthRE;
        if (widthMode === "auto") {
            // Auto-distribute: fit as many columns as possible within MAX_WIDTH_RE
            colWidthRE = Math.floor(MAX_WIDTH_RE / periods.length);
            if (colWidthRE < 1) colWidthRE = 1;
            if (colWidthRE > 10) colWidthRE = 10;
        } else {
            colWidthRE = columnWidthRE;
        }
        
        // Calculate actual width used (may truncate at end)
        const totalWidthRE = colWidthRE * periods.length;
        const usedWidthRE = Math.min(totalWidthRE, MAX_WIDTH_RE);
        const visibleColumns = Math.floor(usedWidthRE / colWidthRE);
        
        await PowerPoint.run(async function(context) {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            
            // Fixed starting position
            const startX = re2pt(LEFT_RE);
            const startY = re2pt(TOP_RE);
            const colWidth = re2pt(colWidthRE);
            const rowHeight = re2pt(3); // 3 RE row height
            
            // Determine if we need a month row
            const showMonthRow = (timeUnit === "days" || timeUnit === "weeks" || timeUnit === "quarters");
            
            // Header row(s) height
            const headerRows = showMonthRow ? 2 : 1;
            const headerHeight = rowHeight * headerRows;
            
            // Total height including phases
            const totalRows = headerRows + phases.length;
            const totalHeight = rowHeight * totalRows;
            
            // Draw month row if needed
            if (showMonthRow) {
                const monthGroups = groupPeriodsByMonth(periods.slice(0, visibleColumns));
                let monthX = startX;
                
                for (let i = 0; i < monthGroups.length; i++) {
                    const group = monthGroups[i];
                    const groupWidth = colWidth * group.count;
                    
                    // Check if we exceed max width
                    if (monthX + groupWidth - startX > re2pt(MAX_WIDTH_RE)) break;
                    
                    // Month cell background
                    const monthCell = slide.shapes.addGeometricShape(
                        PowerPoint.GeometricShapeType.rectangle,
                        {
                            left: monthX,
                            top: startY,
                            width: groupWidth,
                            height: rowHeight
                        }
                    );
                    monthCell.fill.setSolidColor("#d0d0d0");
                    monthCell.lineFormat.color = "#666666";
                    monthCell.lineFormat.weight = 0.5;
                    
                    // Month label
                    const monthLabel = slide.shapes.addTextBox(group.label, {
                        left: monthX,
                        top: startY,
                        width: groupWidth,
                        height: rowHeight
                    });
                    monthLabel.textFrame.textRange.font.size = 11;
                    monthLabel.textFrame.textRange.font.bold = true;
                    monthLabel.textFrame.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;
                    monthLabel.fill.clear();
                    monthLabel.lineFormat.visible = false;
                    
                    monthX += groupWidth;
                }
            }
            
            // Draw time unit header row
            const timeHeaderY = showMonthRow ? startY + rowHeight : startY;
            
            for (let i = 0; i < visibleColumns; i++) {
                const period = periods[i];
                const x = startX + (i * colWidth);
                
                // Check max width
                if (x + colWidth - startX > re2pt(MAX_WIDTH_RE)) break;
                
                // Header cell background
                const headerCell = slide.shapes.addGeometricShape(
                    PowerPoint.GeometricShapeType.rectangle,
                    {
                        left: x,
                        top: timeHeaderY,
                        width: colWidth,
                        height: rowHeight
                    }
                );
                headerCell.fill.setSolidColor("#e0e0e0");
                headerCell.lineFormat.color = "#666666";
                headerCell.lineFormat.weight = 0.5;
                
                // Header label
                const headerLabel = slide.shapes.addTextBox(period.label, {
                    left: x,
                    top: timeHeaderY,
                    width: colWidth,
                    height: rowHeight
                });
                headerLabel.textFrame.textRange.font.size = 11;
                headerLabel.textFrame.textRange.font.bold = true;
                headerLabel.textFrame.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;
                headerLabel.fill.clear();
                headerLabel.lineFormat.visible = false;
            }
            
            // Draw phase rows
            const phaseStartY = startY + headerHeight;
            
            for (let p = 0; p < phases.length; p++) {
                const phase = phases[p];
                const rowY = phaseStartY + (p * rowHeight);
                
                // Find start and end columns for this phase
                let phaseStartCol = -1;
                let phaseEndCol = -1;
                
                for (let i = 0; i < visibleColumns; i++) {
                    const period = periods[i];
                    if (phase.start <= period.end && phase.end >= period.start) {
                        if (phaseStartCol === -1) phaseStartCol = i;
                        phaseEndCol = i;
                    }
                }
                
                if (phaseStartCol >= 0 && phaseEndCol >= 0) {
                    const barX = startX + (phaseStartCol * colWidth);
                    const barWidth = (phaseEndCol - phaseStartCol + 1) * colWidth;
                    
                    // Phase bar
                    const phaseBar = slide.shapes.addGeometricShape(
                        PowerPoint.GeometricShapeType.rectangle,
                        {
                            left: barX,
                            top: rowY,
                            width: barWidth,
                            height: rowHeight
                        }
                    );
                    phaseBar.fill.setSolidColor(phase.color);
                    phaseBar.lineFormat.color = "#333333";
                    phaseBar.lineFormat.weight = 0.5;
                    
                    // Phase label
                    const phaseLabel = slide.shapes.addTextBox(phase.name, {
                        left: barX,
                        top: rowY,
                        width: barWidth,
                        height: rowHeight
                    });
                    phaseLabel.textFrame.textRange.font.size = 11;
                    phaseLabel.textFrame.textRange.font.bold = true;
                    phaseLabel.textFrame.textRange.font.color = "#ffffff";
                    phaseLabel.textFrame.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;
                    phaseLabel.fill.clear();
                    phaseLabel.lineFormat.visible = false;
                }
            }
            
            // Draw vertical grid lines
            for (let i = 0; i <= visibleColumns; i++) {
                const x = startX + (i * colWidth);
                
                // Check max width
                if (x - startX > re2pt(MAX_WIDTH_RE)) break;
                
                const gridLine = slide.shapes.addLine(
                    PowerPoint.ConnectorType.straight,
                    {
                        left: x,
                        top: startY,
                        width: 0.01,
                        height: totalHeight
                    }
                );
                gridLine.lineFormat.color = "#666666";
                gridLine.lineFormat.weight = 0.5;
            }
            
            // Draw today line
            const today = new Date();
            today.setHours(0, 0, 0, 0);
            
            for (let i = 0; i < visibleColumns; i++) {
                const period = periods[i];
                if (today >= period.start && today <= period.end) {
                    // Calculate position within the period
                    const periodDuration = period.end.getTime() - period.start.getTime();
                    const todayOffset = today.getTime() - period.start.getTime();
                    const ratio = periodDuration > 0 ? todayOffset / periodDuration : 0.5;
                    
                    const todayX = startX + (i * colWidth) + (colWidth * ratio);
                    
                    // Today line (red)
                    const todayLine = slide.shapes.addLine(
                        PowerPoint.ConnectorType.straight,
                        {
                            left: todayX,
                            top: startY,
                            width: 0.01,
                            height: totalHeight + re2pt(2)
                        }
                    );
                    todayLine.lineFormat.color = "#ff0000";
                    todayLine.lineFormat.weight = 1.5;
                    
                    // "Heute" label
                    const heuteLabel = slide.shapes.addTextBox("Heute", {
                        left: todayX - re2pt(3),
                        top: startY + totalHeight + re2pt(1),
                        width: re2pt(6),
                        height: re2pt(2)
                    });
                    heuteLabel.textFrame.textRange.font.size = 11;
                    heuteLabel.textFrame.textRange.font.color = "#ff0000";
                    heuteLabel.textFrame.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;
                    heuteLabel.fill.clear();
                    heuteLabel.lineFormat.visible = false;
                    
                    break;
                }
            }
            
            await context.sync();
        });
        
        status.className = "status success";
        let statusText = "GANTT-Diagramm erstellt! ";
        statusText += "Spaltenbreite: " + colWidthRE + " RE, ";
        statusText += "Sichtbare Spalten: " + visibleColumns + "/" + periods.length;
        if (visibleColumns < periods.length) {
            statusText += " (abgeschnitten)";
        }
        status.textContent = statusText;
        
    } catch (error) {
        status.className = "status error";
        status.textContent = "Fehler: " + error.message;
        console.error(error);
    }
}

function generatePeriods(startDate, endDate, timeUnit) {
    const periods = [];
    const current = new Date(startDate);
    
    switch (timeUnit) {
        case "days":
            while (current <= endDate) {
                const dayStart = new Date(current);
                const dayEnd = new Date(current);
                dayEnd.setHours(23, 59, 59, 999);
                
                periods.push({
                    label: current.getDate().toString(),
                    start: dayStart,
                    end: dayEnd,
                    month: current.getMonth(),
                    year: current.getFullYear()
                });
                
                current.setDate(current.getDate() + 1);
            }
            break;
            
        case "weeks":
            // Start from Monday of the week containing startDate
            const dayOfWeek = current.getDay();
            const diffToMonday = (dayOfWeek === 0 ? -6 : 1) - dayOfWeek;
            current.setDate(current.getDate() + diffToMonday);
            
            while (current <= endDate) {
                const weekStart = new Date(current);
                const weekEnd = new Date(current);
                weekEnd.setDate(weekEnd.getDate() + 6);
                weekEnd.setHours(23, 59, 59, 999);
                
                // Calculate week number
                const jan1 = new Date(current.getFullYear(), 0, 1);
                const days = Math.floor((current - jan1) / (24 * 60 * 60 * 1000));
                const weekNum = Math.ceil((days + jan1.getDay() + 1) / 7);
                
                periods.push({
                    label: "KW" + weekNum,
                    start: weekStart,
                    end: weekEnd,
                    month: current.getMonth(),
                    year: current.getFullYear()
                });
                
                current.setDate(current.getDate() + 7);
            }
            break;
            
        case "months":
            current.setDate(1);
            while (current <= endDate) {
                const monthStart = new Date(current);
                const monthEnd = new Date(current.getFullYear(), current.getMonth() + 1, 0, 23, 59, 59, 999);
                
                const monthNames = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", 
                                    "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"];
                
                periods.push({
                    label: monthNames[current.getMonth()] + " " + current.getFullYear().toString().slice(-2),
                    start: monthStart,
                    end: monthEnd,
                    month: current.getMonth(),
                    year: current.getFullYear()
                });
                
                current.setMonth(current.getMonth() + 1);
            }
            break;
            
        case "quarters":
            // Start from beginning of quarter containing startDate
            const startQuarter = Math.floor(current.getMonth() / 3);
            current.setMonth(startQuarter * 3, 1);
            
            while (current <= endDate) {
                const qStart = new Date(current);
                const qEnd = new Date(current.getFullYear(), current.getMonth() + 3, 0, 23, 59, 59, 999);
                
                const qNum = Math.floor(current.getMonth() / 3) + 1;
                
                periods.push({
                    label: "Q" + qNum,
                    start: qStart,
                    end: qEnd,
                    month: current.getMonth(),
                    year: current.getFullYear()
                });
                
                current.setMonth(current.getMonth() + 3);
            }
            break;
    }
    
    return periods;
}

function groupPeriodsByMonth(periods) {
    const groups = [];
    const monthNames = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun",
                        "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"];
    
    let currentGroup = null;
    
    for (let i = 0; i < periods.length; i++) {
        const period = periods[i];
        const monthKey = period.year + "-" + period.month;
        
        if (!currentGroup || currentGroup.key !== monthKey) {
            if (currentGroup) {
                groups.push(currentGroup);
            }
            currentGroup = {
                key: monthKey,
                label: monthNames[period.month],
                count: 1
            };
        } else {
            currentGroup.count++;
        }
    }
    
    if (currentGroup) {
        groups.push(currentGroup);
    }
    
    return groups;
}
