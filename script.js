let lastSelectedCell = null; // Store last selected cell
let isDragging = false; // Flag to check if dragging
let startCell = null; // Store the initial cell
let selectedCells = []; // Store selected cells for multi-selection
let cellDependencies = {}; // Stores formula dependencies
document.addEventListener("DOMContentLoaded", function () {
    const tbody = document.querySelector("#spreadsheet tbody");

    // Create 10 rows dynamically
    for (let i = 1; i <= 10; i++) {
        let row = document.createElement("tr");

        // Row number column
        let th = document.createElement("th");
        th.innerText = i;
        row.appendChild(th);

        // Create 4 columns (A, B, C, D)
        for (let j = 0; j < 4; j++) {
            let cell = document.createElement("td");
            let input = document.createElement("input");
            input.setAttribute("data-row", i);
            input.setAttribute("data-col", String.fromCharCode(65 + j)); // A, B, C, D

            // Dragging functionality
            input.addEventListener("mousedown", function () {
                lastSelectedCell = input;
                isDragging = true;
                startCell = input;
                selectedCells = [input];
            });

            input.addEventListener("mouseover", function () {
                if (isDragging) {
                    selectedCells.push(input);
                    input.style.backgroundColor = "#e8f0fe"; // Highlight selected cells
                }
            });

            input.addEventListener("mouseup", function () {
                isDragging = false;
                if (startCell && selectedCells.length > 1) {
                    applyDragContent();
                }
            });

            input.addEventListener("blur", handleFormula);
            input.addEventListener("input", saveData);
            cell.appendChild(input);
            row.appendChild(cell);
        }
        tbody.appendChild(row);
    }

    loadData();
});

// Apply dragged content to selected cells
function applyDragContent() {
    if (!startCell || selectedCells.length < 2) return;

    let value = startCell.value; // Get original cell value

    selectedCells.forEach(cell => {
        if (value.startsWith("=")) {
            // Adjust formula references dynamically
            let newFormula = value.replace(/\d+/g, match => {
                return parseInt(match) + (parseInt(cell.getAttribute("data-row")) - parseInt(startCell.getAttribute("data-row")));
            });
            cell.value = newFormula;
        } else {
            cell.value = value;
        }

        cell.style.backgroundColor = "";
    });

    saveData();
}

// Handle formulas
function handleFormula(event) {
    let input = event.target;
    let value = input.value.trim();
    let cellKey = input.getAttribute("data-col") + input.getAttribute("data-row");

    if (value.startsWith("=")) {
        try {
            let formula = value.substring(1); // Remove '='
            let result = evaluateFormula(formula, cellKey);
            input.value = result;

            // Store formula for future updates
            input.setAttribute("data-formula", formula);

            // Track dependencies
            trackDependencies(cellKey, formula);
        } catch (error) {
            input.value = "Error";
        }
    }
    saveData();
}


// Track dependencies for auto-updating formulas
function trackDependencies(cellKey, formula) {
    cellDependencies[cellKey] = [];

    let cellRefs = formula.match(/[A-D]\d+/g); // Match cell references (e.g., A1, B3)
    if (cellRefs) {
        cellDependencies[cellKey] = cellRefs;
    }
}

// Auto-update dependent formulas
function updateDependentCells(changedCellKey) {
    Object.keys(cellDependencies).forEach(dependentCell => {
        if (cellDependencies[dependentCell].includes(changedCellKey)) {
            let input = document.querySelector(`input[data-col="${dependentCell.charAt(0)}"][data-row="${dependentCell.slice(1)}"]`);
            if (input) {
                let formula = "=" + input.dataset.formula; // Recalculate formula
                input.value = evaluateFormula(formula);
            }
        }
    });
}

// Update formula when cell value changes
document.querySelectorAll("input").forEach(input => {
    input.addEventListener("input", function () {
        let cellKey = input.getAttribute("data-col") + input.getAttribute("data-row");
        updateDependentCells(cellKey);
    });
});


// Evaluate formulas like SUM(A1:A3), AVERAGE(A1:A3), etc.
function evaluateFormula(formula, currentCell = "") {
    try {
        // Handle SUM, AVERAGE, MAX, MIN, COUNT functions
        let match;
        if ((match = formula.match(/SUM\((\w)(\d+):(\w)(\d+)\)/))) return calculateRange(match, "SUM");
        if ((match = formula.match(/AVERAGE\((\w)(\d+):(\w)(\d+)\)/))) return calculateRange(match, "AVERAGE");
        if ((match = formula.match(/MAX\((\w)(\d+):(\w)(\d+)\)/))) return calculateRange(match, "MAX");
        if ((match = formula.match(/MIN\((\w)(\d+):(\w)(\d+)\)/))) return calculateRange(match, "MIN");
        if ((match = formula.match(/COUNT\((\w)(\d+):(\w)(\d+)\)/))) return calculateRange(match, "COUNT");

        // Replace cell references with actual values (supports A1, $A$1)
        formula = formula.replace(/\$?([A-D])\$?(\d+)/g, (match, col, row) => {
            let cell = document.querySelector(`input[data-row="${row}"][data-col="${col}"]`);
            return cell && !isNaN(cell.value) ? parseFloat(cell.value) : 0;
        });

        return eval(formula); // Evaluate the mathematical expression
    } catch (error) {
        return "Error";
    }
}


// Calculate SUM, AVERAGE, MAX, etc.
function calculateRange(match, type) {
    let [_, startCol, startRow, endCol, endRow] = match;
    startRow = parseInt(startRow);
    endRow = parseInt(endRow);

    let values = [];
    for (let i = startRow; i <= endRow; i++) {
        let cell = document.querySelector(`input[data-row="${i}"][data-col="${startCol}"]`);
        if (cell && cell.value) {
            values.push(parseFloat(cell.value) || 0);
        }
    }

    if (type === "SUM") return values.reduce((a, b) => a + b, 0);
    if (type === "AVERAGE") return values.length ? values.reduce((a, b) => a + b, 0) / values.length : 0;
    if (type === "MAX") return Math.max(...values);
    if (type === "MIN") return Math.min(...values);
    if (type === "COUNT") return values.length;
}
function saveSpreadsheet() {
    let spreadsheetName = prompt("Enter a name for your spreadsheet:");
    if (!spreadsheetName) return;

    let data = {};
    document.querySelectorAll("input").forEach(input => {
        let row = input.getAttribute("data-row");
        let col = input.getAttribute("data-col");
        let key = `${col}${row}`;
        data[key] = {
            value: input.value || "",
            formula: input.getAttribute("data-formula") || ""
        };
    });

    localStorage.setItem(spreadsheetName, JSON.stringify(data));
    alert(`Spreadsheet "${spreadsheetName}" saved successfully!`);
}
function loadSpreadsheet() {
    let spreadsheetName = prompt("Enter the name of the spreadsheet to load:");
    if (!spreadsheetName) return;

    let data = JSON.parse(localStorage.getItem(spreadsheetName));
    if (!data) {
        alert("No spreadsheet found with that name.");
        return;
    }

    document.querySelectorAll("input").forEach(input => {
        let row = input.getAttribute("data-row");
        let col = input.getAttribute("data-col");
        let key = `${col}${row}`;

        if (data[key]) {
            input.value = data[key].value;
            if (data[key].formula) {
                input.setAttribute("data-formula", data[key].formula);
            }
        }
    });

    alert(`Spreadsheet "${spreadsheetName}" loaded successfully!`);
}

// Save & Load Data
function saveData() {
    let data = {};
    document.querySelectorAll("input").forEach(input => {
        let row = input.getAttribute("data-row");
        let col = input.getAttribute("data-col");
        let key = `${col}${row}`;
        data[key] = {
            value: input.value // Store only the text, not an object
        };
    });
    localStorage.setItem("spreadsheetData", JSON.stringify(data));
}


function loadData() {
    let data = JSON.parse(localStorage.getItem("spreadsheetData"));
    if (!data) return;

    document.querySelectorAll("input").forEach(input => {
        let row = input.getAttribute("data-row");
        let col = input.getAttribute("data-col");
        let key = `${col}${row}`;

        if (data[key] && typeof data[key] === "object") {
            input.value = data[key].value || ""; // Extract value correctly
        }
    });
}



// Clear Sheet
function clearSheet() {
    document.querySelectorAll("input").forEach(input => {
        input.value = "";
    });
    localStorage.removeItem("spreadsheetData");
    location.reload();
}
// Function to format text (Bold, Color, Font Size)
function formatText(type) {
    if (!lastSelectedCell) {
        alert("Click on a cell first before applying formatting!");
        return;
    }

    if (type === "bold") {
        lastSelectedCell.style.fontWeight = lastSelectedCell.style.fontWeight === "bold" ? "normal" : "bold";
    }
    if (type === "color") {
        let color = prompt("Enter a color (e.g., red, blue, #ff0000):", "black");
        if (color) lastSelectedCell.style.color = color;
    }
    if (type === "size") {
        let size = prompt("Enter font size (e.g., 14px, 20px):", "16px");
        if (size) lastSelectedCell.style.fontSize = size;
    }

    saveData();
}

// Function to reset formatting
function resetFormatting() {
    if (!lastSelectedCell) {
        alert("Click on a cell first before resetting formatting!");
        return;
    }

    lastSelectedCell.style.fontWeight = "normal";
    lastSelectedCell.style.color = "black";
    lastSelectedCell.style.fontSize = "14px";

    saveData();
}
function formatText(type) {
    if (!lastSelectedCell) {
        alert("Click on a cell first before applying formatting!");
        return;
    }

    if (type === "bold") {
        lastSelectedCell.style.fontWeight = lastSelectedCell.style.fontWeight === "bold" ? "normal" : "bold";
    }
    if (type === "italic") {
        lastSelectedCell.style.fontStyle = lastSelectedCell.style.fontStyle === "italic" ? "normal" : "italic";
    }
    if (type === "color") {
        let color = prompt("Enter a color (e.g., red, blue, #ff0000):", "black");
        if (color) lastSelectedCell.style.color = color;
    }
    if (type === "size") {
        let size = prompt("Enter font size (e.g., 14px, 20px):", "16px");
        if (size) lastSelectedCell.style.fontSize = size;
    }

    saveData();
}
function addRow() {
    let tbody = document.querySelector("#spreadsheet tbody");
    let rowCount = tbody.rows.length + 1;
    let row = document.createElement("tr");

    let th = document.createElement("th");
    th.innerText = rowCount;
    row.appendChild(th);

    for (let j = 0; j < 4; j++) {
        let cell = document.createElement("td");
        let input = document.createElement("input");
        input.setAttribute("data-row", rowCount);
        input.setAttribute("data-col", String.fromCharCode(65 + j));
        input.addEventListener("blur", handleFormula);
        input.addEventListener("input", saveData);
        cell.appendChild(input);
        row.appendChild(cell);
    }
    tbody.appendChild(row);
}
function deleteRow() {
    let tbody = document.querySelector("#spreadsheet tbody");
    if (tbody.rows.length > 1) {
        tbody.deleteRow(-1);
    } else {
        alert("At least one row must remain!");
    }
}
document.querySelectorAll("th").forEach(header => {
    header.style.cursor = "col-resize";
    header.addEventListener("mousedown", function (e) {
        let startX = e.clientX;
        let startWidth = header.offsetWidth;

        function resizeColumn(event) {
            let newWidth = startWidth + (event.clientX - startX);
            header.style.width = newWidth + "px";
        }

        function stopResize() {
            document.removeEventListener("mousemove", resizeColumn);
            document.removeEventListener("mouseup", stopResize);
        }

        document.addEventListener("mousemove", resizeColumn);
        document.addEventListener("mouseup", stopResize);
    });
});

// Export to Excel
function exportToExcel() {
    let table = document.createElement("table");
    let thead = document.createElement("thead");
    let tbody = document.createElement("tbody");

    let headerRow = document.createElement("tr");
    headerRow.appendChild(document.createElement("th"));
    ["A", "B", "C", "D"].forEach(letter => {
        let th = document.createElement("th");
        th.innerText = letter;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    for (let i = 1; i <= 10; i++) {
        let row = document.createElement("tr");
        let th = document.createElement("th");
        th.innerText = i;
        row.appendChild(th);

        for (let j = 0; j < 4; j++) {
            let cell = document.createElement("td");
            let input = document.querySelector(`input[data-row="${i}"][data-col="${String.fromCharCode(65 + j)}"]`);
            cell.innerText = input ? input.value : "";
            row.appendChild(cell);
        }
        tbody.appendChild(row);
    }

    table.appendChild(thead);
    table.appendChild(tbody);
    let ws = XLSX.utils.table_to_sheet(table);
    let wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    XLSX.writeFile(wb, "Google_Sheets_Clone.xlsx");
}
function generateChart() {
    let column = document.getElementById("columnSelect").value; // Get selected column
    let chartType = document.getElementById("chartTypeSelect").value; // Get selected chart type
    let labels = [];
    let values = [];

    // Get all input fields from the selected column
    document.querySelectorAll(`input[data-col='${column}']`).forEach(input => {
        let row = input.getAttribute("data-row");
        let value = parseFloat(input.value.trim()); // Convert to number

        if (!isNaN(value) && input.value.trim() !== "") { // Ensure valid number input
            labels.push("Row " + row);
            values.push(value);
        }
    });

    // ✅ Debugging: Check if data is being collected properly
    console.log("Chart Labels:", labels);
    console.log("Chart Values:", values);

    // ❌ Prevents chart creation if no valid data is found
    if (values.length === 0) {
        alert("No valid numerical data found in the selected column!");
        return;
    }

    // ✅ Check if the canvas exists before drawing the chart
    let canvas = document.getElementById("myChart");
    if (!canvas) {
        alert("Error: Chart canvas not found!");
        return;
    }

    let ctx = canvas.getContext("2d");

    // ✅ Fix: Ensure the chart is an instance of Chart.js before destroying it
    if (window.myChart instanceof Chart) {
        window.myChart.destroy();
    }

    // ✅ Create a new chart with the selected type
    window.myChart = new Chart(ctx, {
        type: chartType, // Dynamically select chart type
        data: {
            labels: labels,
            datasets: [{
                label: `Values in Column ${column}`,
                data: values,
                backgroundColor: [
                    "rgba(54, 162, 235, 0.6)",
                    "rgba(255, 99, 132, 0.6)",
                    "rgba(255, 206, 86, 0.6)",
                    "rgba(75, 192, 192, 0.6)",
                    "rgba(153, 102, 255, 0.6)"
                ],
                borderColor: "rgba(54, 162, 235, 1)",
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: chartType === "pie" || chartType === "doughnut" ? {} : {
                y: {
                    beginAtZero: true
                }
            }
        }
    });

    console.log(`Chart generated successfully! Type: ${chartType}`);
}
// Function to TRIM text (removes leading & trailing spaces)
function trimText() {
    if (!lastSelectedCell) {
        alert("Click on a cell first before applying TRIM!");
        return;
    }
    lastSelectedCell.value = lastSelectedCell.value.trim();
    saveData();
}

// Function to convert text to UPPERCASE
function toUpperCaseText() {
    if (!lastSelectedCell) {
        alert("Click on a cell first before applying UPPER!");
        return;
    }
    lastSelectedCell.value = lastSelectedCell.value.toUpperCase();
    saveData();
}

// Function to convert text to lowercase
function toLowerCaseText() {
    if (!lastSelectedCell) {
        alert("Click on a cell first before applying LOWER!");
        return;
    }
    lastSelectedCell.value = lastSelectedCell.value.toLowerCase();
    saveData();
}

// Function to remove duplicate rows in a selected column
function removeDuplicates() {
    let column = document.getElementById("columnSelect").value; // Get selected column
    let seenValues = new Set();

    document.querySelectorAll(`input[data-col='${column}']`).forEach((input, index) => {
        let value = input.value.trim();
        if (value && seenValues.has(value)) {
            input.value = ""; // Clear the duplicate value but keep the row
        } else {
            seenValues.add(value);
        }
    });

    saveData();
}


// Function to find and replace text within a selected column
function findAndReplace() {
    let column = document.getElementById("columnSelect").value; // Get selected column
    let findText = prompt("Enter text to find:");
    if (!findText) return;

    let replaceText = prompt("Enter replacement text:");

    document.querySelectorAll(`input[data-col='${column}']`).forEach(input => {
        if (input.value.includes(findText)) {
            input.value = input.value.replace(new RegExp(findText, "g"), replaceText);
        }
    });

    saveData();
}

// Expose new functions globally
window.trimText = trimText;
window.toUpperCaseText = toUpperCaseText;
window.toLowerCaseText = toLowerCaseText;
window.removeDuplicates = removeDuplicates;
window.findAndReplace = findAndReplace;
window.generateChart = generateChart;
window.exportToExcel = exportToExcel;
window.formatText = formatText;
window.resetFormatting = resetFormatting;
window.clearSheet = clearSheet;
window.addRow = addRow;
window.deleteRow = deleteRow;
window.saveSpreadsheet = saveSpreadsheet;
window.loadSpreadsheet = loadSpreadsheet;

