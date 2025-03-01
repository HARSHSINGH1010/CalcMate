let lastSelectedCell = null; 
let isDragging = false; 
let startCell = null; 
let selectedCells = []; 
let cellDependencies = {}; 
document.addEventListener("DOMContentLoaded", function () {
    const tbody = document.querySelector("#spreadsheet tbody");

    
    for (let i = 1; i <= 10; i++) {
        let row = document.createElement("tr");

       
        let th = document.createElement("th");
        th.innerText = i;
        row.appendChild(th);

       
        for (let j = 0; j < 4; j++) {
            let cell = document.createElement("td");
            let input = document.createElement("input");
            input.setAttribute("data-row", i);
            input.setAttribute("data-col", String.fromCharCode(65 + j)); // A, B, C, D

           
            input.addEventListener("mousedown", function () {
                lastSelectedCell = input;
                isDragging = true;
                startCell = input;
                selectedCells = [input];
            });

            input.addEventListener("mouseover", function () {
                if (isDragging) {
                    selectedCells.push(input);
                    input.style.backgroundColor = "#e8f0fe"; 
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


function applyDragContent() {
    if (!startCell || selectedCells.length < 2) return;

    let value = startCell.value; 

    selectedCells.forEach(cell => {
        if (value.startsWith("=")) {
            
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


function handleFormula(event) {
    let input = event.target;
    let value = input.value.trim();
    let cellKey = input.getAttribute("data-col") + input.getAttribute("data-row");

    if (value.startsWith("=")) {
        try {
            let formula = value.substring(1); 
            let result = evaluateFormula(formula, cellKey);
            input.value = result;

            
            input.setAttribute("data-formula", formula);

            
            trackDependencies(cellKey, formula);
        } catch (error) {
            input.value = "Error";
        }
    }
    saveData();
}



function trackDependencies(cellKey, formula) {
    cellDependencies[cellKey] = [];

    let cellRefs = formula.match(/[A-D]\d+/g); 
    if (cellRefs) {
        cellDependencies[cellKey] = cellRefs;
    }
}


function updateDependentCells(changedCellKey) {
    Object.keys(cellDependencies).forEach(dependentCell => {
        if (cellDependencies[dependentCell].includes(changedCellKey)) {
            let input = document.querySelector(`input[data-col="${dependentCell.charAt(0)}"][data-row="${dependentCell.slice(1)}"]`);
            if (input) {
                let formula = "=" + input.dataset.formula; 
                input.value = evaluateFormula(formula);
            }
        }
    });
}


document.querySelectorAll("input").forEach(input => {
    input.addEventListener("input", function () {
        let cellKey = input.getAttribute("data-col") + input.getAttribute("data-row");
        updateDependentCells(cellKey);
    });
});



function evaluateFormula(formula, currentCell = "") {
    try {
        
        let match;
        if ((match = formula.match(/SUM\((\w)(\d+):(\w)(\d+)\)/))) return calculateRange(match, "SUM");
        if ((match = formula.match(/AVERAGE\((\w)(\d+):(\w)(\d+)\)/))) return calculateRange(match, "AVERAGE");
        if ((match = formula.match(/MAX\((\w)(\d+):(\w)(\d+)\)/))) return calculateRange(match, "MAX");
        if ((match = formula.match(/MIN\((\w)(\d+):(\w)(\d+)\)/))) return calculateRange(match, "MIN");
        if ((match = formula.match(/COUNT\((\w)(\d+):(\w)(\d+)\)/))) return calculateRange(match, "COUNT");

        
        formula = formula.replace(/\$?([A-D])\$?(\d+)/g, (match, col, row) => {
            let cell = document.querySelector(`input[data-row="${row}"][data-col="${col}"]`);
            return cell && !isNaN(cell.value) ? parseFloat(cell.value) : 0;
        });

        return eval(formula); 
    } catch (error) {
        return "Error";
    }
}



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


function saveData() {
    let data = {};
    document.querySelectorAll("input").forEach(input => {
        let row = input.getAttribute("data-row");
        let col = input.getAttribute("data-col");
        let key = `${col}${row}`;
        data[key] = {
            value: input.value 
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
            input.value = data[key].value || ""; 
        }
    });
}




function clearSheet() {
    document.querySelectorAll("input").forEach(input => {
        input.value = "";
    });
    localStorage.removeItem("spreadsheetData");
    location.reload();
}

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
    let column = document.getElementById("columnSelect").value; 
    let chartType = document.getElementById("chartTypeSelect").value; 
    let labels = [];
    let values = [];

    
    document.querySelectorAll(`input[data-col='${column}']`).forEach(input => {
        let row = input.getAttribute("data-row");
        let value = parseFloat(input.value.trim()); 

        if (!isNaN(value) && input.value.trim() !== "") { 
            labels.push("Row " + row);
            values.push(value);
        }
    });

    
    console.log("Chart Labels:", labels);
    console.log("Chart Values:", values);

    
    if (values.length === 0) {
        alert("No valid numerical data found in the selected column!");
        return;
    }

    
    let canvas = document.getElementById("myChart");
    if (!canvas) {
        alert("Error: Chart canvas not found!");
        return;
    }

    let ctx = canvas.getContext("2d");

    
    if (window.myChart instanceof Chart) {
        window.myChart.destroy();
    }

    
    window.myChart = new Chart(ctx, {
        type: chartType, 
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

function trimText() {
    if (!lastSelectedCell) {
        alert("Click on a cell first before applying TRIM!");
        return;
    }
    lastSelectedCell.value = lastSelectedCell.value.trim();
    saveData();
}


function toUpperCaseText() {
    if (!lastSelectedCell) {
        alert("Click on a cell first before applying UPPER!");
        return;
    }
    lastSelectedCell.value = lastSelectedCell.value.toUpperCase();
    saveData();
}


function toLowerCaseText() {
    if (!lastSelectedCell) {
        alert("Click on a cell first before applying LOWER!");
        return;
    }
    lastSelectedCell.value = lastSelectedCell.value.toLowerCase();
    saveData();
}


function removeDuplicates() {
    let column = document.getElementById("columnSelect").value; 
    let seenValues = new Set();

    document.querySelectorAll(`input[data-col='${column}']`).forEach((input, index) => {
        let value = input.value.trim();
        if (value && seenValues.has(value)) {
            input.value = ""; 
        } else {
            seenValues.add(value);
        }
    });

    saveData();
}



function findAndReplace() {
    let column = document.getElementById("columnSelect").value; 
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

