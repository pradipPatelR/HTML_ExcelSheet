let sheetData = []; // Store Excel data
let stateId = -1;
let districtId = -1;

// Load Excel File
document.getElementById("uploadExcel").addEventListener("change", function(event) {
    let reader = new FileReader();
    reader.readAsBinaryString(event.target.files[0]);
    reader.onload = function(event) {
        let data = event.target.result;
        let workbook = XLSX.read(data, { type: "binary" });
        let sheetName = workbook.SheetNames[0];
        let sheet = workbook.Sheets[sheetName];
        sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        renderTable(sheetData);
    };
});

// Change Search File Type

document.getElementById("search").addEventListener("change", function(event) {
    
    document.getElementById("searchInput").value = "";
    searchTable();
});


// Render Data into Table
function renderTable(data) {
    let table = document.getElementById("excelTable");
    table.style.borderCollapse = 'collapse';
    table.style.width = '100%';
    table.innerHTML = ""; // Clear table
    
    let search = document.getElementById("search");
    search.innerHTML = ""; // Clear search
    
    let selectOption = document.createElement("option");
    selectOption.value = "Select";
    selectOption.innerText = "Select";
    selectOption.id = "";
    
    search.appendChild(selectOption);
    
    data.forEach((row, rowIndex) => {
        
        let tr = document.createElement("tr");
        row.forEach((cell, colIndex) => {
            
            if (rowIndex === 0) {
                let newOption = document.createElement("option");
                newOption.value = cell;
                newOption.innerText = cell;
                newOption.id = "" + colIndex;
                search.appendChild(newOption);
            }
            
            let td = document.createElement(rowIndex === 0 ? "th" : "td");
            td.innerText = cell;
            if (rowIndex !== 0) {
                td.contentEditable = true; // Allow editing
                td.oninput = function() {
                    sheetData[rowIndex][colIndex] = this.innerText; // Update array
                };

                // new else design part header row data
                td.style.border = '1px solid black';
                td.style.padding = '8px';

            } else {
                // new else design part other row data
                td.style.border = '1px solid black';
                td.style.backgroundColor = '#f2f2f2';
                td.style.padding = '8px';
                td.style.textAlign = 'center';
            }
            tr.appendChild(td);
        });
        
        tr.id = "row_" + rowIndex; // Assign unique ID
        
        if ((rowIndex !== 0) && (row.length > 0)) {
            let deleteButton = document.createElement("button");
            deleteButton.id = "deleteBtn_" + rowIndex; // Assign unique ID
            deleteButton.innerText = "Delete";
            deleteButton.style.backgroundColor = "red";
            deleteButton.style.color = "white";
            deleteButton.style.border = "none";
            deleteButton.style.padding = "10px";
            deleteButton.style.cursor = "pointer";

            deleteButton.addEventListener("click", function() {
                
                let row = document.getElementById("row_" + rowIndex);
                
                if (row) {
                    row.remove();  // Remove row if it exists
                }
            });
            
            tr.appendChild(deleteButton);
        }
        
        table.appendChild(tr);
    });
}

// Search Function
function searchTable() {
    let input = document.getElementById("searchInput").value.toLowerCase();
    let table = document.getElementById("excelTable");
    let rows = table.getElementsByTagName("tr");
    
    let rowIdx = document.getElementById('search').selectedIndex
    
    if (rowIdx == 0) {
        rowIdx = -1
    } else {
        rowIdx = rowIdx - 1
    }
    
    for (let i = 1; i < rows.length; i++) {
        let row = rows[i].getElementsByTagName("td");
        let found = false;
        
        if (rowIdx == -1) {
            for (let j = 0; j < row.length; j++) {
                if (row[j].innerText.toLowerCase().includes(input)) {
                    found = true;
                    break;
                }
            }
        } else {
            if (row[rowIdx].innerText.toLowerCase().includes(input)) {
                found = true;
            }
        }
        
        rows[i].style.display = found ? "" : "none";
        
        let idStr = "deleteBtn_" + i;
        document.getElementById(idStr).style.display = found ? "block" : "none";
        
    }
}

// Save Updated Excel File
function saveExcel() {
    let ws = XLSX.utils.aoa_to_sheet(sheetData);
    let wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "UpdatedSheet");
    XLSX.writeFile(wb, "UpdatedExcel.xlsx");
}
