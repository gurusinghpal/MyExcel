//save
let downloadBtn = document.querySelector(".fa-save");
downloadBtn.addEventListener("click", function () {
    let wb = XLSX.utils.book_new(); // Create new workbook

    sheetsDB.forEach((sheet, index) => {
        let sheetData = sheet.map(row => row.map(cell => cell.value || "")); // Extract cell values
        let ws = XLSX.utils.aoa_to_sheet(sheetData); // Convert to sheet
        XLSX.utils.book_append_sheet(wb, ws, `Sheet${index + 1}`); // Add sheet
    });

    XLSX.writeFile(wb, "MyExcel.xlsx"); // Save file
});

//open
let openBtn = document.querySelector(".fa-external-link-square-alt");
let openInput = document.querySelector(".open_input");

openBtn.addEventListener("click", function () {
    openInput.click();
});

openInput.addEventListener("change", function (e) {
    let file = openInput.files[0];
    if (!file) return;

    let reader = new FileReader();
    reader.readAsBinaryString(file);

    reader.onload = function (e) {
        let data = e.target.result;
        let workbook = XLSX.read(data, { type: "binary" });

        sheetsDB = []; // Clear existing sheets
        sheetList.innerHTML = ""; // Remove previous UI elements

        workbook.SheetNames.forEach((sheetName, index) => {
            let ws = workbook.Sheets[sheetName];
            let sheetData = XLSX.utils.sheet_to_json(ws, { header: 1 });

            let newSheetDB = sheetData.map(row => row.map(value => ({ value })));
            sheetsDB.push(newSheetDB);

            sheetOpenHandler(); // Add new sheet button
        });

        db = sheetsDB[0]; // Set first sheet as active
        setUI();
    };
});

// new
newInput = document.querySelector(".fa-plus-square");
newInput.addEventListener("click", function () {
    var confirmNewSheet = confirm("Opening New Excel Sheet")
    if (confirmNewSheet) {
        sheetsDB = [];
        initDB();    //db reset
        db = sheetsDB[0];
        setUI();
        setSheets();
    }
});

// Open Find & Replace Modal
document.getElementById("openFindReplace").addEventListener("click", function () {
    document.getElementById("findReplaceBox").classList.add("show-find-replace");
});

// Close Modal
document.getElementById("closeFindReplace").addEventListener("click", function () {
    document.getElementById("findReplaceBox").classList.remove("show-find-replace");
    // Reset highlighted cells
    let cells = document.querySelectorAll(".grid .cell");
    cells.forEach(cell => {
        cell.style.backgroundColor = ""; // Remove highlight
    });
});

// Find Functionality (Case-Insensitive)
document.getElementById("findBtn").addEventListener("click", function () {
    let findText = document.getElementById("findInput").value.toLowerCase();
    let cells = document.querySelectorAll(".grid .cell");

    cells.forEach(cell => {
        let cellText = cell.innerText.toLowerCase();
        if (cellText.includes(findText) && findText !== "") {
            cell.style.backgroundColor = "yellow"; // Highlight found cells
        } else {
            cell.style.backgroundColor = ""; // Reset if no match
        }
    });
});

// Replace Functionality (Case-Insensitive)
document.getElementById("replaceBtn").addEventListener("click", function () {
    let findText = document.getElementById("findInput").value.toLowerCase();
    let replaceText = document.getElementById("replaceInput").value;
    let cells = document.querySelectorAll(".grid .cell");

    cells.forEach(cell => {
        let cellText = cell.innerText;
        let cellTextLower = cellText.toLowerCase();

        if (cellTextLower.includes(findText) && findText !== "") {
            // Replace text while maintaining original case for non-matching parts
            let regex = new RegExp(findText, "gi"); // 'g' for global, 'i' for case-insensitive
            cell.innerText = cellText.replace(regex, replaceText);
            cell.style.backgroundColor = "lightgreen"; // Show replaced cells
        }
    });
});


// auto save
let autoSaveEnabled = false;
let autoSaveButton = document.querySelector(".auto-save-toggle");

// Toggle Auto-Save
autoSaveButton.addEventListener("click", function () {
    autoSaveEnabled = !autoSaveEnabled;
    autoSaveButton.textContent = `Auto-Save: ${autoSaveEnabled ? "ON" : "OFF"}`;
});

// Auto-Save every 10 seconds if enabled
setInterval(() => {
    if (autoSaveEnabled) {
        let wb = XLSX.utils.book_new();
        sheetsDB.forEach((sheet, index) => {
            let sheetData = sheet.map(row => row.map(cell => cell.value || ""));
            let ws = XLSX.utils.aoa_to_sheet(sheetData);
            XLSX.utils.book_append_sheet(wb, ws, `Sheet${index + 1}`);
        });

        XLSX.writeFile(wb, "AutoSave_ExcelClone.xlsx");
    }
}, 10000);


// export options
function exportFile(type) {
    if (!db || db.length === 0) {
        alert("No data available to export.");
        return;
    }

    if (type === "pdf") {
        let { jsPDF } = window.jspdf;
        let doc = new jsPDF();

        let sheetData = db.map(row => row.map(cell => cell.value || ""));
        doc.autoTable({ body: sheetData });
        doc.save("MyExcel.pdf");
    } 
    else if (type === "csv") {
        let csvContent = "data:text/csv;charset=utf-8,";
        db.forEach(row => {
            let rowData = row.map(cell => (cell.value || "").replace(/,/g, ""));
            csvContent += rowData.join(",") + "\n";
        });

        let blob = new Blob([csvContent], { type: "text/csv" });
        let a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = "MyExcel.csv";
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    }
}


document.querySelector(".export-pdf").addEventListener("click", function () {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    if (!window.jspdf || !doc.autoTable) {
        alert("Error: jsPDF or autoTable is not available.");
        return;
    }

    let sheetData = db.map(row => row.map(cell => cell.value || ""));
    doc.autoTable({ body: sheetData });
    doc.save("MyExcel.pdf");
});


document.querySelector(".export-csv").addEventListener("click", function () {
    if (!db || db.length === 0) {
        alert("No data available to export.");
        return;
    }

    let csvContent = "data:text/csv;charset=utf-8,";
    db.forEach(row => {
        let rowData = row.map(cell => (cell.value || "").replace(/,/g, "")); // Remove commas
        csvContent += rowData.join(",") + "\n";
    });

    let blob = new Blob([csvContent], { type: "text/csv" });
    let a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "MyExcel.csv";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
});



// spell check

function spellCheck() {
    let cells = document.querySelectorAll('.grid .cell');

    cells.forEach(cell => {
        let text = cell.innerText;
        let words = text.split(/\s+/);

        words.forEach(word => {
            if (!SpellChecker.check(word)) {
                cell.style.border = "2px solid red"; // Highlight misspelled words
            }
        });
    });

    alert("Spell check complete. Misspelled words highlighted.");
};



function setUI() {
    for (let i = 0; i < 100; i++) {
        for (let j = 0; j < 26; j++) {
            let cellObject = db[i][j];
            let tobeChangedCell = document.querySelector(`.grid .cell[rid='${i}'][cid='${j}']`);
            
            // Reset to default styles before applying new ones
            tobeChangedCell.style.fontWeight = "normal";
            tobeChangedCell.style.fontStyle = "normal";
            tobeChangedCell.style.textDecoration = "none";

            // Apply saved styles
            tobeChangedCell.innerText = cellObject.value || "";
            tobeChangedCell.style.color = cellObject.color || "black";
            tobeChangedCell.style.backgroundColor = cellObject.backgroundColor || "white";
            tobeChangedCell.style.fontFamily = cellObject.fontFamily || "Arial";
            tobeChangedCell.style.fontSize = (cellObject.fontSize || 16) + "px";
            tobeChangedCell.style.textAlign = cellObject.halign || "center";

            // Fix underline and italic issue
            tobeChangedCell.style.textDecoration = cellObject.underline ? "underline" : "none";
            tobeChangedCell.style.fontStyle = cellObject.italic ? "italic" : "normal";
            tobeChangedCell.style.fontWeight = cellObject.bold ? "bold" : "normal";
        }
    }
}
