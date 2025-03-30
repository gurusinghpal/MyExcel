// **************************MENU ELEMENTS***************************
let fontSizeInput = document.querySelector(".font_size_input");
let fontFamilyInput = document.querySelector(".font_family_input");
let boldInput = document.querySelector(".fa-bold");
let ItalicInput = document.querySelector(".fa-italic");
let UnderlineInput = document.querySelector(".fa-underline");
let alignInput = document.querySelector(".alignment_container");
let backgroundHInput = document.querySelector(".background_color");
let backgroundInput = document.querySelector(".fa-fill-drip");
let textColorHInput = document.querySelector(".text_color");
let textColorInput = document.querySelector(".fa-font");
let createSheetIcon = document.querySelector(".newSheet");
let sheetList = document.querySelector(".sheets-list");
let firstSheet = document.querySelector(".sheet");

//*******************************************************************

// Creating cells (A-Z)
let topRow = document.querySelector(".top_row");
for (let i = 0; i < 26; i++) {
    let div = document.createElement("div");
    div.setAttribute("class", "cell");
    div.textContent = String.fromCharCode(65 + i); // ASCII for A-Z
    topRow.appendChild(div);
}

// Creating Cells (1-100)
let leftCol = document.querySelector(".left_col");
for (let i = 1; i <= 100; i++) {
    let div = document.createElement("div");
    div.setAttribute("class", "cell");
    div.textContent = i;
    leftCol.appendChild(div);
}

// Creating empty cells
let grid = document.querySelector(".grid");
for (let i = 0; i < 100; i++) {
    let row = document.createElement("div");
    row.setAttribute("class", "row");
    for (let j = 0; j < 26; j++) {
        let div = document.createElement("div");
        div.setAttribute("class", "cell");
        div.setAttribute("contentEditable", "true");
        div.setAttribute("rId", i);
        div.setAttribute("cId", j);
        row.appendChild(div);
    }
    grid.appendChild(row);
}

// Creating Database of cell's values
let sheetsDB = [];
initDB();
let db = sheetsDB[0];

// Clicking on cell will give address using rID, cID to address Bar
let allCells = document.querySelectorAll(".grid .cell");
let addressInput = document.querySelector(".address_input");
let formulaInput = document.querySelector(".formula_input");

for (let i = 0; i < allCells.length; i++) {
    allCells[i].addEventListener("click", function (e) {
        let r = allCells[i].getAttribute("rId");
        let c = allCells[i].getAttribute("cId");
        r = Number(r);
        c = Number(c);
        addressInput.value = String.fromCharCode(c + 65) + (r + 1);

        // *********** TWO WAY BINDING ***********
        let cellObject = db[r][c];
        fontSizeInput.value = cellObject.fontSize;
        fontFamilyInput.value = cellObject.fontFamily;
        boldInput.classList.toggle("selected", cellObject.bold);
        ItalicInput.classList.toggle("selected", cellObject.italic);
        UnderlineInput.classList.toggle("selected", cellObject.underline);

        let options = alignInput.children;
        for (let i = 0; i < options.length; i++) {
            options[i].classList.remove("selected");
        }
        if (cellObject.halign) {
            for (let i = 0; i < options.length; i++) {
                let elementClasses = options[i].classList;
                let hAlignment = elementClasses[elementClasses.length - 1];
                if (hAlignment == cellObject.halign) {
                    elementClasses.add("selected");
                }
            }
        }

        textColorInput.style.color = cellObject.color || "black";
        backgroundInput.style.color = cellObject.backgroundColor || "black";
        formulaInput.value = cellObject.formula;
    });
}

// Get First Cell
let firstCell = document.querySelector(".grid .cell[rid='0'][cid='0']");
firstCell.focus();
firstCell.click();

// Get RID CID from address
function getRidCid(address) {
    let ASCI = address.charCodeAt(0);
    let cid = ASCI - 65;
    let rid = Number(address.substring(1)) - 1;
    return { rId: rid, cId: cid };
}

// Initializing Database
function initDB() {
    let db = [];
    for (let i = 0; i < 100; i++) {
        let rowArr = [];
        for (let j = 0; j < 26; j++) {
            let cellObject = {
                color: "",
                backgroundColor: "",
                fontFamily: "Arial",
                fontSize: 16,
                halign: "center",
                underline: false,
                bold: false,
                italic: false,
                value: "",
                formula: "",
                children: []
            };
            rowArr.push(cellObject);
        }
        db.push(rowArr);
    }
    sheetsDB.push(db);
}

document.addEventListener("DOMContentLoaded", function () {
    const menuBtn = document.querySelector(".menu-btn");
    const menuOptions = document.querySelector(".menu-options");
    const dropdownBtn = document.getElementById("dropdownBtn");
    const dropdownContent = document.getElementById("dropdownContent");

    // Toggle Floating Menu
    menuBtn.addEventListener("click", (event) => {
        menuOptions.style.display = menuOptions.style.display === "flex" ? "none" : "flex";
        event.stopPropagation(); // Prevents immediate closure
    });

    // Toggle Dropdown on Click
    dropdownBtn.addEventListener("click", (event) => {
        dropdownContent.classList.toggle("show");
        event.stopPropagation(); // Prevents immediate closure
    });

    // Close Floating Menu & Dropdown when clicking outside
    document.addEventListener("click", function (event) {
        if (!menuBtn.contains(event.target) && !menuOptions.contains(event.target)) {
            menuOptions.style.display = "none";
        }
        if (!dropdownBtn.contains(event.target) && !dropdownContent.contains(event.target)) {
            dropdownContent.classList.remove("show");
        }
    });

    // Apply Color Blind Mode
    function applyColorBlindMode(mode) {
        const colorClasses = [
            "color-blind-protanopia",
            "color-blind-deuteranopia",
            "color-blind-tritanopia",
            "color-blind-achromatopsia",
            "color-blind-achromatomaly",
            "high-contrast-mode"
        ];
        document.body.classList.remove(...colorClasses);
        if (mode) {
            document.body.classList.add(mode);
        }
    }

    // Event Listener for all contrast options
    document.querySelectorAll(".contrast-option").forEach(button => {
        button.addEventListener("click", function () {
            let mode = this.getAttribute("data-mode");
            applyColorBlindMode(mode);
        });
    });
});


// // // Text to Speech (Read full grid text)
// document.getElementById("textToSpeech").addEventListener("click", function () {
//     let text = document.querySelector(".grid").innerText;
//     let speech = new SpeechSynthesisUtterance(text);
//     speechSynthesis.speak(speech);
// });

// Voice Command Activation
document.getElementById("voiceCommand").addEventListener("click", function () {
    let recognition = new (window.SpeechRecognition || window.webkitSpeechRecognition)();
    recognition.start();
    recognition.onresult = function (event) {
        let command = event.results[0][0].transcript.toLowerCase();
        alert("You said: " + command);
    };
});

// Hover Effect for Cells (Zoom & Shadow)
document.querySelectorAll(".grid .cell").forEach(cell => {
    cell.addEventListener("mouseover", function () {
        this.style.transform = "scale(1.5)";
        this.style.zIndex = "10";
        this.style.boxShadow = "0 0 5px rgba(0, 0, 0, 0.3)";
    });

    cell.addEventListener("mouseout", function () {
        this.style.transform = "scale(1)";
        this.style.zIndex = "1";
        this.style.boxShadow = "none";
    });
});

//text to speech
document.addEventListener("DOMContentLoaded", function () {
    let synth = window.speechSynthesis;
    let hoverTimeout;
    let lastSpokenText = "";
    let ttsEnabled = false; // Flag to track TTS state

    const toggleTTSButton = document.getElementById("toggleTTS");

    // Toggle TTS state
    toggleTTSButton.addEventListener("click", function () {
        ttsEnabled = !ttsEnabled;
        if (!ttsEnabled) {
            synth.cancel(); // Stop ongoing speech when turning off
        }
        this.textContent = ttsEnabled ? "Disable Text-to-Speech" : "Enable Text-to-Speech";
    });

    function speakText(text) {
        if (!synth || synth.speaking) {
            synth.cancel(); // Stop any ongoing speech
        }
        if (text.trim() !== "" && text !== lastSpokenText) {
            let speech = new SpeechSynthesisUtterance(text);
            speech.lang = "en-US";
            synth.speak(speech);
            lastSpokenText = text;
        }
    }

    document.body.addEventListener("mouseover", function (event) {
        if (!ttsEnabled) return; // Only speak if TTS is enabled

        let target = event.target; // Get the hovered element
        let text = target.innerText || target.getAttribute("alt") || target.getAttribute("aria-label") || "";

        text = text.trim();

        if (text !== "") {
            hoverTimeout = setTimeout(() => {
                speakText(text);
            }, 200); // Small delay to prevent rapid speech triggers
        }
    });

    document.body.addEventListener("mouseout", function () {
        clearTimeout(hoverTimeout);
        if (!ttsEnabled) {
            synth.cancel(); // Stop speech immediately when moving away if TTS is disabled
        }
    });

    // // Apply Color Blind or High Contrast Mode
    // function applyMode(mode) {
    //     document.body.className = ""; // Reset previous modes

    //     if (mode === "high-contrast") {
    //         document.body.classList.add("high-contrast-mode");
    //     } else if (mode) {
    //         document.body.classList.add(`color-blind-${mode}`);
    //     }
    // }

    // // Event Listener for all contrast options
    // document.querySelectorAll(".contrast-option").forEach(button => {
    //     button.addEventListener("click", function () {
    //         let mode = this.getAttribute("data-mode");
    //         applyMode(mode);
    //     });
    // });

});



