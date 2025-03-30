firstSheet.addEventListener("click", function(){
    for(let i = 0; i<sheetList.children.length;i++){
        sheetList.children[i].classList.remove("active-sheet");
    }
    firstSheet.classList.add("class", "active-sheet")
    db = sheetsDB[0];
    setUI();
})
createSheetIcon.addEventListener("click", function(){
    let childrenNum = sheetList.children.length;
    let newSheet = document.createElement("div");
    newSheet.setAttribute("class", "sheet");
    newSheet.setAttribute("sheetIdx", childrenNum);
    newSheet.textContent = `Sheet ${childrenNum + 1}`;
    sheetList.appendChild(newSheet);
    initDB();
    //active switch
    newSheet.addEventListener("click",function(){
        for(let i = 0; i<childrenNum;i++){
            sheetList.children[i].classList.remove("active-sheet");
        }
        newSheet.classList.add("class", "active-sheet")
        let index = newSheet.getAttribute("sheetIdx");
        db = sheetsDB[index];
        setUI();
    })
})
function sheetOpenHandler(){
    let childrenNum = sheetList.children.length;
    let newSheet = document.createElement("div");
    newSheet.setAttribute("class", "sheet");
    newSheet.setAttribute("sheetIdx", childrenNum);
    newSheet.textContent = `Sheet ${childrenNum + 1}`;
    sheetList.appendChild(newSheet);
    //active switch
    newSheet.addEventListener("click",function(){
        for(let i = 0; i<childrenNum;i++){
            sheetList.children[i].classList.remove("active-sheet");
        }
        newSheet.classList.add("class", "active-sheet")
        let index = newSheet.getAttribute("sheetIdx");
        db = sheetsDB[index];
        setUI();
    })
}

// Add event listeners to magnify cells
function enableCellMagnification() {
    const cells = document.querySelectorAll('.grid .cell');

    cells.forEach((cell) => {
        // On mouseover, apply magnification
        cell.addEventListener('mouseover', function () {
            cell.style.transform = "scale(1.5)"; // Scale up the cell
            cell.style.zIndex = "10"; // Ensure it appears above other elements
            cell.style.boxShadow = "0px 4px 10px rgba(0, 0, 0, 0.2)"; // Add a shadow for effect
            cell.style.transition = "transform 0.2s ease, box-shadow 0.2s ease"; // Smooth animation
            cell.style.fontWeight = "bold"; // Make the text bold

        });

        // On mouseout, remove magnification
        cell.addEventListener('mouseout', function () {
            cell.style.transform = "scale(1)";
            cell.style.zIndex = "1";
            cell.style.boxShadow = "none";
            cell.style.fontWeight = "normal"; // Revert to normal font weight
        });
    });
}

// Call the function when the grid is initialized
enableCellMagnification();

