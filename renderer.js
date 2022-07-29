const information = document.getElementById("info");
information.innerText = `This app is using Chrome (v${versions.chrome()}), Node.js (v${versions.node()}), and Electron (v${versions.electron()})`;

const func = async () => {
    const response = await window.versions.ping();
    console.log(response); // prints out "pong"
}

func();

const excelButton = document.getElementById("excel-file-button");
let excelFilePath = "";
let excelFileIsValid = false;
const filePathExcelElement = document.getElementById("file-path-excel");
excelButton.addEventListener("click", async () => {
    const filePath = await window.electronAPI.openFile();
    filePathExcelElement.innerText = getFileName(filePath);
    excelFilePath = filePath;
    excelFileIsValid = getFileEnding(filePath) === "xlsx";
});

const wordButton = document.getElementById("word-file-button");
let wordFilePath = "";
let wordFileIsValid = false;
const filePathWordElement = document.getElementById("file-path-word");
wordButton.addEventListener("click", async () => {
    const filePath = await window.electronAPI.openFile();
    filePathWordElement.innerText = getFileName(filePath);
    wordFilePath = filePath;
    wordFileIsValid = getFileEnding(filePath) === "docx";
});

function getFileName(path) {
    return path.split('\\').pop().split('/').pop();
}

function getFileEnding(path) {
    return path.split('.').pop();
}

const allDOMExcelGroups = document.getElementById("all-excel-groups");
const addButton = document.getElementById("add-excel-group");
const runButton = document.getElementById("run");

let maxGroupId = 0;  // keep a unique identifier for each group for ease of deletion
let excelGroupObjects = [];

addButton.addEventListener("click", () => {
    createExcelGroup();
});

function createExcelGroup() {
    const DOMGroup = document.createElement("div");
    DOMGroup.classList.add("excelGroup");
    const groupId = maxGroupId;
    let excelGroupObject = {id: groupId};  // save for later so we can retrieve the inputs
    maxGroupId += 1;

    const columnInput = document.createElement("input");
    excelGroupObject.columnInput = columnInput;
    DOMGroup.appendChild(columnInput);

    const startRowInput = document.createElement("input");
    startRowInput.type = "number";
    startRowInput.min = "1";
    excelGroupObject.startRowInput = startRowInput;
    DOMGroup.appendChild(startRowInput);

    const endRowInput = document.createElement("input");
    endRowInput.type = "number";
    endRowInput.min = "1";
    excelGroupObject.endRowInput = endRowInput;
    DOMGroup.appendChild(endRowInput);

    // create div for dropdown and optional input field below
    const dropdownDiv = document.createElement("div");
    dropdownDiv.style.display = "flex";
    dropdownDiv.style.flexDirection = "column";

    const selectionType = document.createElement("select");
    const optionAll = document.createElement("option");
    optionAll.value = "all";
    optionAll.text = "All Cells";
    selectionType.appendChild(optionAll);
    const optionRandom = document.createElement("option");
    optionRandom.value = "random";
    optionRandom.text = "Randomly Pick";
    selectionType.appendChild(optionRandom);

    const numOfRandomCellsLabel = document.createElement("p");
    numOfRandomCellsLabel.innerHTML = "Pick how many cells?";
    numOfRandomCellsLabel.style.fontSize = "12px";
    numOfRandomCellsLabel.style.display = "none";

    const numOfRandomCells = document.createElement("input");
    numOfRandomCells.type = "number";
    numOfRandomCells.min = "1";
    numOfRandomCells.style.display = "none";
    numOfRandomCells.style.width = "7.5rem";
    // make input invisible unless the randomly pick option is selected
    selectionType.addEventListener("click", () => {
        // sigh... this checks the previous value instead of the new one
        // TODO find a better solution
        if (selectionType.value !== "random") {
            numOfRandomCells.style.display = "none";
            numOfRandomCellsLabel.style.display = "none";
        }
        else {
            numOfRandomCells.style.display = "block";
            numOfRandomCellsLabel.style.display = "block";
        }
    });
    
    excelGroupObject.selectionType = selectionType;
    excelGroupObject.numOfRandomCells = numOfRandomCells;
    dropdownDiv.appendChild(selectionType);
    dropdownDiv.appendChild(numOfRandomCellsLabel);
    dropdownDiv.appendChild(numOfRandomCells);
    DOMGroup.appendChild(dropdownDiv);

    const tagToReplace = document.createElement("input");
    tagToReplace.value = "replace" + groupId;  // default tag
    excelGroupObject.tagToReplace = tagToReplace;
    DOMGroup.appendChild(tagToReplace);

    const deleteButton = document.createElement("button");
    deleteButton.textContent = "Delete";
    deleteButton.addEventListener("click", () => {
        console.log("delete");
        allDOMExcelGroups.removeChild(DOMGroup);
        excelGroupObjects = excelGroupObjects.filter(obj => obj.id != groupId);
    });
    DOMGroup.appendChild(deleteButton);

    allDOMExcelGroups.appendChild(DOMGroup);
    excelGroupObjects.push(excelGroupObject);
}



createExcelGroup();

const statusText = document.getElementById("status");

function setStatusError() {
    statusText.style.color = "red";
}

function setStatusNeutral() {
    statusText.style.color = "black";
}

function setStatusSuccess() {
    statusText.style.color = "green";
}

runButton.addEventListener("click", () => {
    insertTextIntoWord();
});

// for use in saving presets
// list of objects with same data as excelGroupObject, but without the html elements
let presetObjects = [];
let processing = false;

// TODO
// for now, just checking data
function insertTextIntoWord() {
    if (processing) return;  // don't do anything if process has already started

    if (!excelFileIsValid) {
        setStatusError();
        statusText.innerHTML = "<b>Error</b>: there is no valid Excel file.";
        return;
    } else if (!wordFileIsValid) {
        setStatusError();
        statusText.innerHTML = "<b>Error</b>: there is no valid Word file.";
        return;
    }

    let allFieldsFilled = true;
    processing = true;
    setStatusNeutral();
    statusText.innerHTML = "Processing...";

    excelGroupObjects.forEach(obj => {
        let presetObject = {};

        console.log(obj.id);
        presetObject.id = obj.id;

        console.log(obj.columnInput);
        console.log(obj.columnInput.value);
        if (!obj.columnInput.value) {
            allFieldsFilled = false;
            return;
        } else {
            presetObject.column = obj.columnInput.value;
        }

        console.log(obj.startRowInput.value);
        if (!obj.startRowInput.value) {
            allFieldsFilled = false;
            return;
        } else {
            presetObject.startRow = parseInt(obj.startRowInput.value);
        }

        console.log(obj.endRowInput.value);
        if (!obj.endRowInput.value) {
            allFieldsFilled = false;
            return;
        } else {
            presetObject.endRow = parseInt(obj.endRowInput.value);
        }

        console.log(obj.selectionType.value);
        if (!obj.selectionType.value) {
            allFieldsFilled = false;
            return;
        } else {
            presetObject.useAllCells = obj.selectionType.value === "all";
        }

        console.log(obj.numOfRandomCells.value);
        if (!obj.numOfRandomCells.value && !presetObject.useAllCells) {
            allFieldsFilled = false;
            return;
        } else {
            presetObject.numOfRandomCells = parseInt(obj.numOfRandomCells.value);
        }

        console.log(obj.tagToReplace.value);
        if (!obj.tagToReplace.value) {
            allFieldsFilled = false;
            return;
        } else {
            presetObject.tagToReplace = obj.tagToReplace.value;
        }

        presetObjects.push(presetObject);
        console.log(presetObject);
    });
    
    if (!allFieldsFilled) {
        setStatusError();
        presetObjects = [];
        statusText.innerHTML = "<b>Error</b>: at least one value was not given.";
        processing = false;
        return;
    }
    setStatusNeutral();
    statusText.innerHTML = "Processing...";
    console.log(presetObjects);

    // TODO create async function to get text from Excel

    processing = false;
}
