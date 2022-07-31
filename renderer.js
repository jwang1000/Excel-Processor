/// Section - create callback functions for excel and word files

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
let wordFileDir = "";
let wordFileIsValid = false;
const filePathWordElement = document.getElementById("file-path-word");
wordButton.addEventListener("click", async () => {
    const filePath = await window.electronAPI.openFile();
    filePathWordElement.innerText = getFileName(filePath);
    wordFilePath = filePath;
    wordFileDir = getFileDir(filePath);
    wordFileIsValid = getFileEnding(filePath) === "docx";
});

function getFileName(path) {
    return path.split('\\').pop().split('/').pop();
}

function getFileDir(path) {
    return path.match(/(.*)[\/\\]/)[1] || '';
}

function getFileEnding(path) {
    return path.split('.').pop();
}




/// Section - Create HTML elements for rows

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
    let excelGroupObject = { id: groupId };  // save for later so we can retrieve the inputs
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
        allDOMExcelGroups.removeChild(DOMGroup);
        excelGroupObjects = excelGroupObjects.filter(obj => obj.id != groupId);
    });
    DOMGroup.appendChild(deleteButton);

    allDOMExcelGroups.appendChild(DOMGroup);
    excelGroupObjects.push(excelGroupObject);
}

createExcelGroup();




/// Section - move data from Excel to Word

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
    moveExcelDataToWord();
});

function convertColumnToIndex(val) {
    let base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', i, j, result = 0;

    for (i = 0, j = val.length - 1; i < val.length; i += 1, j -= 1) {
        result += Math.pow(base.length, j) * (base.indexOf(val[i]));
    }

    return result;
};

// for use in saving presets
// list of objects with same data as excelGroupObject, but without the html elements
let presetObjects = [];
let processing = false;

async function moveExcelDataToWord() {
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
    presetObjects = [];
    setStatusNeutral();
    statusText.innerHTML = "Processing...";

    // fill preset objects array
    excelGroupObjects.forEach(obj => {
        let presetObject = {};

        presetObject.id = obj.id;

        // check if column is empty or contains non-alpha characters
        if (!obj.columnInput.value || !/^[a-z]+$/i.test(obj.columnInput.value)) {
            allFieldsFilled = false;
            return;
        } else {
            // save column character for JSON save data
            presetObject.column = obj.columnInput.value.toUpperCase();
            // save column index (0-indexed) for ease of processing excel data
            presetObject.columnIndex = convertColumnToIndex(presetObject.column);
        }

        // validate rows: make sure 0 < startRow <= endRow
        if (!obj.startRowInput.value || !obj.endRowInput.value) {
            allFieldsFilled = false;
            return;
        } else {
            const a = parseInt(obj.startRowInput.value);
            const b = parseInt(obj.endRowInput.value);
            const startRow = Math.min(a, b);
            const endRow = Math.max(a, b);

            if (startRow === 0) {
                allFieldsFilled = false;
                return;
            }

            presetObject.startRow = startRow;
            presetObject.endRow = endRow;
        }

        if (!obj.selectionType.value) {
            allFieldsFilled = false;
            return;
        } else {
            presetObject.useAllCells = obj.selectionType.value === "all";
        }

        if (!obj.numOfRandomCells.value && !presetObject.useAllCells) {
            allFieldsFilled = false;
            return;
        } else {
            presetObject.numOfRandomCells = parseInt(obj.numOfRandomCells.value);
        }

        if (!obj.tagToReplace.value) {
            allFieldsFilled = false;
            return;
        } else {
            presetObject.tagToReplace = obj.tagToReplace.value;
        }

        presetObjects.push(presetObject);
    });

    if (!allFieldsFilled) {
        setStatusError();
        presetObjects = [];
        statusText.innerHTML = "<b>Error</b>: at least one value was not given or is invalid.";
        processing = false;
        return;
    }
    setStatusNeutral();
    statusText.innerHTML = "Processing...";
    console.log(presetObjects);

    // async function to get text from Excel
    const data = await getExcelData();

    const info = await insertTextIntoWord(data);

    if (info === "") {
        statusText.innerHTML = "Success!"
        setStatusSuccess();
    } else {
        statusText.innerHTML = info;
    }
    processing = false;
}

function getExcelData() {
    return new Promise(resolve => {
        let data = [];
        for (let i = 0; i < presetObjects.length; i++) {
            data.push("");
        }
        window.electronAPI.readXlsxFile(excelFilePath).then((allRows) => {
            // `rows` is an array of rows
            // each row being an array of cells.
            allRows.forEach((rowData, rowIndex) => {
                presetObjects.forEach((presetObject, objIndex) => {
                    // rowIndex starts from 0 but Excel rows start from 1
                    if (presetObject.startRow <= rowIndex + 1 && rowIndex + 1 <= presetObject.endRow) {
                        data[objIndex] += rowData[presetObject.columnIndex] + "\n";
                    }
                });
            });
            resolve(data);
        });
    });
}

async function insertTextIntoWord(data) {
    return await window.electronAPI.insertTextIntoWord(data, presetObjects, wordFilePath, wordFileDir);
}
