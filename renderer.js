const information = document.getElementById("info");
information.innerText = `This app is using Chrome (v${versions.chrome()}), Node.js (v${versions.node()}), and Electron (v${versions.electron()})`;

const func = async () => {
    const response = await window.versions.ping();
    console.log(response); // prints out "pong"
}

func();

const btn = document.getElementById("btn")
const filePathElement = document.getElementById("filePath")

btn.addEventListener("click", async () => {
    const filePath = await window.electronAPI.openFile();
    filePathElement.innerText = filePath;
})

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

    const selectionType = document.createElement("select");
    const optionAll = document.createElement("option");
    optionAll.value = "all";
    optionAll.text = "All Cells";
    selectionType.appendChild(optionAll);
    const optionRandom = document.createElement("option");
    optionRandom.value = "random";
    optionRandom.text = "Randomly Pick";
    selectionType.appendChild(optionRandom);
    excelGroupObject.selectionType = selectionType;
    DOMGroup.appendChild(selectionType);

    // TODO create element that is hidden unless "random" is selected
    const numOfRandomCells = document.createElement("input");
    numOfRandomCells.type = "number";
    numOfRandomCells.min = "1";
    excelGroupObject.numOfRandomCells = numOfRandomCells;
    DOMGroup.appendChild(numOfRandomCells);

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

runButton.addEventListener("click", () => {
    insertTextIntoWord();
});

function insertTextIntoWord() {
    // TODO
    // for now, just checking data
    excelGroupObjects.forEach(obj => {
        console.log(obj.id);
        console.log(obj.columnInput);
        console.log(obj.columnInput.value);
        console.log(obj.startRowInput.value);
        console.log(obj.endRowInput.value);
    })
}
