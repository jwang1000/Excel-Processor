const { app, Menu, BrowserWindow, ipcMain, dialog } = require("electron")
const fs = require("fs");
const path = require("path")

const readXlsxFile = require("read-excel-file/node");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const Store = require("electron-store");
const { SocketAddress } = require("net");

// Save data format:
// store->presetList is a list of strings of all the names of presets
// store->preset.<name> is the list of presetObjects saved under the preset named <name>
const store = new Store();

// presetList is passed to the renderer to fill the dropdown options
let presetList;

/// Loads all presets already saved - presets are loaded on startup
function getPresetList(_event) {
    return presetList;
}

/// For use when loading a specific preset
async function loadPreset(_event, presetName) {
    return store.get("preset")[presetName];
}

/// For saving a preset - updates the preset list as well
async function savePreset(_event, presetName, presetObjects) {
    let presets = store.get("preset");
    presets[presetName] = presetObjects;
    store.set(presets);
}

async function handleFileOpen() {
    const { canceled, filePaths } = await dialog.showOpenDialog();
    if (canceled) {
        return;
    } else {
        return filePaths[0];
    }
}

async function insertTextIntoWord(_event, data, presetObjects, wordFilePath, wordFileDir) {
    try {
        const content = fs.readFileSync(
            wordFilePath,
            "binary"
        );

        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true
        });

        // Render the document (replace tags)
        let renderObj = {};
        presetObjects.forEach((presetObject, objIndex) => {
            renderObj[presetObject.tagToReplace] = data[objIndex];
        });
        doc.render(renderObj);

        const buf = doc.getZip().generate({
            type: "nodebuffer",
            // compression: DEFLATE adds a compression step.
            // For a 50MB output document, expect 500ms additional CPU time
            compression: "DEFLATE",
        });

        // buf is a nodejs Buffer, you can either write it to a
        // file or res.send it with express for example.
        fs.writeFileSync(path.resolve(wordFileDir, "output.docx"), buf);
    } catch (error) {
        return error;
    }
    return "";
}

let aboutWindow = null;
function openAboutWindow() {
    if (aboutWindow) {
        aboutWindow.focus();
        return;
    }

    aboutWindow = new BrowserWindow({
        height: 400,
        resizable: false,
        width: 500,
        title: "About",
        minimizable: false,
        fullscreenable: false,
        autoHideMenuBar: true,
    });

    aboutWindow.loadURL("file://" + __dirname + "/about/about.html");

    aboutWindow.on("closed", () => {
        aboutWindow = null;
    });
}

const isMac = process.platform === "darwin";
const template = [
    // { role: "appMenu" }
    ...(isMac ? [{
        label: app.name,
        submenu: [
            { role: "about" },
            { type: "separator" },
            { role: "services" },
            { type: "separator" },
            { role: "hide" },
            { role: "hideOthers" },
            { role: "unhide" },
            { type: "separator" },
            { role: "quit" }
        ]
    }] : []),
    // { role: "fileMenu" }
    {
        label: "File",
        submenu: [
            isMac ? { role: "close" } : { role: "quit" }
        ]
    },
    // { role: "editMenu" }
    {
        label: "Edit",
        submenu: [
            { role: "undo" },
            { role: "redo" },
            { type: "separator" },
            { role: "cut" },
            { role: "copy" },
            { role: "paste" },
            ...(isMac ? [
                { role: "pasteAndMatchStyle" },
                { role: "delete" },
                { role: "selectAll" },
                { type: "separator" },
                {
                    label: "Speech",
                    submenu: [
                        { role: "startSpeaking" },
                        { role: "stopSpeaking" }
                    ]
                }
            ] : [
                { role: "delete" },
                { type: "separator" },
                { role: "selectAll" }
            ])
        ]
    },
    // { role: "viewMenu" }
    {
        label: "View",
        submenu: [
            { role: "reload" },
            { role: "forceReload" },
            { role: "toggleDevTools" },
            { type: "separator" },
            { role: "resetZoom" },
            { role: "zoomIn" },
            { role: "zoomOut" },
            { type: "separator" },
            { role: "togglefullscreen" }
        ]
    },
    // { role: "windowMenu" }
    {
        label: "Window",
        submenu: [
            { role: "minimize" },
            { role: "zoom" },
            ...(isMac ? [
                { type: "separator" },
                { role: "front" },
                { type: "separator" },
                { role: "window" }
            ] : [
                { role: "close" }
            ])
        ]
    },
    {
        role: "help",
        submenu: [
            {
                label: "About",
                click: openAboutWindow
            },
            {
                label: "Contact",
                click: async () => {
                    const { shell } = require("electron")
                    await shell.openExternal("https://www.jwang1000.com/contact")
                }
            },
            {
                label: "See GitHub repo...",
                click: async () => {
                    const { shell } = require("electron")
                    await shell.openExternal("https://github.com/jwang1000/Excel-Processor")
                }
            }
        ]
    }
];

const menu = Menu.buildFromTemplate(template);
Menu.setApplicationMenu(menu);

const createWindow = () => {
    const win = new BrowserWindow({
        minWidth: 900,
        minHeight: 500,
        title: "Excel to Word",
        show: false,
        icon: "excel-word-logo.png",
        webPreferences: {
            preload: path.join(__dirname, "preload.js")
        }
    });
    win.maximize();
    win.loadFile("index.html");
    win.show();
}

app.whenReady().then(() => {
    ipcMain.handle("dialog:openFile", handleFileOpen);
    ipcMain.handle("readXlsxFile", (_event, filePath) => {
        return readXlsxFile(filePath);
    });
    ipcMain.handle("insertTextIntoWord", insertTextIntoWord);
    ipcMain.handle("getPresetList", getPresetList);
    ipcMain.handle("loadPreset", loadPreset);
    ipcMain.handle("savePreset", savePreset);

    // load list of presets
    presetList = store.get("presetList");

    createWindow();

    // On macOS it"s common to re-create a window in the app when the
    // dock icon is clicked and there are no other windows open.
    app.on("activate", () => {
        if (BrowserWindow.getAllWindows().length === 0) createWindow();
    });
})

// Quit when all windows are closed, except on macOS. There, it"s common
// for applications and their menu bar to stay active until the user quits
// explicitly with Cmd + Q.
app.on("window-all-closed", () => {
    if (process.platform !== "darwin") app.quit();
})
