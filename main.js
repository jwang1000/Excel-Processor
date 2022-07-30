const { app, BrowserWindow, ipcMain, dialog } = require('electron')
const fs = require("fs");
const path = require('path')

const readXlsxFile = require('read-excel-file/node');
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

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

const createWindow = () => {
    const win = new BrowserWindow({
        minWidth: 900,
        minHeight: 500,
        title: "Excel to Word",
        show: false,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js')
        }
    });
    win.maximize();
    win.loadFile('index.html');
    win.show();
}

app.whenReady().then(() => {
    ipcMain.handle('dialog:openFile', handleFileOpen);
    ipcMain.handle('readXlsxFile', (_event, filePath) => {
        return readXlsxFile(filePath);
    });
    ipcMain.handle('insertTextIntoWord', insertTextIntoWord);

    createWindow();

    // On macOS it's common to re-create a window in the app when the
    // dock icon is clicked and there are no other windows open.
    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) createWindow();
    });
})

// Quit when all windows are closed, except on macOS. There, it's common
// for applications and their menu bar to stay active until the user quits
// explicitly with Cmd + Q.
app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
})
