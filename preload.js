// All of the Node.js APIs are available in the preload process.
// It has the same sandbox as a Chrome extension.
const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("electronAPI", {
    openFile: () => ipcRenderer.invoke("dialog:openFile"),
    readXlsxFile: (filePath) => ipcRenderer.invoke("readXlsxFile", filePath),
    insertTextIntoWord: (data,
        presetObjects,
        wordFilePath,
        wordFileDir,
        outputFile) => ipcRenderer.invoke("insertTextIntoWord",
            data,
            presetObjects,
            wordFilePath,
            wordFileDir,
            outputFile),
    getPresetList: () => ipcRenderer.invoke("getPresetList"),
    loadPreset: (presetName) => ipcRenderer.invoke("loadPreset", presetName),
    savePreset: (presetName, presetObjects) => ipcRenderer.invoke("savePreset", presetName, presetObjects),
    deletePreset: (presetName) => ipcRenderer.invoke("deletePreset", presetName)
});
