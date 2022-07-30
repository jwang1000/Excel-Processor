// All of the Node.js APIs are available in the preload process.
// It has the same sandbox as a Chrome extension.
const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('versions', {
    node: () => process.versions.node,
    chrome: () => process.versions.chrome,
    electron: () => process.versions.electron
})

contextBridge.exposeInMainWorld('electronAPI', {
    openFile: () => ipcRenderer.invoke('dialog:openFile'),
    readXlsxFile: (filePath) => ipcRenderer.invoke('readXlsxFile', filePath),
    insertTextIntoWord: (data,
        presetObjects,
        wordFilePath,
        wordFileDir) => ipcRenderer.invoke('insertTextIntoWord',
            data,
            presetObjects,
            wordFilePath,
            wordFileDir)
})
