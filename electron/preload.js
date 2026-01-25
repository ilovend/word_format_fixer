const { contextBridge, ipcRenderer } = require('electron');

// 向渲染进程暴露IPC方法
contextBridge.exposeInMainWorld('electronAPI', {
    processDocument: (documentPath, activeRules) => ipcRenderer.invoke('process-document', documentPath, activeRules),
    getRules: () => ipcRenderer.invoke('get-rules'),
    getPresets: () => ipcRenderer.invoke('get-presets'),
    configureRules: (configs) => ipcRenderer.invoke('configure-rules', configs),
    selectFile: () => ipcRenderer.invoke('select-file'),
    savePreset: (presetId, presetData) => ipcRenderer.invoke('save-preset', presetId, presetData),
    deletePreset: (presetId) => ipcRenderer.invoke('delete-preset', presetId),
    importConfigFile: () => ipcRenderer.invoke('importConfigFile'),
    readFile: (filePath) => ipcRenderer.invoke('readFile', filePath),
    selectFolder: () => ipcRenderer.invoke('select-folder'),
    scanFolder: (folderPath) => ipcRenderer.invoke('scan-folder', folderPath)
});