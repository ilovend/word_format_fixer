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
    scanFolder: (folderPath) => ipcRenderer.invoke('scan-folder', folderPath),
    // Diff/对比功能
    prepareDiff: (documentPath) => ipcRenderer.invoke('prepare-diff', documentPath),
    generateDiff: (documentPath) => ipcRenderer.invoke('generate-diff', documentPath),
    getPreview: (documentPath) => ipcRenderer.invoke('get-preview', documentPath),
    // 自动更新功能
    checkForUpdates: () => ipcRenderer.invoke('check-for-updates'),
    quitAndInstall: () => ipcRenderer.invoke('quit-and-install'),
    onUpdateChecking: (callback) => ipcRenderer.on('update-checking', callback),
    onUpdateAvailable: (callback) => ipcRenderer.on('update-available', (event, info) => callback(info)),
    onUpdateNotAvailable: (callback) => ipcRenderer.on('update-not-available', callback),
    onUpdateError: (callback) => ipcRenderer.on('update-error', (event, message) => callback(message)),
    onUpdateDownloadProgress: (callback) => ipcRenderer.on('update-download-progress', (event, progress) => callback(progress)),
    onUpdateDownloaded: (callback) => ipcRenderer.on('update-downloaded', (event, info) => callback(info))
});