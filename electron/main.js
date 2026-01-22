const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const { spawn } = require('child_process');
const axios = require('axios');

let mainWindow;
let pythonProcess;
let apiPort = null; // 动态获取端口

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1000,  // 调整默认宽度
        height: 750,  // 调整默认高度
        minWidth: 700,  // 设置最小宽度，保证基础内容不挤压
        minHeight: 550, // 设置最小高度
        show: false,     // 先隐藏，加载完再显示防止白屏
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            nodeIntegration: false,
            contextIsolation: true,
            enableRemoteModule: true
        }
    });

    // 加载前端页面
    mainWindow.loadFile(path.join(__dirname, 'index.html'));

    // 加载完成后显示窗口
    mainWindow.once('ready-to-show', () => {
        mainWindow.show();
    });

    mainWindow.on('closed', function () {
        mainWindow = null;
    });
}

function startPythonServer() {
    console.log('Starting Python server...');
    pythonProcess = spawn('python', ['python-backend/main.py'], {
        cwd: path.join(__dirname, '..'),
        stdio: 'inherit'
    });

    pythonProcess.on('error', (err) => {
        console.error('Failed to start Python server:', err);
        dialog.showErrorBox('错误', '无法启动Python服务器');
    });

    pythonProcess.on('close', (code) => {
        console.log(`Python server exited with code ${code}`);
    });
}

app.on('ready', () => {
    createWindow();
    startPythonServer();
});

app.on('window-all-closed', function () {
    if (process.platform !== 'darwin') {
        if (pythonProcess) {
            pythonProcess.kill();
        }
        app.quit();
    }
});

app.on('activate', function () {
    if (mainWindow === null) {
        createWindow();
    }
});

// 读取端口文件获取后端端口
async function getApiPort() {
    const fs = require('fs');
    const path = require('path');
    const portFilePath = path.join(__dirname, '..', '.port');

    try {
        // 尝试读取端口文件
        if (fs.existsSync(portFilePath)) {
            const port = parseInt(fs.readFileSync(portFilePath, 'utf-8').trim());
            if (!isNaN(port)) {
                console.log(`Read port from file: ${port}`);
                return port;
            }
        }
    } catch (error) {
        console.log('Failed to read port file:', error.message);
    }

    // 如果端口文件不存在或读取失败，使用默认端口
    console.log('Using default port 7777');
    return 7777;
}

// IPC处理
ipcMain.handle('process-document', async (event, documentPath, activeRules) => {
    try {
        // 获取正确的端口
        if (!apiPort) {
            apiPort = await getApiPort();
        }
        
        const response = await axios.post(`http://127.0.0.1:${apiPort}/api/process`, {
            file_path: documentPath,
            active_rules: activeRules
        });
        return response.data;
    } catch (error) {
        console.error('Error processing document:', error);
        throw error;
    }
});

ipcMain.handle('get-rules', async () => {
    try {
        // 获取正确的端口
        if (!apiPort) {
            apiPort = await getApiPort();
        }
        
        const response = await axios.get(`http://127.0.0.1:${apiPort}/api/rules`);
        return response.data;
    } catch (error) {
        console.error('Error getting rules:', error);
        throw error;
    }
});

ipcMain.handle('get-presets', async () => {
    try {
        // 获取正确的端口
        if (!apiPort) {
            apiPort = await getApiPort();
        }
        
        const response = await axios.get(`http://127.0.0.1:${apiPort}/api/presets`);
        return response.data;
    } catch (error) {
        console.error('Error getting presets:', error);
        throw error;
    }
});

ipcMain.handle('configure-rules', async (event, configs) => {
    try {
        // 获取正确的端口
        if (!apiPort) {
            apiPort = await getApiPort();
        }
        
        const response = await axios.post(`http://127.0.0.1:${apiPort}/api/rules/config`, configs);
        return response.data;
    } catch (error) {
        console.error('Error configuring rules:', error);
        throw error;
    }
});

ipcMain.handle('select-file', async () => {
    try {
        const result = await dialog.showOpenDialog({
            properties: ['openFile'],
            filters: [
                { name: 'Word Documents', extensions: ['docx'] }
            ]
        });
        
        if (!result.canceled && result.filePaths.length > 0) {
            return result.filePaths[0];
        }
        return null;
    } catch (error) {
        console.error('Error selecting file:', error);
        throw error;
    }
});

ipcMain.handle('save-preset', async (event, presetId, presetData) => {
    try {
        // 获取正确的端口
        if (!apiPort) {
            apiPort = await getApiPort();
        }
        
        const response = await axios.post(`http://127.0.0.1:${apiPort}/api/presets/save`, {
            preset_id: presetId,
            preset_data: presetData
        });
        return response.data;
    } catch (error) {
        console.error('Error saving preset:', error);
        throw error;
    }
});

ipcMain.handle('delete-preset', async (event, presetId) => {
    try {
        // 获取正确的端口
        if (!apiPort) {
            apiPort = await getApiPort();
        }
        
        const response = await axios.delete(`http://127.0.0.1:${apiPort}/api/presets/delete`, {
            data: { preset_id: presetId }
        });
        return response.data;
    } catch (error) {
        console.error('Error deleting preset:', error);
        throw error;
    }
});

// 配置文件导入
ipcMain.handle('importConfigFile', async () => {
    try {
        const result = await dialog.showOpenDialog({
            properties: ['openFile'],
            filters: [
                { name: '配置文件', extensions: ['json'] },
                { name: '所有文件', extensions: ['*'] }
            ]
        });
        
        if (!result.canceled && result.filePaths.length > 0) {
            return result.filePaths[0];
        }
        return null;
    } catch (error) {
        console.error('Error selecting config file:', error);
        throw error;
    }
});

// 读取文件内容
ipcMain.handle('readFile', async (event, filePath) => {
    try {
        const fs = require('fs');
        return fs.readFileSync(filePath, 'utf-8');
    } catch (error) {
        console.error('Error reading file:', error);
        throw error;
    }
});