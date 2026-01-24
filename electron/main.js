const { app, BrowserWindow, ipcMain, dialog, Menu } = require('electron');
const path = require('path');
const { spawn } = require('child_process');
const fs = require('fs');

let mainWindow;

// 直接调用Python CLI工具处理请求
function callPythonCLI(command, data) {
    return new Promise((resolve, reject) => {
        // 确定Python后端可执行文件的路径
        let backendPath;
        // 判断是否为开发环境
        // 在开发环境中，__dirname指向electron目录
        // 在打包环境中，process.resourcesPath指向resources目录
        const isDev = !app.isPackaged;
        
        if (isDev) {
            // 开发环境：使用electron目录下的后端目录
            backendPath = path.join(__dirname, 'word_format_fixer_backend', 'word_format_fixer_backend.exe');
        } else {
            // 打包环境：使用resources目录下的后端目录
            backendPath = path.join(process.resourcesPath, 'word_format_fixer_backend', 'word_format_fixer_backend.exe');
        }
        
        // 构建命令行参数
        const args = [command, JSON.stringify(data)];
        
        console.log(`Calling Python CLI: ${backendPath} ${args.join(' ')}`);
        
        // 启动Python进程
        const pythonProcess = spawn(backendPath, args, {
            cwd: path.join(path.dirname(backendPath)),
            stdio: ['pipe', 'pipe', 'pipe']
        });
        
        let output = '';
        let errorOutput = '';
        
        // 收集输出
        pythonProcess.stdout.on('data', (data) => {
            output += data.toString();
        });
        
        pythonProcess.stderr.on('data', (data) => {
            errorOutput += data.toString();
        });
        
        // 处理进程结束
        pythonProcess.on('close', (code) => {
            console.log(`Python CLI exited with code ${code}`);
            
            if (code === 0) {
                try {
                    // 解析JSON输出
                    const result = JSON.parse(output);
                    resolve(result);
                } catch (e) {
                    console.error('Failed to parse Python output:', e.message);
                    console.error('Raw output:', output);
                    reject(new Error(`Failed to parse Python output: ${e.message}`));
                }
            } else {
                const errorMsg = errorOutput || `Python process exited with code ${code}`;
                console.error('Python CLI error:', errorMsg);
                reject(new Error(errorMsg));
            }
        });
        
        // 处理进程错误
        pythonProcess.on('error', (err) => {
            console.error('Failed to start Python process:', err);
            reject(err);
        });
    });
}

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1000,  // 调整默认宽度
        height: 750,  // 调整默认高度
        minWidth: 700,  // 设置最小宽度，保证基础内容不挤压
        minHeight: 550, // 设置最小高度
        show: false,     // 先隐藏，加载完再显示防止白屏
        icon: path.join(__dirname, 'build', 'icon.ico'), // 设置应用图标
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            nodeIntegration: false,
            contextIsolation: true
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

app.on('ready', () => {
    // 创建窗口前先隐藏菜单栏
    Menu.setApplicationMenu(null);
    createWindow();
});

app.on('window-all-closed', function () {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

app.on('activate', function () {
    if (mainWindow === null) {
        createWindow();
    }
});

// IPC处理 - 直接调用Python CLI

ipcMain.handle('process-document', async (event, documentPath, activeRules) => {
    try {
        return await callPythonCLI('process-document', {
            file_path: documentPath,
            active_rules: activeRules
        });
    } catch (error) {
        console.error('Error processing document:', error);
        throw error;
    }
});

ipcMain.handle('get-rules', async () => {
    try {
        return await callPythonCLI('get-rules', {});
    } catch (error) {
        console.error('Error getting rules:', error);
        throw error;
    }
});

ipcMain.handle('get-presets', async () => {
    try {
        return await callPythonCLI('get-presets', {});
    } catch (error) {
        console.error('Error getting presets:', error);
        throw error;
    }
});

ipcMain.handle('save-preset', async (event, presetId, presetData) => {
    try {
        return await callPythonCLI('save-preset', {
            preset_id: presetId,
            preset_data: presetData
        });
    } catch (error) {
        console.error('Error saving preset:', error);
        throw error;
    }
});

ipcMain.handle('delete-preset', async (event, presetId) => {
    try {
        return await callPythonCLI('delete-preset', {
            preset_id: presetId
        });
    } catch (error) {
        console.error('Error deleting preset:', error);
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
        return fs.readFileSync(filePath, 'utf-8');
    } catch (error) {
        console.error('Error reading file:', error);
        throw error;
    }
});

// 获取版本号
ipcMain.handle('get-version', async () => {
    try {
        // 读取VERSION文件
        const versionPath = path.join(__dirname, '..', 'VERSION');
        if (fs.existsSync(versionPath)) {
            return fs.readFileSync(versionPath, 'utf-8').trim();
        }
        // 如果VERSION文件不存在，使用package.json中的版本号
        const pkgPath = path.join(__dirname, 'package.json');
        const pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf-8'));
        return pkg.version;
    } catch (error) {
        console.error('Error getting version:', error);
        return '1.0.0';
    }
});