
const { app, BrowserWindow, ipcMain, dialog, Menu } = require('electron');
const path = require('path');
const { spawn } = require('child_process');
const fs = require('fs');

// 禁用 GPU 硬件加速，解决部分环境下的 crash 问题
app.disableHardwareAcceleration();

let mainWindow;

// 全局 Python 进程管理器
class PythonProcessManager {
    constructor() {
        /** @type {import('child_process').ChildProcess | null} */
        this.process = null;
        this.pendingRequests = new Map();
        this.rl = null;
        this.isReady = false;
        this.restartCount = 0;
        this.maxRestarts = 3;
    }

    // 获取后端可执行文件路径
    getBackendPath() {
        const isDev = !app.isPackaged;
        if (isDev) {
            return path.join(__dirname, 'word_format_fixer_backend', 'word_format_fixer_backend.exe');
        } else {
            return path.join(process.resourcesPath, 'word_format_fixer_backend', 'word_format_fixer_backend.exe');
        }
    }

    // 启动持久化进程
    start() {
        if (this.process) return;

        const isDev = !app.isPackaged;
        let spawnCmd, spawnArgs, spawnOptions;

        if (isDev) {
            // 开发环境：直接运行 Python 源码
            // 尝试定位 python-backend/cli.py
            const scriptPath = path.join(__dirname, '..', 'python-backend', 'cli.py');
            console.log(`Development Mode: Starting Python Source: ${scriptPath}`);

            // 简单检测系统中的 python 命令
            const pythonCmd = process.platform === 'win32' ? 'python' : 'python3';

            spawnCmd = pythonCmd;
            spawnArgs = [scriptPath, '--interactive'];
            // @ts-ignore: TS definition mismatch for stdio array in JS
            spawnOptions = {
                cwd: path.dirname(scriptPath),
                stdio: 'pipe' // 使用 pipe 进行交互
            };
        } else {
            // 生产环境：运行打包好的 exe
            const backendPath = this.getBackendPath();
            console.log(`Production Mode: Starting Python Backend: ${backendPath}`);

            spawnCmd = backendPath;
            spawnArgs = ['--interactive'];
            // @ts-ignore: TS definition mismatch for stdio array in JS
            spawnOptions = {
                cwd: path.dirname(backendPath),
                stdio: 'pipe'
            };
        }

        try {
            console.log(`Spawning: ${spawnCmd} ${spawnArgs.join(' ')}`);
            const proc = spawn(spawnCmd, spawnArgs, spawnOptions);
            this.process = proc;

            if (!proc) {
                console.error('Failed to spawn process');
                return;
            }

            // 设置编码
            if (proc.stdout) proc.stdout.setEncoding('utf8');
            if (proc.stderr) proc.stderr.setEncoding('utf8');

            // 处理输出流
            const readline = require('readline');
            if (proc.stdout) {
                this.rl = readline.createInterface({
                    input: proc.stdout,
                    terminal: false
                });

                this.rl.on('line', (line) => {
                    this.handleResponse(line);
                });
            }

            // 错误日志
            if (proc.stderr) {
                proc.stderr.on('data', (data) => {
                    console.error(`Python Backend Error: ${data}`);
                });
            }

            // 进程退出处理
            proc.on('close', (code) => {
                console.log(`Python Backend exited with code ${code}`);
                // 只有当当前进程仍然是这个进程时才清理（防止并发启动导致的问题）
                if (this.process === proc) {
                    this.cleanup();

                    // 异常退出尝试重启
                    if (code !== 0 && this.restartCount < this.maxRestarts) {
                        console.log('Attempting to restart backend...');
                        this.restartCount++;
                        setTimeout(() => this.start(), 1000);
                    }
                }
            });

        } catch (error) {
            console.error('Failed to start Python backend:', error);
        }
    }

    // 处理后端响应
    handleResponse(line) {
        try {
            const response = JSON.parse(line);

            // 处理就绪信号
            if (response.status === 'ready') {
                console.log('Python Backend Ready, PID:', response.pid);
                this.isReady = true;
                this.restartCount = 0;
                return;
            }

            // 处理请求响应
            const { id, success, result, error } = response;
            if (id && this.pendingRequests.has(id)) {
                const { resolve, reject } = this.pendingRequests.get(id);
                this.pendingRequests.delete(id);

                if (success) {
                    resolve(result);
                } else {
                    reject(new Error(error || 'Unknown error from backend'));
                }
            }
        } catch (e) {
            console.error('Failed to parse backend response:', e, line);
        }
    }

    // 发送请求
    send(command, data) {
        return new Promise((resolve, reject) => {
            if (!this.process || !this.isReady) {
                // 如果进程未就绪，尝试启动
                if (!this.process) this.start();

                // 等待就绪（最多2秒）
                const checkReady = setInterval(() => {
                    if (this.isReady) {
                        clearInterval(checkReady);
                        this._sendInternal(command, data, resolve, reject);
                    }
                }, 100);

                setTimeout(() => {
                    clearInterval(checkReady);
                    if (!this.isReady) {
                        reject(new Error('Backend not ready'));
                    }
                }, 2000);
            } else {
                this._sendInternal(command, data, resolve, reject);
            }
        });
    }

    _sendInternal(command, data, resolve, reject) {
        // Double check process and stdin
        if (!this.process || !this.process.stdin) {
            reject(new Error('Backend process not available'));
            return;
        }

        const requestId = Date.now().toString(36) + Math.random().toString(36).substr(2);
        const payload = JSON.stringify({
            id: requestId,
            command: command,
            data: data
        }) + '\n';

        this.pendingRequests.set(requestId, { resolve, reject });

        try {
            this.process.stdin.write(payload);
        } catch (error) {
            this.pendingRequests.delete(requestId);
            reject(error);
        }
    }

    // 清理资源
    cleanup() {
        this.isReady = false;
        this.process = null;
        this.rl = null;
        // 拒绝所有挂起的请求
        for (const [id, { reject }] of this.pendingRequests) {
            reject(new Error('Backend process terminated'));
        }
        this.pendingRequests.clear();
    }

    // 停止进程
    stop() {
        if (this.process) {
            this.process.kill();
            this.cleanup();
        }
    }
}

// 实例化管理器
const pythonManager = new PythonProcessManager();

// 修改 callPythonCLI 使用持久化进程
async function callPythonCLI(command, data) {
    try {
        return await pythonManager.send(command, data);
    } catch (error) {
        console.warn('Persistent process failed, falling back to one-off spawn:', error.message);
        return callPythonCLIOneOff(command, data);
    }
}

// 原有的一次性调用函数（作为回退）
function callPythonCLIOneOff(command, data) {
    return new Promise((resolve, reject) => {
        // ... (原有的实现逻辑保持不变，用于兼容或降级)
        const backendPath = pythonManager.getBackendPath();
        const args = [command, JSON.stringify(data)];

        console.log(`Calling Python CLI (One-off): ${backendPath} ${args.join(' ')}`);

        const pythonProcess = spawn(backendPath, args, {
            cwd: path.join(path.dirname(backendPath)),
            stdio: ['pipe', 'pipe', 'pipe']
        });

        let output = '';
        let errorOutput = '';

        pythonProcess.stdout.on('data', (data) => { output += data.toString(); });
        pythonProcess.stderr.on('data', (data) => { errorOutput += data.toString(); });

        pythonProcess.on('close', (code) => {
            if (code === 0) {
                try {
                    resolve(JSON.parse(output));
                } catch (e) {
                    reject(new Error(`Failed to parse Python output: ${e.message}`));
                }
            } else {
                reject(new Error(errorOutput || `Python process exited with code ${code}`));
            }
        });

        pythonProcess.on('error', (err) => reject(err));
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
        // 窗口显示后启动后端进程
        pythonManager.start();
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
    // 退出前清理后端进程
    pythonManager.stop();
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

app.on('quit', () => {
    pythonManager.stop();
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

ipcMain.handle('configure-rules', async (event, configs) => {
    try {
        return await callPythonCLI('configure-rules', { configs: configs });
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