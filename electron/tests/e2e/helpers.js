/**
 * Electron E2E 测试工具模块
 * 提供 Electron 应用启动和窗口管理功能
 */
const { _electron: electron } = require('playwright');
const path = require('path');

/**
 * 启动 Electron 应用
 * @returns {Promise<{app: import('playwright').ElectronApplication, window: import('playwright').Page}>}
 */
async function launchApp() {
    const mainPath = path.join(__dirname, '../../main.js');

    const app = await electron.launch({
        args: [mainPath],
        cwd: path.join(__dirname, '../..'),
    });

    // 获取第一个窗口
    const window = await app.firstWindow();

    // 等待页面加载完成
    await window.waitForLoadState('domcontentloaded');

    return { app, window };
}

/**
 * 关闭 Electron 应用
 * @param {import('playwright').ElectronApplication} app 
 */
async function closeApp(app) {
    if (app) {
        await app.close();
    }
}

/**
 * 等待后端响应（规则加载完成）
 * @param {import('playwright').Page} window 
 * @param {number} timeout 
 */
async function waitForBackendReady(window, timeout = 10000) {
    // 等待规则加载完成的状态指示
    await window.waitForFunction(
        () => {
            const status = document.getElementById('status');
            return status && (
                status.textContent.includes('规则加载完成') ||
                status.textContent.includes('加载失败')
            );
        },
        { timeout }
    );
}

module.exports = {
    launchApp,
    closeApp,
    waitForBackendReady,
};
