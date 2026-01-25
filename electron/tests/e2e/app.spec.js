/**
 * Electron 应用启动测试
 * 验证应用能够正常启动并加载基础界面
 */
const { test, expect } = require('@playwright/test');
const { launchApp, closeApp, waitForBackendReady } = require('./helpers');

let app;
let window;

test.describe('应用启动测试', () => {
    test.beforeAll(async () => {
        const result = await launchApp();
        app = result.app;
        window = result.window;
    });

    test.afterAll(async () => {
        await closeApp(app);
    });

    test('应用窗口正常加载', async () => {
        // 验证窗口存在
        expect(window).toBeTruthy();

        // 验证窗口标题
        const title = await window.title();
        expect(title).toContain('Word');
    });

    test('主界面元素加载', async () => {
        // 验证主要 UI 元素存在

        // 文件选择区域
        const dropZone = await window.$('#dropZone');
        expect(dropZone).toBeTruthy();

        // 处理按钮
        const processBtn = await window.$('#processFile');
        expect(processBtn).toBeTruthy();

        // 状态指示器
        const status = await window.$('#status');
        expect(status).toBeTruthy();
    });

    test('规则从后端加载成功', async () => {
        // 等待后端响应
        await waitForBackendReady(window, 15000);

        // 获取状态文本
        const statusText = await window.$eval('#status', el => el.textContent);

        // 验证规则加载成功（而非失败）
        expect(statusText).toContain('规则加载完成');
    });

    test('规则分类正确渲染', async () => {
        // 等待规则加载完成
        await waitForBackendReady(window, 15000);

        // 验证至少存在一个规则分类
        const categoryHeaders = await window.$$('.rule-category-header');
        expect(categoryHeaders.length).toBeGreaterThan(0);
    });

    test('规则计数正确显示', async () => {
        // 等待规则加载完成
        await waitForBackendReady(window, 15000);

        // 获取总规则数
        const totalRulesText = await window.$eval('#totalRules', el => el.textContent);
        const totalRules = parseInt(totalRulesText, 10);

        // 验证规则数量大于0
        expect(totalRules).toBeGreaterThan(0);
    });
});
