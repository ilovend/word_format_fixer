// @ts-check
const { defineConfig } = require('@playwright/test');

/**
 * Playwright 测试配置 - Electron 应用测试
 * @see https://playwright.dev/docs/test-configuration
 */
module.exports = defineConfig({
    testDir: './tests/e2e',

    // 测试超时时间
    timeout: 30 * 1000,

    // 期望超时
    expect: {
        timeout: 5000
    },

    // 测试报告
    reporter: [
        ['list'],
        ['html', { outputFolder: 'tests/report', open: 'never' }]
    ],

    // 重试次数（CI 环境）
    retries: process.env.CI ? 2 : 0,

    // 并行测试数量
    workers: 1, // Electron 测试建议串行

    // 全局配置
    use: {
        // 截图（仅失败时）
        screenshot: 'only-on-failure',

        // 视频录制
        video: 'retain-on-failure',

        // 追踪
        trace: 'on-first-retry',
    },
});
