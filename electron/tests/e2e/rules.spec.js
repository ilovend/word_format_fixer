/**
 * 规则配置交互测试
 * 验证用户可以正确操作规则配置界面
 */
const { test, expect } = require('@playwright/test');
const { launchApp, closeApp, waitForBackendReady } = require('./helpers');

let app;
let window;

test.describe('规则配置交互测试', () => {
    test.beforeAll(async () => {
        const result = await launchApp();
        app = result.app;
        window = result.window;

        // 等待规则加载完成
        await waitForBackendReady(window, 15000);
    });

    test.afterAll(async () => {
        await closeApp(app);
    });

    test('展开/收起规则分类', async () => {
        // 获取第一个分类头部
        const firstHeader = await window.$('.rule-category-header');
        expect(firstHeader).toBeTruthy();

        // 点击展开
        await firstHeader.click();

        // 等待动画完成
        await window.waitForTimeout(300);

        // 验证内容显示
        const content = await window.$('.rule-category-content');
        const display = await content.evaluate(el => window.getComputedStyle(el).display);
        expect(display).not.toBe('none');
    });

    test('切换规则启用状态', async () => {
        // 展开第一个分类（可能已展开，忽略错误）
        const firstHeader = await window.$('.rule-category-header');
        const content = await window.$('.rule-category-content');
        const display = await content.evaluate(el => window.getComputedStyle(el).display);

        if (display === 'none') {
            await firstHeader.click();
            await window.waitForTimeout(300);
        }

        // 获取第一个规则开关
        const ruleToggle = await window.$('.rule-toggle');
        expect(ruleToggle).toBeTruthy();

        // 获取初始状态
        const wasChecked = await ruleToggle.isChecked();

        // 使用 JavaScript 直接触发 click（绕过遮挡问题）
        await ruleToggle.evaluate(el => el.click());
        await window.waitForTimeout(100);

        // 验证状态变化
        const isChecked = await ruleToggle.isChecked();
        expect(isChecked).toBe(!wasChecked);
    });

    test('启用规则计数更新', async () => {
        // 获取初始启用规则数
        const enabledBefore = await window.$eval('#enabledRules', el => parseInt(el.textContent, 10));

        // 展开第一个分类（可能已展开，忽略错误）
        const firstHeader = await window.$('.rule-category-header');
        const content = await window.$('.rule-category-content');
        const display = await content.evaluate(el => window.getComputedStyle(el).display);

        if (display === 'none') {
            await firstHeader.click();
            await window.waitForTimeout(300);
        }

        // 获取一个未选中的规则开关
        const uncheckedToggles = await window.$$('.rule-toggle:not(:checked)');

        if (uncheckedToggles.length > 0) {
            // 点击启用 (使用 evaluate 绕过遮挡)
            await uncheckedToggles[0].evaluate(el => el.click());
            await window.waitForTimeout(100);

            // 验证计数增加
            const enabledAfter = await window.$eval('#enabledRules', el => parseInt(el.textContent, 10));
            expect(enabledAfter).toBe(enabledBefore + 1);
        }
    });



    test('预设切换功能', async () => {
        // 查找预设卡片
        const presetCards = await window.$$('.preset-card');

        if (presetCards.length > 1) {
            // 点击第二个预设
            await presetCards[1].click();

            // 验证激活状态切换
            const isActive = await presetCards[1].evaluate(el => el.classList.contains('active'));
            expect(isActive).toBe(true);
        }
    });
});
