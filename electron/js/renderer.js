let selectedFilePath = '';
let apiPort = null;
let backendRules = [];
let presets = {};  // 预设数据将从后端加载

// 初始化
// 初始化
document.addEventListener('DOMContentLoaded', function () {
    initializeEventListeners();

    // @ts-ignore
    if (window.electronAPI) {
        fetchRulesFromBackend();
        fetchPresetsFromBackend();
    } else {
        console.error('Electron API not available');
        updateStatus('错误: 无法连接到应用程序核心', 'error');
    }
});

// 初始化事件监听器
// 初始化事件监听器
function initializeEventListeners() {
    // 辅助函数：安全绑定事件
    const safeBind = (id, event, handler) => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener(event, handler);
        } else {
            console.warn(`Element #${id} not found, skipping event binding.`);
        }
    };

    // 文件选择
    safeBind('selectFile', 'click', selectFile);
    safeBind('clearFile', 'click', clearFile);

    // 处理文档
    safeBind('processFile', 'click', processFile);

    // 预设选择
    document.querySelectorAll('.preset-card').forEach(card => {
        card.addEventListener('click', function () {
            // @ts-ignore
            selectPreset(this.dataset.preset);
        });
    });

    // 拖拽功能
    const dropZone = document.getElementById('dropZone');
    if (dropZone) {
        dropZone.addEventListener('dragover', handleDragOver);
        dropZone.addEventListener('drop', handleDrop);
        dropZone.addEventListener('dragenter', handleDragEnter);
        dropZone.addEventListener('dragleave', handleDragLeave);
    }

    // 配置管理
    safeBind('saveConfig', 'click', saveCurrentConfig);
    safeBind('loadConfig', 'click', loadConfig);
    safeBind('exportConfig', 'click', exportConfig);
    safeBind('importConfig', 'click', importConfig);
}

// 保存当前配置到本地存储
function saveCurrentConfig() {
    try {
        const enabledRules = getEnabledRules();
        if (enabledRules.length === 0) {
            updateStatus('没有启用的规则，无需保存配置', 'warning');
            return;
        }

        const config = {
            enabledRules: enabledRules,
            timestamp: new Date().toISOString()
        };

        // 保存到localStorage
        localStorage.setItem('wordFormatFixerConfig', JSON.stringify(config));

        // 提示用户
        updateStatus('配置已成功保存', 'success');
        logDev('当前配置已保存到本地存储');
    } catch (e) {
        updateStatus(`配置保存失败: ${e.message}`, 'error');
        logDev(`配置保存失败: ${e.message}`);
    }
}

// 从本地存储加载配置
function loadConfig() {
    try {
        const configStr = localStorage.getItem('wordFormatFixerConfig');
        if (!configStr) {
            updateStatus('没有找到保存的配置', 'warning');
            return;
        }

        const config = JSON.parse(configStr);
        if (!config.enabledRules || !Array.isArray(config.enabledRules)) {
            updateStatus('配置格式不正确', 'error');
            return;
        }

        applyConfig(config.enabledRules);
        updateStatus('配置已成功加载', 'success');
        logDev('从本地存储加载配置');
    } catch (e) {
        updateStatus(`加载配置失败: ${e.message}`, 'error');
        logDev(`加载配置失败: ${e.message}`);
    }
}

// 应用配置
function applyConfig(enabledRules) {
    if (!enabledRules || enabledRules.length === 0) {
        updateStatus('配置中没有启用的规则', 'warning');
        return;
    }

    // 先禁用所有规则
    document.querySelectorAll('.rule-toggle').forEach(toggle => {
        toggle.checked = false;
    });

    // 启用配置中的规则并设置参数
    let appliedRuleCount = 0;
    enabledRules.forEach(ruleConfig => {
        const ruleId = ruleConfig.rule_id;
        const params = ruleConfig.params;

        // 启用规则
        const toggle = document.querySelector(`[data-rule="${ruleId}"].rule-toggle`);
        if (toggle) {
            toggle.checked = true;
            appliedRuleCount++;

            // 设置参数
            for (const [paramName, paramValue] of Object.entries(params)) {
                const paramControl = document.querySelector(`[data-rule="${ruleId}"][data-param="${paramName}"]`);
                if (paramControl) {
                    if (paramControl.type === 'checkbox') {
                        paramControl.checked = paramValue;
                    } else {
                        paramControl.value = paramValue;
                    }
                }
            }
        }
    });

    updateRuleCounts();
    logDev(`成功应用了 ${appliedRuleCount} 个规则配置`);
}

// 导出配置为JSON文件
function exportConfig() {
    try {
        const enabledRules = getEnabledRules();
        if (enabledRules.length === 0) {
            updateStatus('没有启用的规则，无需导出配置', 'warning');
            return;
        }

        const config = {
            enabledRules: enabledRules,
            timestamp: new Date().toISOString(),
            version: '1.0.0'
        };

        const configStr = JSON.stringify(config, null, 2);
        const blob = new Blob([configStr], { type: 'application/json' });
        const url = URL.createObjectURL(blob);

        // 创建下载链接
        const a = document.createElement('a');
        a.href = url;
        a.download = `word_format_fixer_config_${new Date().toISOString().slice(0, 10)}.json`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        updateStatus('配置已成功导出', 'success');
        logDev('配置已导出为JSON文件');
    } catch (e) {
        updateStatus(`配置导出失败: ${e.message}`, 'error');
        logDev(`配置导出失败: ${e.message}`);
    }
}

// 导入配置
async function importConfig() {
    try {
        const filePath = await window.electronAPI.importConfigFile();
        if (!filePath) {
            logDev('用户取消了配置导入');
            return;
        }

        const configStr = await window.electronAPI.readFile(filePath);
        if (!configStr) {
            updateStatus('导入的文件为空', 'error');
            return;
        }

        const config = JSON.parse(configStr);

        if (config.enabledRules && Array.isArray(config.enabledRules)) {
            applyConfig(config.enabledRules);
            updateStatus('配置已成功导入', 'success');
            logDev(`从文件导入配置: ${filePath}`);
        } else {
            updateStatus('导入的配置格式不正确，缺少enabledRules字段', 'error');
            logDev('导入的配置格式不正确');
        }
    } catch (e) {
        updateStatus(`导入配置失败: ${e.message}`, 'error');
        logDev(`导入配置失败: ${e.message}`);
    }
}

// 选择文件
async function selectFile() {
    try {
        logDev('选择文件...');
        const filePath = await window.electronAPI.selectFile();
        if (filePath) {
            selectedFilePath = filePath;
            updateFileInfo();
            logDev(`文件选择成功: ${filePath}`);
        }
    } catch (error) {
        console.error('Error selecting file:', error);
        updateStatus('错误: ' + error.message, 'error');
        logDev(`文件选择错误: ${error.message}`);
    }
}

// 清除文件选择
function clearFile() {
    selectedFilePath = '';
    updateFileInfo();
    updateStatus('');
    document.getElementById('result').textContent = '';
    logDev('文件选择已清除');
}

// 更新文件信息
function updateFileInfo() {
    const filePathElement = document.getElementById('filePath');
    const processButton = document.getElementById('processFile');

    if (selectedFilePath) {
        filePathElement.textContent = selectedFilePath;
        processButton.disabled = false;
    } else {
        filePathElement.textContent = '未选择文件';
        processButton.disabled = true;
    }
}

// 处理文档
async function processFile() {
    if (!selectedFilePath) {
        updateStatus('请先选择要处理的文档', 'warning');
        return;
    }

    const enabledRules = getEnabledRules();
    if (enabledRules.length === 0) {
        updateStatus('请至少启用一个规则', 'warning');
        return;
    }

    updateStatus('正在处理...', 'processing');
    showProgressBar();
    logDev('开始处理文档...');
    logDev(`文档路径: ${selectedFilePath}`);
    logDev(`启用的规则数量: ${enabledRules.length}`);

    try {
        // 发送请求到后端
        const result = await window.electronAPI.processDocument(selectedFilePath, enabledRules);
        logDev(`处理结果: ${JSON.stringify(result)}`);

        hideProgressBar();

        if (result.save_success) {
            updateStatus('处理完成，文件已保存', 'success');
        } else {
            updateStatus('处理完成，保存失败', 'error');
        }

        // 显示详细结果
        displayResult(result);
    } catch (error) {
        console.error('Error processing document:', error);
        hideProgressBar();

        let errorMessage = '处理错误';
        if (error.message) {
            if (error.message.includes('connect')) {
                errorMessage = `无法连接到后端服务: ${error.message}`;
            } else if (error.message.includes('permission')) {
                errorMessage = `权限不足: ${error.message}`;
            } else {
                errorMessage = `处理错误: ${error.message}`;
            }
        }

        updateStatus(errorMessage, 'error');
        document.getElementById('result').textContent = errorMessage;
        logDev(`处理错误: ${error.message || error}`);
    }
}

// 获取启用的规则
function getEnabledRules() {
    const enabledRules = [];
    document.querySelectorAll('.rule-toggle:checked').forEach(toggle => {
        const ruleId = toggle.dataset.rule;
        const params = {};

        // 获取该规则的所有参数控件
        const paramControls = document.querySelectorAll(`[data-rule="${ruleId}"][data-param]`);
        paramControls.forEach(control => {
            const paramName = control.dataset.param;
            let paramValue;

            // 根据控件类型获取值
            if (control.type === 'checkbox') {
                paramValue = control.checked;
            } else if (control.type === 'number') {
                paramValue = parseFloat(control.value);
            } else if (control.tagName === 'SELECT') {
                paramValue = control.value;
            } else {
                paramValue = control.value;
                // 尝试转换为JSON（如果是复杂类型）
                try {
                    paramValue = JSON.parse(control.value);
                } catch (e) {
                    // 保持字符串类型
                }
            }

            params[paramName] = paramValue;
        });

        enabledRules.push({
            rule_id: ruleId,
            params: params
        });
    });
    return enabledRules;
}

// 更新规则计数
function updateRuleCounts() {
    const totalRules = document.querySelectorAll('.rule-toggle').length;
    const enabledRules = document.querySelectorAll('.rule-toggle:checked').length;

    document.getElementById('totalRules').textContent = totalRules;
    document.getElementById('enabledRules').textContent = enabledRules;
}

// 选择预设
function selectPreset(presetId) {
    // 移除所有预设的激活状态
    document.querySelectorAll('.preset-card').forEach(card => {
        card.classList.remove('active');
    });

    // 激活当前预设
    document.querySelector(`[data-preset="${presetId}"]`).classList.add('active');

    // 根据预设设置规则
    applyPreset(presetId);
    logDev(`选择预设: ${presetId}`);
}

// 应用预设
function applyPreset(presetId) {
    // 重置所有规则
    document.querySelectorAll('.rule-toggle').forEach(toggle => {
        toggle.checked = false;
    });

    // 根据预设启用规则
    if (presets[presetId]) {
        const presetRules = presets[presetId].rules;
        const enabledRuleCount = Object.keys(presetRules).filter(ruleId =>
            presetRules[ruleId].enabled
        ).length;

        for (const ruleId in presetRules) {
            if (presetRules[ruleId].enabled) {
                const toggle = document.querySelector(`[data-rule="${ruleId}"]`);
                if (toggle) {
                    toggle.checked = true;
                }
            }
        }

        updateRuleCounts();
        logDev(`应用预设: ${presetId}, 启用规则数: ${enabledRuleCount}`);
    }
}

// 更新状态
function updateStatus(message, type = '') {
    const statusElement = document.getElementById('status');
    statusElement.textContent = message;

    // 移除所有状态类
    statusElement.className = 'status-indicator';

    // 添加对应状态类
    if (type) {
        statusElement.classList.add('status-' + type);
    }
}

// 显示结果
function displayResult(result) {
    const resultElement = document.getElementById('result');

    // 构建结果HTML
    let resultHtml = '';

    // 统计信息
    resultHtml += '<h4>处理统计</h4>';

    if (result.summary) {
        resultHtml += `<p>总修复数: ${result.summary.total_fixed || 0}</p>`;
        resultHtml += `<p>处理时间: ${result.summary.time_taken || 'N/A'}</p>`;
    }

    // 详细结果
    resultHtml += '<h4>详细结果</h4>';
    resultHtml += '<ul>';

    if (result.results) {
        result.results.forEach(item => {
            const success = item.success ? '✓' : '✗';
            const ruleName = getRuleNameById(item.rule_id);
            resultHtml += `<li>${success} ${ruleName} - 修复: ${item.fixed_count}个元素</li>`;

            // 显示详细信息
            if (item.details && item.details.length > 0) {
                resultHtml += '<ul style="margin-left: 20px; margin-top: 5px;">';
                item.details.forEach(detail => {
                    resultHtml += `<li style="font-size: 11px;">${detail}</li>`;
                });
                resultHtml += '</ul>';
            }
        });
    }

    resultHtml += '</ul>';

    // 保存信息
    if (result.save_success) {
        resultHtml += '<h4>保存信息</h4>';
        resultHtml += `<p>保存成功: ${result.saved_to || selectedFilePath}</p>`;
    }

    resultElement.innerHTML = resultHtml;
}

// 根据规则ID获取规则名称
function getRuleNameById(ruleId) {
    const rule = backendRules.find(r => r.id === ruleId);
    return rule ? rule.name : ruleId;
}

// 拖拽功能
function handleDragOver(e) {
    e.preventDefault();
    e.stopPropagation();
}

function handleDragEnter(e) {
    e.preventDefault();
    e.stopPropagation();
    document.getElementById('dropZone').classList.add('active');
}

function handleDragLeave(e) {
    e.preventDefault();
    e.stopPropagation();
    document.getElementById('dropZone').classList.remove('active');
}

function handleDrop(e) {
    e.preventDefault();
    e.stopPropagation();
    document.getElementById('dropZone').classList.remove('active');

    if (e.dataTransfer.files.length > 0) {
        const file = e.dataTransfer.files[0];
        if (file.name.endsWith('.docx')) {
            selectedFilePath = file.path;
            updateFileInfo();
            logDev(`文件拖拽: ${file.path}`);
        } else {
            updateStatus('错误: 请选择.docx文件', 'error');
            logDev(`文件类型错误: ${file.name}`);
        }
    }
}

// 规则分类展开/收起
function toggleRuleCategory() {
    const content = this.nextElementSibling;
    const arrow = this.querySelector('span');

    if (content.style.display === 'none') {
        content.style.display = 'block';
        arrow.textContent = '▼';
    } else {
        content.style.display = 'none';
        arrow.textContent = '▶';
    }
}

// 从后端获取规则列表
async function fetchRulesFromBackend() {
    updateStatus('正在加载规则...', 'processing');
    logDev('正在加载规则...');

    try {
        // 通过Electron IPC获取规则
        backendRules = await window.electronAPI.getRules();
        logDev(`规则加载完成: ${backendRules.length} 条规则`);

        if (backendRules.length === 0) {
            updateStatus('未获取到规则，请检查后端连接', 'warning');
            logDev('警告：未获取到任何规则');
            return;
        }

        // 生成规则UI
        generateRuleUI();

        updateStatus('规则加载完成', 'success');
    } catch (error) {
        console.error('Error fetching rules:', error);
        updateStatus(`规则加载失败: ${error.message}`, 'error');
        logDev(`规则加载错误: ${error.message}`);
        // 可以在这里使用默认规则作为备用
    }
}

// 从后端获取预设配置
async function fetchPresetsFromBackend() {
    logDev('正在加载预设配置...');

    try {
        // 通过Electron IPC获取预设
        const backendPresets = await window.electronAPI.getPresets();
        if (backendPresets) {
            presets = backendPresets;
            const presetCount = Object.keys(backendPresets).length;
            logDev(`预设加载完成: ${presetCount} 个预设`);
            // 刷新预设UI
            updatePresetUI();
            logDev('预设UI已更新');
        } else {
            updateStatus('未获取到预设配置', 'warning');
            logDev('警告：未获取到任何预设配置');
        }
    } catch (error) {
        console.error('Error fetching presets:', error);
        updateStatus(`预设加载失败: ${error.message}`, 'error');
        logDev(`预设加载错误: ${error.message}`);
        // 使用默认预设作为备用
    }
}

// 更新预设UI
function updatePresetUI() {
    // 这里可以添加代码来更新预设UI
    // 例如，更新预设卡片的显示
    logDev('预设UI已更新');

    // 自动应用默认预设
    applyPreset('default');
    logDev('自动应用默认预设');
}

// 生成规则UI
function generateRuleUI() {
    const mainPanel = document.querySelector('.main-panel');

    // 清空现有的规则分类
    const existingCategories = mainPanel.querySelectorAll('.rule-category');
    existingCategories.forEach(category => {
        if (!category.querySelector('h3')) return;
        category.remove();
    });

    // 按类别分组规则（从后端元数据读取类别）
    const rulesByCategory = {};

    backendRules.forEach(rule => {
        // 优先从元数据读取类别，如果没有则根据规则ID推断
        let category = rule.category || '其他规则';

        if (!rule.category) {
            // 后向兼容：如果后端没有提供category字段，则根据ID推断
            if (rule.id.includes('Font') || rule.id.includes('font')) {
                category = '字体规则';
            } else if (rule.id.includes('Paragraph') || rule.id.includes('paragraph') || rule.id.includes('Title') || rule.id.includes('title') || rule.id.includes('List') || rule.id.includes('list')) {
                category = '段落规则';
            } else if (rule.id.includes('Table') || rule.id.includes('table')) {
                category = '表格规则';
            } else if (rule.id.includes('Page') || rule.id.includes('page')) {
                category = '页面规则';
            }
        }

        if (!rulesByCategory[category]) {
            rulesByCategory[category] = [];
        }

        rulesByCategory[category].push(rule);
    });

    // 为每个类别创建UI
    Object.entries(rulesByCategory).forEach(([categoryName, rules]) => {
        const categoryElement = document.createElement('div');
        categoryElement.className = 'rule-category';

        // 类别头部
        const header = document.createElement('div');
        header.className = 'rule-category-header';
        header.innerHTML = `
                <h3>${categoryName}</h3>
                <span>▶</span>
            `;

        // 类别内容 - 默认隐藏
        const content = document.createElement('div');
        content.className = 'rule-category-content';
        content.style.display = 'none';

        // 添加规则项
        rules.forEach(rule => {
            const ruleItem = document.createElement('div');
            ruleItem.className = 'rule-item';

            // 生成规则控件HTML
            let ruleControlsHtml = `
                    <div class="rule-info">
                        <h4>${rule.name}</h4>
                        <p>${rule.description}</p>
                    </div>
                    <div class="rule-controls">
                        <label class="toggle-switch">
                            <input type="checkbox" class="rule-toggle" data-rule="${rule.id}">
                            <span class="toggle-slider"></span>
                        </label>
                    </div>`;

            // 生成参数配置HTML - 使用动态表单生成器
            let paramsHtml = '';
            if (window.FormGenerator && (rule.param_schema || rule.params)) {
                // 使用新的 schema-based 表单生成器
                paramsHtml = window.FormGenerator.generateRuleParamsPanel(rule);
            } else if (rule.params && Object.keys(rule.params).length > 0) {
                // 回退：使用旧的内联生成方式
                paramsHtml = `
                        <div class="rule-params" style="margin-top: 10px; padding: 10px; background-color: #f8f9fa; border-radius: 4px; border: 1px solid #e0e0e0;">
                            <h5 style="margin: 0 0 10px 0; font-size: 14px; color: #666;">参数配置</h5>
                            <div class="params-grid" style="display: grid; grid-template-columns: repeat(auto-fill, minmax(250px, 1fr)); gap: 10px;">`;

                // 为每个参数生成配置控件
                for (const [paramName, paramValue] of Object.entries(rule.params)) {
                    paramsHtml += generateParamControl(rule.id, paramName, paramValue);
                }

                paramsHtml += `</div></div>`;
            }

            // 组合HTML
            ruleItem.innerHTML = ruleControlsHtml + paramsHtml;

            content.appendChild(ruleItem);
        });

        categoryElement.appendChild(header);
        categoryElement.appendChild(content);
        mainPanel.appendChild(categoryElement);

        // 添加展开/收起事件
        header.addEventListener('click', toggleRuleCategory);
    });

    // 添加规则开关事件监听器
    document.querySelectorAll('.rule-toggle').forEach(toggle => {
        toggle.addEventListener('change', updateRuleCounts);
    });

    updateRuleCounts();
}

// 生成参数控件
function generateParamControl(ruleId, paramName, paramValue) {
    const paramKey = `${ruleId}_${paramName}`;
    let controlHtml = '';

    // 转换参数名显示格式（驼峰转空格，首字母大写）
    const displayName = paramName
        .replace(/([A-Z])/g, ' $1')
        .replace(/^./, str => str.toUpperCase())
        .replace(/_/g, ' ');

    // 根据参数类型生成不同的控件
    const paramType = typeof paramValue;

    // 特殊参数处理
    if (paramName.includes('color') || paramName.endsWith('Color')) {
        // 颜色选择器
        controlHtml = `
                    <div class="param-control" style="display: flex; flex-direction: column; gap: 5px;">
                        <label for="${paramKey}" style="font-size: 12px; color: #666;">${displayName}</label>
                        <div style="display: flex; gap: 10px; align-items: center;">
                            <input 
                                type="color" 
                                id="${paramKey}_color" 
                                value="#${paramValue}" 
                                style="width: 50px; height: 30px; border: 1px solid #ddd; border-radius: 4px; cursor: pointer;">
                            <input 
                                type="text" 
                                id="${paramKey}" 
                                data-rule="${ruleId}" 
                                data-param="${paramName}" 
                                value="${paramValue}" 
                                style="flex: 1; padding: 5px; border: 1px solid #ddd; border-radius: 4px; font-size: 12px; text-transform: uppercase;">
                        </div>
                    </div>`;

        // 添加颜色选择器事件监听（在控件生成后添加）
        setTimeout(() => {
            const colorPicker = document.getElementById(`${paramKey}_color`);
            const colorInput = document.getElementById(`${paramKey}`);
            if (colorPicker && colorInput) {
                colorPicker.addEventListener('input', function () {
                    colorInput.value = this.value.substring(1); // 移除#号
                });
                colorInput.addEventListener('input', function () {
                    colorPicker.value = `#${this.value}`;
                });
            }
        }, 0);
    }
    // 数值类型 - 添加滑块控件
    else if (paramType === 'number') {
        controlHtml = `
                    <div class="param-control" style="display: flex; flex-direction: column; gap: 5px;">
                        <label for="${paramKey}" style="font-size: 12px; color: #666;">${displayName}</label>
                        <input 
                            type="range" 
                            id="${paramKey}_slider" 
                            data-rule="${ruleId}" 
                            data-param="${paramName}" 
                            min="${Math.max(0, paramValue - 10)}" 
                            max="${paramValue + 10}" 
                            step="0.1" 
                            value="${paramValue}" 
                            style="width: 100%; margin-bottom: 5px;">
                        <input 
                            type="number" 
                            id="${paramKey}" 
                            data-rule="${ruleId}" 
                            data-param="${paramName}" 
                            value="${paramValue}" 
                            step="0.1" 
                            style="padding: 5px; border: 1px solid #ddd; border-radius: 4px; font-size: 12px; width: 100px;">
                    </div>`;

        // 添加滑块和输入框的联动事件（在控件生成后添加）
        setTimeout(() => {
            const slider = document.getElementById(`${paramKey}_slider`);
            const input = document.getElementById(`${paramKey}`);
            if (slider && input) {
                slider.addEventListener('input', function () {
                    input.value = this.value;
                });
                input.addEventListener('input', function () {
                    slider.value = this.value;
                });
            }
        }, 0);
    }
    // 字符串类型
    else if (paramType === 'string') {
        // 特殊处理字体选择
        if (paramName.includes('font') || paramName.endsWith('Font')) {
            // 常见中文字体列表
            const chineseFonts = ['宋体', '黑体', '微软雅黑', '仿宋', '楷体', '新宋体', '华文楷体', '华文仿宋', '华文黑体', '华文细黑'];
            // 常见西文字体列表
            const westernFonts = ['Arial', 'Times New Roman', 'Calibri', 'Verdana', 'Helvetica', 'Georgia', 'Courier New', 'Tahoma'];

            let fonts = chineseFonts.concat(westernFonts);

            // 生成下拉选择框
            controlHtml = `
                        <div class="param-control" style="display: flex; flex-direction: column; gap: 5px;">
                            <label for="${paramKey}" style="font-size: 12px; color: #666;">${displayName}</label>
                            <select 
                                id="${paramKey}" 
                                data-rule="${ruleId}" 
                                data-param="${paramName}" 
                                style="padding: 5px; border: 1px solid #ddd; border-radius: 4px; font-size: 12px;">
                                ${fonts.map(font => `<option value="${font}" ${paramValue === font ? 'selected' : ''} style="font-family: ${font};">${font}</option>`).join('')}
                            </select>
                            <div style="margin-top: 5px; font-size: 14px; font-weight: bold;" id="${paramKey}_preview" style="font-family: ${paramValue};">字体预览：这是${paramValue}字体</div>
                        </div>`;

            // 添加字体预览事件监听（在控件生成后添加）
            setTimeout(() => {
                const select = document.getElementById(`${paramKey}`);
                const preview = document.getElementById(`${paramKey}_preview`);
                if (select && preview) {
                    select.addEventListener('change', function () {
                        preview.style.fontFamily = this.value;
                        preview.textContent = `字体预览：这是${this.value}字体`;
                    });
                }
            }, 0);
        }
        // 对齐方式选择
        else if (paramName.includes('align') || paramName.endsWith('Align')) {
            const alignOptions = [
                { value: 'left', text: '左对齐' },
                { value: 'center', text: '居中对齐' },
                { value: 'right', text: '右对齐' },
                { value: 'justify', text: '两端对齐' }
            ];

            controlHtml = `
                        <div class="param-control" style="display: flex; flex-direction: column; gap: 5px;">
                            <label for="${paramKey}" style="font-size: 12px; color: #666;">${displayName}</label>
                            <select 
                                id="${paramKey}" 
                                data-rule="${ruleId}" 
                                data-param="${paramName}" 
                                style="padding: 5px; border: 1px solid #ddd; border-radius: 4px; font-size: 12px;">
                                ${alignOptions.map(option => `<option value="${option.value}" ${paramValue === option.value ? 'selected' : ''}>${option.text}</option>`).join('')}
                            </select>
                        </div>`;
        }
        // 垂直对齐方式选择
        else if (paramName.includes('vertical_alignment') || paramName.endsWith('VerticalAlignment')) {
            const alignOptions = [
                { value: 'top', text: '顶端对齐' },
                { value: 'center', text: '垂直居中' },
                { value: 'bottom', text: '底端对齐' }
            ];

            controlHtml = `
                        <div class="param-control" style="display: flex; flex-direction: column; gap: 5px;">
                            <label for="${paramKey}" style="font-size: 12px; color: #666;">${displayName}</label>
                            <select 
                                id="${paramKey}" 
                                data-rule="${ruleId}" 
                                data-param="${paramName}" 
                                style="padding: 5px; border: 1px solid #ddd; border-radius: 4px; font-size: 12px;">
                                ${alignOptions.map(option => `<option value="${option.value}" ${paramValue === option.value ? 'selected' : ''}>${option.text}</option>`).join('')}
                            </select>
                        </div>`;
        }
        // 普通字符串输入
        else {
            controlHtml = `
                        <div class="param-control" style="display: flex; flex-direction: column; gap: 5px;">
                            <label for="${paramKey}" style="font-size: 12px; color: #666;">${displayName}</label>
                            <input 
                                type="text" 
                                id="${paramKey}" 
                                data-rule="${ruleId}" 
                                data-param="${paramName}" 
                                value="${paramValue}" 
                                style="padding: 5px; border: 1px solid #ddd; border-radius: 4px; font-size: 12px;">
                        </div>`;
        }
    }
    // 布尔类型 - 开关控件
    else if (paramType === 'boolean') {
        controlHtml = `
                    <div class="param-control" style="display: flex; align-items: center; gap: 10px;">
                        <label class="toggle-switch">
                            <input 
                                type="checkbox" 
                                id="${paramKey}" 
                                data-rule="${ruleId}" 
                                data-param="${paramName}" 
                                ${paramValue ? 'checked' : ''}>
                            <span class="toggle-slider"></span>
                        </label>
                        <label for="${paramKey}" style="font-size: 12px; color: #666; margin: 0;">${displayName}</label>
                    </div>`;
    }
    // 其他类型
    else {
        controlHtml = `
                    <div class="param-control" style="display: flex; flex-direction: column; gap: 5px;">
                        <label for="${paramKey}" style="font-size: 12px; color: #666;">${displayName}</label>
                        <input 
                            type="text" 
                            id="${paramKey}" 
                            data-rule="${ruleId}" 
                            data-param="${paramName}" 
                            value="${JSON.stringify(paramValue)}" 
                            style="padding: 5px; border: 1px solid #ddd; border-radius: 4px; font-size: 12px;">
                    </div>`;
    }

    return controlHtml;
}

// 显示进度条
function showProgressBar() {
    const progressBar = document.querySelector('.progress-bar');
    const progressFill = document.querySelector('.progress-fill');
    if (progressBar) {
        progressBar.style.display = 'block';
        progressFill.style.width = '0%';

        // 模拟进度
        let progress = 0;
        const interval = setInterval(() => {
            progress += 10;
            if (progress <= 90) {
                progressFill.style.width = progress + '%';
            } else {
                clearInterval(interval);
            }
        }, 200);
    }
}

// 隐藏进度条
function hideProgressBar() {
    const progressBar = document.querySelector('.progress-bar');
    const progressFill = document.querySelector('.progress-fill');
    if (progressBar) {
        progressFill.style.width = '100%';
        setTimeout(() => {
            progressBar.style.display = 'none';
        }, 300);
    }
}

// 记录开发者日志
function logDev(message) {
    const devLog = document.getElementById('devLog');
    const timestamp = new Date().toLocaleTimeString();
    devLog.textContent += `[${timestamp}] ${message}\n`;
    devLog.scrollTop = devLog.scrollHeight;
}

// 预设编辑器相关函数
function openPresetEditor(presetId = null) {
    const modal = document.getElementById('presetModal');
    const modalTitle = document.getElementById('modalTitle');
    const presetNameInput = document.getElementById('presetName');
    const presetDescriptionInput = document.getElementById('presetDescription');
    const ruleSelection = document.getElementById('ruleSelection');
    const savePresetBtn = document.getElementById('savePreset');

    // 清空表单
    presetNameInput.value = '';
    presetDescriptionInput.value = '';
    ruleSelection.innerHTML = '';

    // 设置模态框标题
    if (presetId) {
        modalTitle.textContent = '编辑预设';
        savePresetBtn.onclick = () => savePreset(presetId);

        // 加载预设数据
        const preset = presets[presetId];
        if (preset) {
            presetNameInput.value = preset.name;
            presetDescriptionInput.value = preset.description;
        }
    } else {
        modalTitle.textContent = '新建预设';
        savePresetBtn.onclick = saveNewPreset;
    }

    // 生成规则选择列表
    generateRuleSelection(presetId);

    // 显示模态框
    modal.style.display = 'flex'; // 使用 flex 居中
    document.body.classList.add('modal-open'); // 防止背景滚动

    // 聚焦到第一个输入框，提升效率
    setTimeout(() => {
        document.getElementById('presetName').focus();
    }, 100);

    logDev(`打开预设编辑器: ${presetId || '新建'}`);
}

function closePresetEditor() {
    const modal = document.getElementById('presetModal');
    modal.style.display = 'none';
    document.body.classList.remove('modal-open'); // 恢复背景滚动
    logDev('关闭预设编辑器');
}

function generateRuleSelection(presetId = null) {
    const ruleSelection = document.getElementById('ruleSelection');

    // 清空现有内容
    ruleSelection.innerHTML = '';

    // 添加标题
    const title = document.createElement('h3');
    title.textContent = '选择规则';
    ruleSelection.appendChild(title);

    // 按类别分组规则
    const rulesByCategory = {};
    backendRules.forEach(rule => {
        let category = '其他规则';
        if (rule.id.includes('Font') || rule.id.includes('font')) {
            category = '字体规则';
        } else if (rule.id.includes('Paragraph') || rule.id.includes('paragraph') || rule.id.includes('Title') || rule.id.includes('title') || rule.id.includes('List') || rule.id.includes('list')) {
            category = '段落规则';
        } else if (rule.id.includes('Table') || rule.id.includes('table')) {
            category = '表格规则';
        } else if (rule.id.includes('Page') || rule.id.includes('page')) {
            category = '页面规则';
        }

        if (!rulesByCategory[category]) {
            rulesByCategory[category] = [];
        }
        rulesByCategory[category].push(rule);
    });

    // 生成规则复选框
    Object.entries(rulesByCategory).forEach(([categoryName, rules]) => {
        const categoryDiv = document.createElement('div');
        categoryDiv.style.marginBottom = '15px';

        const categoryTitle = document.createElement('h4');
        categoryTitle.style.fontSize = '14px';
        categoryTitle.style.marginBottom = '10px';
        categoryTitle.style.color = '#666666';
        categoryTitle.textContent = categoryName;
        categoryDiv.appendChild(categoryTitle);

        rules.forEach(rule => {
            const ruleCheckbox = document.createElement('div');
            ruleCheckbox.className = 'rule-checkbox';

            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.id = `rule_${rule.id}`;
            checkbox.value = rule.id;

            // 检查预设是否启用了该规则
            if (presetId && presets[presetId] && presets[presetId].rules[rule.id] && presets[presetId].rules[rule.id].enabled) {
                checkbox.checked = true;
            }

            const label = document.createElement('label');
            label.htmlFor = `rule_${rule.id}`;
            label.textContent = `${rule.name} - ${rule.description}`;
            label.style.fontSize = '12px';

            ruleCheckbox.appendChild(checkbox);
            ruleCheckbox.appendChild(label);
            categoryDiv.appendChild(ruleCheckbox);
        });

        ruleSelection.appendChild(categoryDiv);
    });
}

async function saveNewPreset() {
    const presetName = document.getElementById('presetName').value.trim();
    const presetDescription = document.getElementById('presetDescription').value.trim();

    if (!presetName) {
        alert('请输入预设名称');
        return;
    }

    // 生成唯一ID
    const presetId = presetName.toLowerCase().replace(/\s+/g, '_');

    // 检查ID是否已存在
    if (presets[presetId]) {
        alert('预设名称已存在');
        return;
    }

    // 获取选中的规则
    const selectedRules = getSelectedRulesFromEditor();

    // 创建新预设
    const newPreset = {
        name: presetName,
        description: presetDescription,
        rules: {}
    };

    // 添加规则
    selectedRules.forEach(ruleId => {
        newPreset.rules[ruleId] = { enabled: true };
    });

    // 保存到后端
    try {
        await window.electronAPI.savePreset(presetId, newPreset);

        // 更新本地预设数据
        presets[presetId] = newPreset;

        // 更新预设UI
        addPresetCard(presetId, newPreset);

        closePresetEditor();
        logDev(`新建预设成功: ${presetId}`);
    } catch (error) {
        console.error('保存预设失败:', error);
        alert('保存预设失败: ' + error.message);
        logDev(`新建预设失败: ${error.message}`);
    }
}

async function savePreset(presetId) {
    const presetName = document.getElementById('presetName').value.trim();
    const presetDescription = document.getElementById('presetDescription').value.trim();

    if (!presetName) {
        alert('请输入预设名称');
        return;
    }

    // 获取选中的规则
    const selectedRules = getSelectedRulesFromEditor();

    // 更新预设
    const updatedPreset = {
        name: presetName,
        description: presetDescription,
        rules: {}
    };

    // 添加规则
    selectedRules.forEach(ruleId => {
        updatedPreset.rules[ruleId] = { enabled: true };
    });

    // 保存到后端
    try {
        await window.electronAPI.savePreset(presetId, updatedPreset);

        // 更新本地预设数据
        presets[presetId] = updatedPreset;

        // 更新预设UI
        updatePresetCard(presetId, updatedPreset);

        closePresetEditor();
        logDev(`更新预设成功: ${presetId}`);
    } catch (error) {
        console.error('更新预设失败:', error);
        alert('更新预设失败: ' + error.message);
        logDev(`更新预设失败: ${error.message}`);
    }
}

function getSelectedRulesFromEditor() {
    const selectedRules = [];
    const checkboxes = document.querySelectorAll('#ruleSelection input[type="checkbox"]:checked');

    checkboxes.forEach(checkbox => {
        selectedRules.push(checkbox.value);
    });

    return selectedRules;
}

function addPresetCard(presetId, preset) {
    const presetCards = document.querySelector('.preset-cards');

    const presetCard = document.createElement('div');
    presetCard.className = 'preset-card';
    presetCard.dataset.preset = presetId;

    presetCard.innerHTML = `
                <div class="preset-header">
                    <h4>${preset.name}</h4>
                    <div class="preset-actions">
                        <button class="btn btn-sm btn-secondary" onclick="openPresetEditor('${presetId}')">编辑</button>
                        <button class="btn btn-sm btn-danger" onclick="confirmDeletePreset('${presetId}')">删除</button>
                    </div>
                </div>
                <p>${preset.description}</p>
            `;

    // 添加点击事件
    presetCard.addEventListener('click', function () {
        selectPreset(this.dataset.preset);
    });

    presetCards.appendChild(presetCard);
    logDev(`添加预设卡片: ${presetId}`);
}

function updatePresetCard(presetId, preset) {
    const presetCard = document.querySelector(`[data-preset="${presetId}"]`);
    if (presetCard) {
        presetCard.innerHTML = `
                    <div class="preset-header">
                        <h4>${preset.name}</h4>
                        <div class="preset-actions">
                            <button class="btn btn-sm btn-secondary" onclick="openPresetEditor('${presetId}')">编辑</button>
                            <button class="btn btn-sm btn-danger" onclick="confirmDeletePreset('${presetId}')">删除</button>
                        </div>
                    </div>
                    <p>${preset.description}</p>
                `;

        // 重新添加点击事件
        presetCard.addEventListener('click', function () {
            selectPreset(this.dataset.preset);
        });

        logDev(`更新预设卡片: ${presetId}`);
    }
}

function confirmDeletePreset(presetId) {
    if (confirm(`确定要删除预设 "${presets[presetId]?.name}" 吗？`)) {
        deletePreset(presetId);
    }
}

async function deletePreset(presetId) {
    // 不允许删除默认预设
    if (presetId === 'default') {
        alert('默认预设不能删除');
        return;
    }

    try {
        await window.electronAPI.deletePreset(presetId);

        // 更新本地预设数据
        delete presets[presetId];

        // 更新预设UI
        removePresetCard(presetId);

        logDev(`删除预设成功: ${presetId}`);
    } catch (error) {
        console.error('删除预设失败:', error);
        alert('删除预设失败: ' + error.message);
        logDev(`删除预设失败: ${error.message}`);
    }
}

function removePresetCard(presetId) {
    const presetCard = document.querySelector(`[data-preset="${presetId}"]`);
    if (presetCard) {
        presetCard.remove();
        logDev(`移除预设卡片: ${presetId}`);
    }
}

// 点击模态框外部关闭
window.onclick = function (event) {
    const modal = document.getElementById('presetModal');
    if (event.target == modal) {
        closePresetEditor();
    }
}
