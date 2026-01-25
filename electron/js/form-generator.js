/**
 * 动态表单生成器
 * 
 * 根据后端规则的 param_schema 自动生成 UI 控件：
 * - range: 滑块 + 数字输入
 * - color: 颜色选择器
 * - font: 字体下拉框
 * - boolean: 开关按钮
 * - enum: 下拉选择框
 * - string: 文本输入框
 * - number/integer: 数字输入框
 */

/**
 * 根据 param_schema 生成参数控件 HTML
 * @param {string} ruleId - 规则ID
 * @param {Object} paramSchema - 参数 schema 定义
 * @param {any} currentValue - 当前参数值
 * @returns {string} HTML 字符串
 */
function generateParamControlFromSchema(ruleId, paramSchema, currentValue) {
    const paramKey = `${ruleId}_${paramSchema.name}`;
    const displayName = paramSchema.display_name || paramSchema.name;
    const description = paramSchema.description || '';
    const value = currentValue !== undefined ? currentValue : paramSchema.default;

    let controlHtml = '';

    switch (paramSchema.param_type) {
        case 'range':
            controlHtml = generateRangeControl(paramKey, ruleId, paramSchema, value);
            break;
        case 'color':
            controlHtml = generateColorControl(paramKey, ruleId, paramSchema, value);
            break;
        case 'font':
            controlHtml = generateFontControl(paramKey, ruleId, paramSchema, value);
            break;
        case 'boolean':
            controlHtml = generateBooleanControl(paramKey, ruleId, paramSchema, value);
            break;
        case 'enum':
            controlHtml = generateEnumControl(paramKey, ruleId, paramSchema, value);
            break;
        case 'number':
        case 'integer':
            controlHtml = generateNumberControl(paramKey, ruleId, paramSchema, value);
            break;
        case 'string':
        default:
            controlHtml = generateStringControl(paramKey, ruleId, paramSchema, value);
            break;
    }

    return controlHtml;
}

/**
 * 生成滑块控件
 */
function generateRangeControl(paramKey, ruleId, schema, value) {
    const min = schema.min_value !== undefined ? schema.min_value : 0;
    const max = schema.max_value !== undefined ? schema.max_value : 100;
    const step = schema.step !== undefined ? schema.step : 1;
    const unit = schema.unit || '';

    const html = `
        <div class="param-control">
            <label for="${paramKey}">${schema.display_name}</label>
            <div class="param-range-container">
                <input 
                    type="range" 
                    id="${paramKey}_slider" 
                    min="${min}" 
                    max="${max}" 
                    step="${step}" 
                    value="${value}"
                    class="param-slider">
                <div class="param-value-display">
                    <input 
                        type="number" 
                        id="${paramKey}" 
                        data-rule="${ruleId}" 
                        data-param="${schema.name}" 
                        value="${value}" 
                        min="${min}"
                        max="${max}"
                        step="${step}"
                        class="param-number-input">
                    <span class="param-unit">${unit}</span>
                </div>
            </div>
            ${schema.description ? `<small class="param-hint">${schema.description}</small>` : ''}
        </div>`;

    // 设置滑块和输入框联动
    setTimeout(() => {
        const slider = document.getElementById(`${paramKey}_slider`);
        const input = document.getElementById(`${paramKey}`);
        if (slider && input) {
            slider.addEventListener('input', () => input.value = slider.value);
            input.addEventListener('input', () => slider.value = input.value);
        }
    }, 0);

    return html;
}

/**
 * 生成颜色选择器
 */
function generateColorControl(paramKey, ruleId, schema, value) {
    // 确保颜色值以 # 开头
    let colorValue = value || '#000000';
    if (!colorValue.startsWith('#')) {
        colorValue = '#' + colorValue;
    }

    const html = `
        <div class="param-control">
            <label for="${paramKey}">${schema.display_name}</label>
            <div class="param-color-container">
                <input 
                    type="color" 
                    id="${paramKey}_picker" 
                    value="${colorValue}"
                    class="param-color-picker">
                <input 
                    type="text" 
                    id="${paramKey}" 
                    data-rule="${ruleId}" 
                    data-param="${schema.name}" 
                    value="${colorValue}"
                    class="param-color-input">
            </div>
            ${schema.description ? `<small class="param-hint">${schema.description}</small>` : ''}
        </div>`;

    // 设置颜色选择器联动
    setTimeout(() => {
        const picker = document.getElementById(`${paramKey}_picker`);
        const input = document.getElementById(`${paramKey}`);
        if (picker && input) {
            picker.addEventListener('input', () => input.value = picker.value);
            input.addEventListener('input', () => {
                if (/^#[0-9A-Fa-f]{6}$/.test(input.value)) {
                    picker.value = input.value;
                }
            });
        }
    }, 0);

    return html;
}

/**
 * 生成字体选择器
 */
function generateFontControl(paramKey, ruleId, schema, value) {
    // 使用 schema 提供的选项或默认字体列表
    const options = schema.options || [
        { value: '宋体', label: '宋体' },
        { value: '黑体', label: '黑体' },
        { value: '楷体', label: '楷体' },
        { value: '仿宋', label: '仿宋' },
        { value: '微软雅黑', label: '微软雅黑' },
        { value: 'Arial', label: 'Arial' },
        { value: 'Times New Roman', label: 'Times New Roman' }
    ];

    const optionsHtml = options.map(opt =>
        `<option value="${opt.value}" ${value === opt.value ? 'selected' : ''} style="font-family: ${opt.value};">${opt.label}</option>`
    ).join('');

    const html = `
        <div class="param-control">
            <label for="${paramKey}">${schema.display_name}</label>
            <select 
                id="${paramKey}" 
                data-rule="${ruleId}" 
                data-param="${schema.name}"
                class="param-select param-font-select">
                ${optionsHtml}
            </select>
            <div class="param-font-preview" id="${paramKey}_preview" style="font-family: ${value};">
                字体预览 Font Preview
            </div>
            ${schema.description ? `<small class="param-hint">${schema.description}</small>` : ''}
        </div>`;

    // 设置字体预览联动
    setTimeout(() => {
        const select = document.getElementById(`${paramKey}`);
        const preview = document.getElementById(`${paramKey}_preview`);
        if (select && preview) {
            select.addEventListener('change', () => {
                preview.style.fontFamily = select.value;
            });
        }
    }, 0);

    return html;
}

/**
 * 生成布尔开关
 */
function generateBooleanControl(paramKey, ruleId, schema, value) {
    const checked = value ? 'checked' : '';

    return `
        <div class="param-control param-control-boolean">
            <div class="param-boolean-container">
                <label class="param-toggle-switch">
                    <input 
                        type="checkbox" 
                        id="${paramKey}" 
                        data-rule="${ruleId}" 
                        data-param="${schema.name}"
                        ${checked}>
                    <span class="param-toggle-slider"></span>
                </label>
                <span class="param-boolean-label">${schema.display_name}</span>
            </div>
            ${schema.description ? `<small class="param-hint">${schema.description}</small>` : ''}
        </div>`;
}

/**
 * 生成枚举下拉框
 */
function generateEnumControl(paramKey, ruleId, schema, value) {
    const options = schema.options || [];
    const optionsHtml = options.map(opt =>
        `<option value="${opt.value}" ${value === opt.value ? 'selected' : ''}>${opt.label}</option>`
    ).join('');

    return `
        <div class="param-control">
            <label for="${paramKey}">${schema.display_name}</label>
            <select 
                id="${paramKey}" 
                data-rule="${ruleId}" 
                data-param="${schema.name}"
                class="param-select">
                ${optionsHtml}
            </select>
            ${schema.description ? `<small class="param-hint">${schema.description}</small>` : ''}
        </div>`;
}

/**
 * 生成数字输入框
 */
function generateNumberControl(paramKey, ruleId, schema, value) {
    const min = schema.min_value !== undefined ? `min="${schema.min_value}"` : '';
    const max = schema.max_value !== undefined ? `max="${schema.max_value}"` : '';
    const step = schema.step !== undefined ? `step="${schema.step}"` : '';
    const unit = schema.unit || '';

    return `
        <div class="param-control">
            <label for="${paramKey}">${schema.display_name}</label>
            <div class="param-number-container">
                <input 
                    type="number" 
                    id="${paramKey}" 
                    data-rule="${ruleId}" 
                    data-param="${schema.name}" 
                    value="${value}"
                    ${min} ${max} ${step}
                    class="param-number-input">
                ${unit ? `<span class="param-unit">${unit}</span>` : ''}
            </div>
            ${schema.description ? `<small class="param-hint">${schema.description}</small>` : ''}
        </div>`;
}

/**
 * 生成文本输入框
 */
function generateStringControl(paramKey, ruleId, schema, value) {
    const placeholder = schema.placeholder || '';

    return `
        <div class="param-control">
            <label for="${paramKey}">${schema.display_name}</label>
            <input 
                type="text" 
                id="${paramKey}" 
                data-rule="${ruleId}" 
                data-param="${schema.name}" 
                value="${value || ''}"
                placeholder="${placeholder}"
                class="param-text-input">
            ${schema.description ? `<small class="param-hint">${schema.description}</small>` : ''}
        </div>`;
}

/**
 * 为规则生成参数配置面板
 * @param {Object} rule - 规则对象（包含 param_schema 和 params）
 * @returns {string} 参数配置面板 HTML
 */
function generateRuleParamsPanel(rule) {
    // 优先使用 param_schema，如果不存在则回退到 params
    const hasSchema = rule.param_schema && rule.param_schema.length > 0;
    const hasParams = rule.params && Object.keys(rule.params).length > 0;

    if (!hasSchema && !hasParams) {
        return ''; // 没有参数
    }

    let paramsHtml = `
        <div class="rule-params">
            <h5>参数配置</h5>
            <div class="params-grid">`;

    if (hasSchema) {
        // 使用 param_schema 生成控件（推荐方式）
        rule.param_schema.forEach(schema => {
            const currentValue = rule.params ? rule.params[schema.name] : schema.default;
            paramsHtml += generateParamControlFromSchema(rule.id, schema, currentValue);
        });
    } else {
        // 回退：使用旧的推断方式
        for (const [paramName, paramValue] of Object.entries(rule.params)) {
            paramsHtml += generateParamControlLegacy(rule.id, paramName, paramValue);
        }
    }

    paramsHtml += '</div></div>';
    return paramsHtml;
}

/**
 * 旧版参数控件生成（兼容没有 schema 的规则）
 */
function generateParamControlLegacy(ruleId, paramName, paramValue) {
    const paramKey = `${ruleId}_${paramName}`;
    const displayName = paramName
        .replace(/([A-Z])/g, ' $1')
        .replace(/^./, str => str.toUpperCase())
        .replace(/_/g, ' ');

    const paramType = typeof paramValue;

    // 颜色
    if (paramName.includes('color') || paramName.endsWith('Color')) {
        return `
            <div class="param-control">
                <label for="${paramKey}">${displayName}</label>
                <div class="param-color-container">
                    <input type="color" id="${paramKey}_picker" value="#${paramValue}">
                    <input type="text" id="${paramKey}" data-rule="${ruleId}" data-param="${paramName}" value="${paramValue}">
                </div>
            </div>`;
    }

    // 数字
    if (paramType === 'number') {
        return `
            <div class="param-control">
                <label for="${paramKey}">${displayName}</label>
                <input type="number" id="${paramKey}" data-rule="${ruleId}" data-param="${paramName}" value="${paramValue}" step="0.1">
            </div>`;
    }

    // 布尔
    if (paramType === 'boolean') {
        return `
            <div class="param-control param-control-boolean">
                <label class="param-toggle-switch">
                    <input type="checkbox" id="${paramKey}" data-rule="${ruleId}" data-param="${paramName}" ${paramValue ? 'checked' : ''}>
                    <span class="param-toggle-slider"></span>
                </label>
                <span>${displayName}</span>
            </div>`;
    }

    // 字符串
    return `
        <div class="param-control">
            <label for="${paramKey}">${displayName}</label>
            <input type="text" id="${paramKey}" data-rule="${ruleId}" data-param="${paramName}" value="${paramValue}">
        </div>`;
}

// 导出函数供全局使用
if (typeof window !== 'undefined') {
    window.FormGenerator = {
        generateParamControlFromSchema,
        generateRuleParamsPanel,
        generateParamControlLegacy
    };
}
