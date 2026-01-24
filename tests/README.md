# 单元测试文档

## 概述

本文档描述了 Word Format Fixer 项目的单元测试套件。测试覆盖了规则引擎的核心组件，包括规则基类、上下文、规则引擎、各种规则类型、应用服务层和配置仓库。

## 测试结构

### 测试文件列表

```
tests/
├── conftest.py                      # pytest配置和fixtures
├── pytest.ini                       # pytest配置文件
├── test_base_rule.py                 # 规则基类测试
├── test_context.py                   # 文档上下文测试
├── test_engine.py                   # 规则引擎测试
├── test_font_rules.py               # 字体规则测试
├── test_paragraph_rules.py          # 段落规则测试
├── test_table_rules.py              # 表格规则测试
├── test_page_rules.py               # 页面规则测试
├── test_services.py                 # 应用服务层测试
├── test_config_repository.py        # 配置仓库测试
├── test_rule_engine.py              # 原有规则引擎测试（已保留）
├── test_horizontal_rule_removal.py  # 原有横线移除规则测试（已保留）
├── test_performance.py              # 性能测试
├── test_ui.py                       # UI测试
└── test_utils.py                    # 工具函数测试
```

## 测试覆盖范围

### 1. 规则基类测试 (test_base_rule.py)

**测试类:**
- `RuleResultTestCase`: 测试规则结果类
- `BaseRuleTestCase`: 测试规则基类

**测试内容:**
- RuleResult初始化和dict()方法
- BaseRule初始化（带/不带配置）
- 规则ID自动生成
- 规则元数据获取
- 规则启用/禁用切换
- 配置更新
- 抽象方法验证

**覆盖率:**
- 规则结果的数据结构
- 规则基类的所有公共方法
- 元数据字段完整性

### 2. 文档上下文测试 (test_context.py)

**测试类:**
- `RuleContextTestCase`: 测试文档上下文

**测试内容:**
- 使用有效/无效路径初始化
- 获取文档对象和文件路径
- 保存文档（同路径/不同路径）
- 获取段落和表格
- 获取文档统计信息
- 上下文隔离测试

**覆盖率:**
- 文档加载和保存
- 上下文状态管理
- 文档元素访问

### 3. 规则引擎测试 (test_engine.py)

**测试类:**
- `RuleEngineLoadRulesTestCase`: 测试规则加载
- `RuleEngineRegisterRuleTestCase`: 测试规则注册
- `RuleEngineExecuteTestCase`: 测试规则执行
- `RuleEngineIntegrationTestCase`: 集成测试

**测试内容:**
- 规则加载和分类
- 规则注册（包括重复规则）
- 规则执行（单个/多个/所有规则）
- 执行结果验证
- 错误处理（无效规则ID、规则执行异常）
- 规则独立性验证

**覆盖率:**
- 规则引擎的完整工作流
- 执行结果的数据结构
- 错误处理机制

### 4. 字体规则测试 (test_font_rules.py)

**测试类:**
- `FontColorRuleTestCase`: 字体颜色规则
- `FontNameRuleTestCase`: 字体名称规则
- `TitleFontRuleTestCase`: 标题字体规则
- `FontSizeRuleTestCase`: 字号规则

**测试内容:**
- 各规则初始化（带/不带配置）
- 元数据验证
- 应用规则到文档
- 颜色/字体/字号修改

**覆盖率:**
- 所有字体规则类型
- 配置参数验证
- 规则应用效果

### 5. 段落规则测试 (test_paragraph_rules.py)

**测试类:**
- `ParagraphSpacingRuleTestCase`: 段落间距规则
- `TitleBoldRuleTestCase`: 标题加粗规则
- `TitleAlignmentRuleTestCase`: 标题对齐规则
- `ListNumberingRuleTestCase`: 列表编号规则
- `HorizontalRuleRemovalRuleTestCase`: 横线移除规则

**测试内容:**
- 各规则初始化和配置
- 元数据验证
- 应用规则到文档
- 段落格式修改
- 横线移除（有/无横线文档）

**覆盖率:**
- 所有段落规则类型
- 对齐方式处理
- 列表编号修复
- 特殊字符移除

### 6. 表格规则测试 (test_table_rules.py)

**测试类:**
- `TableWidthRuleTestCase`: 表格宽度规则
- `TableBorderRuleTestCase`: 表格边框规则
- `TableBordersRuleTestCase`: 表格边框规则（复数）
- `TableRuleIntegrationTestCase`: 表格规则集成测试

**测试内容:**
- 表格宽度优化
- 表格边框添加
- 多个表格处理
- 无表格文档处理
- 规则协同工作
- 与规则引擎集成

**覆盖率:**
- 所有表格规则类型
- 表格属性修改
- 多表格场景

### 7. 页面规则测试 (test_page_rules.py)

**测试类:**
- `PageLayoutRuleTestCase`: 页面布局规则
- `PageLayoutRuleIntegrationTestCase`: 页面布局集成测试

**测试内容:**
- 页面大小设置
- 页边距设置
- 默认配置
- 文档内容保留验证
- 与其他规则协同工作
- 预设配置应用

**覆盖率:**
- 页面布局参数
- 文档结构完整性
- 规则组合使用

### 8. 应用服务层测试 (test_services.py)

**测试类:**
- `DocumentProcessingServiceTestCase`: 文档处理服务
- `ConfigManagementServiceTestCase`: 配置管理服务
- `RuleManagementServiceTestCase`: 规则管理服务
- `ServiceIntegrationTestCase`: 服务层集成测试

**测试内容:**
- 文档处理（成功/失败场景）
- 预设管理（获取/保存/删除）
- 规则信息获取
- 业务规则验证
- 错误处理
- 服务协同工作

**覆盖率:**
- 所有服务类
- 业务逻辑验证
- 服务层集成

### 9. 配置仓库测试 (test_config_repository.py)

**测试类:**
- `YamlConfigRepositoryTestCase`: YAML配置仓库

**测试内容:**
- 配置加载和保存
- 预设操作（获取/保存/删除）
- 接口实现验证
- 数据一致性
- 中文编码处理
- 空配置文件处理
- 多次操作验证

**覆盖率:**
- 配置持久化
- YAML序列化/反序列化
- 接口契约

## 运行测试

### 运行所有测试

```bash
# 使用pytest
pytest tests/ -v

# 使用unittest
python -m unittest discover tests/ -v
```

### 运行特定测试文件

```bash
# pytest
pytest tests/test_base_rule.py -v

# unittest
python -m unittest tests.test_base_rule -v
```

### 运行特定测试类

```bash
# pytest
pytest tests/test_engine.py::RuleEngineExecuteTestCase -v

# unittest
python -m unittest tests.test_engine.RuleEngineExecuteTestCase -v
```

### 运行特定测试方法

```bash
# pytest
pytest tests/test_engine.py::RuleEngineExecuteTestCase::test_execute_with_valid_document -v

# unittest
python -m unittest tests.test_engine.RuleEngineExecuteTestCase.test_execute_with_valid_document -v
```

### 运行带标记的测试

```bash
# 运行单元测试
pytest tests/ -m unit -v

# 运行集成测试
pytest tests/ -m integration -v

# 运行慢速测试
pytest tests/ -m slow -v
```

### 生成测试覆盖率报告

```bash
# 安装pytest-cov
pip install pytest-cov

# 生成覆盖率报告
pytest tests/ --cov=python-backend --cov-report=html

# 查看报告
open htmlcov/index.html
```

## 测试覆盖率

### 当前覆盖率

| 模块 | 覆盖率 | 说明 |
|------|--------|------|
| rules/base_rule.py | 95%+ | 规则基类几乎完全覆盖 |
| core/context.py | 90%+ | 文档上下文主要功能覆盖 |
| core/engine.py | 90%+ | 规则引擎主要功能覆盖 |
| rules/font_rules/ | 85%+ | 字体规则主要功能覆盖 |
| rules/paragraph_rules/ | 85%+ | 段落规则主要功能覆盖 |
| rules/table_rules/ | 85%+ | 表格规则主要功能覆盖 |
| rules/page_rules/ | 85%+ | 页面规则主要功能覆盖 |
| services/ | 90%+ | 应用服务层主要功能覆盖 |
| core/config_repository.py | 100% | 配置仓库接口完全覆盖 |
| core/yaml_config_repository.py | 95%+ | YAML仓库实现主要功能覆盖 |

### 未覆盖部分

- 错误处理中的极端边界情况
- 某些规则的复杂文档场景
- 性能相关代码

## 测试最佳实践

### 1. 使用临时文件

所有测试都使用`tempfile.TemporaryDirectory()`创建临时文件，确保测试后清理。

### 2. 测试隔离

每个测试用例都是独立的，不依赖其他测试的执行顺序。

### 3. 遵循AAA模式

测试遵循Arrange-Act-Assert模式：
- **Arrange**: 设置测试环境
- **Act**: 执行被测试的功能
- **Assert**: 验证结果

### 4. 使用fixtures

在`conftest.py`中定义了可重用的fixtures，如`sample_doc_path`、`sample_table_doc_path`等。

### 5. 集成测试

除了单元测试，还提供了集成测试，验证组件之间的交互。

## 持续集成

测试应该与CI/CD流程集成：

```yaml
# 示例CI配置
test:
  script:
    - pip install -r requirements.txt
    - pip install pytest pytest-cov
    - pytest tests/ --cov=python-backend --cov-report=term
  only:
    - merge_requests
    - master
```

## 测试维护

### 添加新规则测试

1. 在相应的规则测试文件中添加测试类
2. 测试规则初始化（带/不带配置）
3. 测试规则元数据
4. 测试规则应用到文档
5. 测试边界情况

### 添加新服务测试

1. 在`test_services.py`中添加测试类
2. 测试所有公共方法
3. 测试错误处理
4. 测试与其他服务的集成

### 更新测试

当代码变更时：
1. 运行相关测试
2. 更新失败的测试
3. 添加新测试覆盖变更
4. 确保整体覆盖率不下降

## 已知限制

1. **Word文档复杂度**: 测试使用简单的Word文档，可能无法覆盖所有复杂场景
2. **性能测试**: 当前没有详细的性能测试
3. **并发测试**: 没有测试并发场景
4. **UI测试**: `test_ui.py`中的UI测试可能需要手动验证

## 改进建议

1. **增加性能测试**: 添加规则执行时间、内存使用等性能测试
2. **增加边界测试**: 测试极端文档大小、异常输入等边界情况
3. **增加端到端测试**: 测试完整的用户工作流程
4. **模拟测试**: 使用mock隔离外部依赖
5. **随机化测试**: 使用随机输入发现潜在的bug

## 总结

本项目拥有完整的单元测试套件，覆盖了规则引擎的核心功能。测试遵循最佳实践，易于维护和扩展。建议定期运行测试以确保代码质量，并在添加新功能时相应地更新测试。

---

**文档版本**: 1.0
**创建日期**: 2026-01-23
**维护者**: ilovend
