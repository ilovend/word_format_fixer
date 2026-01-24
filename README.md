# Word文档格式修复工具

一个用于修复Markdown转换为Word文档后格式问题的Python工具，采用规则引擎架构，支持多种预设配置和自定义规则。

## 文档目录

### 1. 架构与设计
- [架构方案](docs/architecture/01-架构方案.md) - 项目的整体架构设计
- [架构问题分析](docs/analysis/架构问题分析.md) - 项目中存在的逻辑混杂和架构问题
- [架构改进对比](docs/refactoring/架构改进对比.md) - 重构前后的架构对比

### 2. 重构与优化
- [重构分析](docs/refactoring/重构分析.md) - CodeBuddy重构方案的正确性分析
- [重构总结](docs/refactoring/重构总结.md) - 重构工作的总结
- [构建错误分析](docs/refactoring/构建错误分析.md) - 构建过程中遇到的错误及解决方案
- [重构最佳实践](docs/refactoring/重构最佳实践.md) - 重构过程中遵循的最佳实践

### 3. 功能与规范
- [功能与UI设计规范](docs/specifications/功能与UI设计规范.md) - 系统功能和UI设计规范
- [UI改进提案](docs/specifications/UI改进提案.md) - UI改进建议

### 4. 开发与迭代
- [迭代计划](docs/development/迭代计划.md) - 项目迭代计划
- [开发日志](docs/development/开发日志.md) - 开发过程中的记录
- [项目总结](docs/development/项目总结.md) - 项目的总结报告

### 5. 发布与变更
- [发布说明](docs/changelogs/发布说明.md) - 版本发布说明

## 架构演进

### 原始架构
- 单体应用设计
- 逻辑混杂，违反单一职责原则
- 规则依赖不清晰
- 配置管理混乱

### 重构后架构
- 分层架构设计
- 规则引擎模式
- 清晰的依赖关系
- 模块化设计

### 核心架构组件
1. **Presentation Layer** - Electron UI界面
2. **IPC Layer** - 前后端通信
3. **Application Service Layer** - 业务逻辑
4. **Domain Layer** - 规则引擎和核心逻辑
5. **Infrastructure Layer** - 配置管理和持久化

## 技术栈

- **前端**：Electron, HTML, CSS, JavaScript
- **后端**：Python, FastAPI (IPC通信)
- **文档处理**：python-docx
- **规则引擎**：自定义规则引擎
- **配置管理**：YAML

## 核心特性

- ✅ 基于规则引擎的架构设计
- ✅ 模块化的规则实现
- ✅ 完整的预设管理功能
- ✅ 支持自定义规则配置
- ✅ 前后端分离架构
- ✅ 清晰的API接口设计
- ✅ 支持多种预设配置
- ✅ 可扩展的规则系统

## 功能特性

### 文本格式修复
- ✅ **文本颜色修复**：统一文本颜色为黑色
- ✅ **标题格式修复**：一级标题居中，其他标题左对齐
- ✅ **字体定制**：中文字体使用宋体/黑体，西文字体使用Arial
- ✅ **字号标准化**：统一正文和标题字号

### 列表与编号
- ✅ **编号格式修复**：修复不正确的编号格式
- ✅ **项目符号清理**：移除不需要的项目符号（如·）
- ✅ **列表样式统一**：统一列表的样式和格式

### 表格优化
- ✅ **表格边框添加**：为表格添加边框
- ✅ **表格宽度优化**：自动调整表格宽度和列宽
- ✅ **表格居中**：设置表格居中对齐
- ✅ **表格标题格式化**：优化表格标题样式

### 页面布局
- ✅ **页面边距设置**：统一页面边距
- ✅ **页面大小设置**：设置标准页面大小

### 系统功能
- ✅ **响应式GUI界面**：支持窗口大小调整，美观易用
- ✅ **配置参数可调**：可自定义字体、字号、页面边距等
- ✅ **多预设配置**：内置多种预设配置
- ✅ **命令行支持**：可通过命令行批量处理文件
- ✅ **预设管理**：支持创建、编辑和删除预设
- ✅ **规则引擎**：基于规则引擎的架构设计
- ✅ **可扩展规则**：支持添加新的规则
- ✅ **执行报告**：生成详细的执行报告

### 预设配置
- **default**：默认配置，适合大多数文档
- **minimal**：最小配置，仅包含必要规则
- **comprehensive**：全面配置，包含所有规则
- **academic**：学术论文格式，符合学术论文规范
- **business**：企业公文规范，适合企业正式文档
- **bid**：竞标书标准，专业竞标文档格式

## 系统要求

- Windows 7+ / macOS / Linux
- Python 3.6+

## 安装方法

### 方法一：直接运行可执行文件

1. 从 [GitHub Releases](https://github.com/ilovend/word_format_fixer/releases) 下载最新的可执行文件
2. 双击运行 `WordFormatFixer.exe`（Windows）或相应的可执行文件

### 方法二：从源码运行

1. 克隆仓库：
   ```bash
   git clone https://github.com/ilovend/word_format_fixer.git
   cd word_format_fixer
   ```

2. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```

3. 运行工具：
   ```bash
   python run.py
   ```

## 使用方法

### 1. Electron GUI模式

1. **启动应用**：
   - 双击运行 `win启动应用.bat`（Windows）
   - 或使用命令：`npm start`（在electron目录下）

2. **选择文件**：
   - 点击"选择Word文档"按钮选择输入文件
   - 或直接拖拽.docx文件到应用窗口

3. **选择预设**：
   - 在左侧预设管理面板中选择合适的预设
   - 或点击"新建预设"创建自定义预设

4. **配置规则**：
   - 在中间规则配置面板中查看和修改规则
   - 启用或禁用特定规则
   - 调整规则参数

5. **执行修复**：
   - 点击右侧"处理文档"按钮执行修复
   - 查看执行报告和详细结果

6. **查看结果**：
   - 修复完成后，系统会生成详细的执行报告
   - 报告包含修复的详细信息和统计数据

### 2. 命令行模式

**注意**：命令行模式正在更新中，请优先使用GUI模式。

```bash
# 基本用法
python run.py input.docx

# 输出到指定文件
python run.py input.docx -o output.docx

# 使用预设配置
python run.py input.docx --preset default
```

## 配置选项

### 1. 规则配置

每个规则都有自己的配置参数，可以在UI中直接调整：

#### 字体规则
- **字体名称标准化**：设置中文字体和西文字体
- **标题字体设置**：设置标题专用字体
- **字号标准化**：设置正文和标题字号
- **字体颜色统一**：设置文本颜色

#### 段落规则
- **段落间距设置**：设置行间距、段前段后间距
- **标题加粗**：设置标题是否加粗
- **标题对齐**：设置标题对齐方式
- **列表编号修复**：修复列表编号

#### 表格规则
- **表格宽度优化**：设置表格宽度百分比和列宽自动调整
- **表格边框添加**：设置表格边框大小和颜色

#### 页面规则
- **页面布局**：设置页面大小、边距等

### 2. 预设配置

预设是一组规则配置的集合，可以快速应用到文档：

| 预设名称 | 描述 |
|---------|------|
| default | 默认配置，包含所有常用规则 |
| minimal | 最小配置，仅包含必要规则 |
| comprehensive | 全面配置，包含所有规则 |
| academic | 学术论文格式，符合学术规范 |
| business | 企业公文规范，适合正式文档 |
| bid | 竞标书标准，专业文档格式 |

### 3. 自定义预设

可以创建、编辑和删除预设：
- 点击"新建预设"按钮
- 输入预设名称和描述
- 选择要启用的规则
- 调整规则参数
- 保存预设

## 快捷键

| 快捷键 | 功能 |
|-------|------|
| Ctrl+R | 开始修复文档 |
| Ctrl+O | 选择输入文件 |
| Ctrl+Q | 退出程序 |
| Ctrl+N | 新建预设 |
| Ctrl+E | 编辑当前预设 |
| Ctrl+D | 删除当前预设 |

## 构建与部署

### 1. 开发环境搭建

#### Python后端
```bash
# 安装依赖
pip install -r requirements.txt

# 启动后端服务
python python-backend/main.py
```

#### Electron前端
```bash
# 进入electron目录
cd electron

# 安装依赖
npm install

# 启动前端应用
npm start
```

### 2. 构建可执行文件

#### Windows平台

1. **构建Python后端**：
   - 使用 `build.bat` 脚本构建后端可执行文件
   - 命令：`.uild.bat`

2. **构建Electron应用**：
   - 进入electron目录
   - 命令：`npm run build`
   - 构建完成后，可执行文件会在 `electron/dist` 目录中生成

#### macOS/Linux平台

1. **构建Python后端**：
   - 使用 `build.sh` 脚本构建后端可执行文件
   - 命令：`bash build.sh`

2. **构建Electron应用**：
   - 进入electron目录
   - 命令：`npm run build`
   - 构建完成后，可执行文件会在 `electron/dist` 目录中生成

### 3. 部署

- 将构建好的可执行文件和必要的配置文件打包
- 确保用户系统中已安装必要的依赖
- 提供详细的安装和使用说明

## 常见问题

### 1. 文档格式问题

**Q: 修复后的文档打开时提示格式错误**
A: 这可能是由于文档结构过于复杂导致的。请尝试使用"minimal"预设配置，或检查文档是否有损坏。

**Q: 表格没有完全修复**
A: 复杂表格的修复可能需要手动调整。工具会尽量修复基本的表格边框和列宽问题。

**Q: 字体显示不正确**
A: 请确保您的系统中安装了所需的字体（宋体、黑体、Arial）。

**Q: 规则不生效**
A: 请检查：
- 该规则是否已启用
- 规则配置是否正确
- 文档是否符合规则的适用条件

### 2. 系统问题

**Q: 应用无法启动**
A: 请检查：
- 是否已安装Python 3.6+
- 是否已安装必要的依赖
- 端口是否被占用（默认使用7777端口）

**Q: 应用运行缓慢**
A: 请尝试：
- 关闭其他占用系统资源的应用
- 使用"minimal"预设配置
- 减少文档的复杂度

## 开发者指南

### 1. 项目结构

```
word_format_fixer/
├── AIPoliDoc/           # AIPoliDoc相关代码
├── docs/               # 项目文档
├── electron/           # Electron前端
├── python-backend/     # Python后端
├── tests/              # 测试代码
├── word_format_fixer/  # 原始项目代码
├── LICENSE            # 许可证文件
├── README.md          # 项目说明文档
├── build.bat          # Windows构建脚本
├── requirements.txt   # Python依赖
└── run.py             # 运行脚本
```

### 2. 规则开发

1. **创建新规则**：
   - 在 `python-backend/rules/` 目录下创建新的规则文件
   - 继承 `BaseRule` 类
   - 实现 `apply` 方法
   - 设置规则元数据

2. **注册规则**：
   - 在 `python-backend/rules/__init__.py` 中注册新规则

3. **测试规则**：
   - 编写单元测试
   - 在UI中测试规则效果

### 3. API文档

#### 后端API

| 端点 | 方法 | 功能 |
|------|------|------|
| /api/health | GET | 健康检查 |
| /api/rules | GET | 获取所有规则 |
| /api/presets | GET | 获取所有预设 |
| /api/presets/save | POST | 保存预设 |
| /api/presets/delete | DELETE | 删除预设 |
| /api/process | POST | 处理文档 |

## 贡献

欢迎提交Issue和Pull Request来改进这个工具！

### 贡献流程

1. Fork本仓库
2. 创建您的特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交您的修改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开一个Pull Request

### 贡献指南

- 遵循项目的代码风格
- 编写单元测试
- 更新文档
- 确保所有测试通过
- 提供详细的PR描述

## 许可证

本项目采用 MIT 许可证 - 详见 [LICENSE](LICENSE) 文件

## 联系方式

- GitHub: [ilovend](https://github.com/ilovend)
- 邮箱: ilovendme@outlook.com

---

**注意**：本工具仅用于修复文档格式问题，不会修改文档内容。建议在使用前备份原始文档。

**更新日期**：2026-01-22