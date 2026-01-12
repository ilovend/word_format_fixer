# Word文档格式修复工具

一个用于修复Markdown转换为Word文档后格式问题的Python工具。

## 功能特性

- ✅ **文本颜色修复**：统一文本颜色为黑色
- ✅ **标题格式修复**：一级标题居中，其他标题左对齐
- ✅ **字体定制**：中文字体使用宋体/黑体，西文字体使用Arial
- ✅ **编号格式修复**：修复不正确的编号格式
- ✅ **项目符号清理**：移除不需要的项目符号（如·）
- ✅ **表格优化**：添加表格边框，自动调整列宽
- ✅ **响应式GUI界面**：支持窗口大小调整，美观易用
- ✅ **配置参数可调**：可自定义字体、字号、页面边距等
- ✅ **多预设配置**：内置默认、标书专用、紧凑、打印就绪等预设
- ✅ **命令行支持**：可通过命令行批量处理文件

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

### GUI模式

1. 双击运行 `WordFormatFixer.exe` 或 `python run.py`
2. 点击"浏览"按钮选择输入Word文件
3. 选择或修改输出文件路径（默认在原文件同目录下添加 `_fixed` 后缀）
4. 选择预设配置或自定义配置选项
5. 点击"开始修复"按钮执行修复
6. 修复完成后会显示成功提示，并在日志窗口显示详细信息

### 命令行模式

```bash
# 基本用法
python run.py --cli input.docx -o output.docx

# 使用预设配置
python run.py --cli input.docx -o output.docx --preset bid_document
```

## 配置选项

### 预设配置
- **default**：默认配置，适合大多数文档
- **bid_document**：标书专用配置，格式更加规范
- **compact**：紧凑配置，节省空间
- **print_ready**：打印就绪配置，优化打印效果

### 自定义配置
- **中文字体**：默认使用宋体
- **西文字体**：默认使用Arial
- **标题字体**：默认使用黑体
- **字号设置**：可调整正文、一级标题、二级标题、三级标题的字号
- **页面边距**：可调整上下左右边距（单位：厘米）
- **表格设置**：可调整表格宽度百分比和自动调整列宽选项

## 快捷键

- **Ctrl+R**：开始修复
- **Ctrl+O**：选择输入文件
- **Ctrl+Q**：退出程序

## 构建可执行文件

如果需要从源码构建可执行文件：

1. 安装PyInstaller：
   ```bash
   pip install pyinstaller
   ```

2. 运行构建脚本：
   ```bash
   # Windows
   .\build.bat
   
   # macOS/Linux
   bash build.sh
   ```

3. 构建完成后，可执行文件会在 `dist` 目录中生成

## 常见问题

### Q: 修复后的文档打开时提示格式错误
A: 这可能是由于文档结构过于复杂导致的。请尝试使用"标书专用"预设配置，或联系作者获取支持。

### Q: 表格没有完全修复
A: 复杂表格的修复可能需要手动调整。工具会尽量修复基本的表格边框和列宽问题。

### Q: 字体显示不正确
A: 请确保您的系统中安装了所需的字体（宋体、黑体、Arial）。

## 贡献

欢迎提交Issue和Pull Request来改进这个工具！

## 许可证

本项目采用 MIT 许可证 - 详见 [LICENSE](LICENSE) 文件

## 联系方式

- GitHub: [ilovend](https://github.com/ilovend)
- 邮箱: [your-email@example.com]（请替换为实际邮箱）

---

**注意**：本工具仅用于修复文档格式问题，不会修改文档内容。建议在使用前备份原始文档。