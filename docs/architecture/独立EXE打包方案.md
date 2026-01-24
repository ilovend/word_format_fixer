# Word格式修复工具独立EXE安装包打包方案

## 一、项目现状分析

1. **当前架构**：
   - 前端：Electron应用
   - 后端：Python服务，通过HTTP API与前端通信
   - 启动方式：start.bat脚本先启动Python后端，再启动Electron前端

2. **打包需求**：
   - 无需外部后端
   - 单个EXE安装包
   - 可在本机直接运行
   - 无需安装Python环境

## 二、解决方案

### 1. 使用PyInstaller将Python后端打包成EXE

**步骤**：
- 创建PyInstaller配置文件
- 将Python后端代码和依赖打包成单个EXE
- 确保所有规则和配置文件被正确打包

**修改点**：
- 编写PyInstaller.spec文件
- 确保配置文件和规则文件被正确包含

### 2. 修改Electron主进程代码

**步骤**：
- 修改`startPythonServer`函数，使用打包后的Python EXE而不是直接调用python命令
- 确保Electron应用能够找到打包后的Python EXE
- 处理Python EXE的启动和关闭

**修改点**：
- `electron/main.js`：修改Python服务器启动逻辑

### 3. 使用electron-builder打包完整应用

**步骤**：
- 修改electron-builder配置，将Python EXE和资源文件包含在内
- 配置NSIS安装包生成选项
- 确保所有必要文件被正确打包

**修改点**：
- `electron/package.json`：更新build配置
- 添加必要的构建脚本

### 4. 处理资源文件

**步骤**：
- 确保配置文件、规则文件和图标文件被正确打包
- 处理应用运行时资源文件的访问路径

**修改点**：
- 更新Python代码中资源文件的访问方式
- 更新Electron代码中资源文件的访问方式

## 三、具体实现细节

### 1. 后端打包配置

- 创建`python-backend/pyinstaller.spec`文件
- 配置打包选项，包括：
  - 单文件模式
  - 包含所有依赖
  - 包含配置文件和规则文件
  - 设置正确的入口点

### 2. Electron主进程修改

- 修改`startPythonServer`函数：
  - 检查当前环境，使用打包后的Python EXE
  - 设置正确的工作目录和参数
  - 处理进程的启动、错误和关闭事件

### 3. Electron打包配置

- 更新`electron/package.json`中的build配置：
  - 添加extraResources配置，包含Python EXE和资源文件
  - 更新NSIS配置，设置安装选项
  - 配置应用图标和版本信息

### 4. 测试和验证

- 测试打包后的应用在不同环境下的运行情况
- 验证所有功能正常工作
- 确保资源文件正确访问
- 测试安装和卸载过程

## 四、预期结果

- 生成单个EXE安装包
- 无需安装Python环境
- 应用启动时自动启动内嵌的Python后端
- 所有功能正常工作
- 支持安装、卸载和更新

## 五、文件修改列表

1. **python-backend/pyinstaller.spec**：新增，PyInstaller打包配置
2. **electron/main.js**：修改，更新Python服务器启动逻辑
3. **electron/package.json**：修改，更新electron-builder配置
4. **python-backend/config_loader.py**：修改，处理资源文件路径
5. **python-backend/ipc/adapter.py**：修改，处理资源文件路径

## 六、优势

- 无需安装Python环境
- 单个安装包，便于分发和使用
- 无需手动启动后端服务
- 应用启动更快
- 更好的用户体验

## 七、风险和注意事项

- 打包后的EXE文件可能较大
- 需要确保所有依赖被正确包含
- 需要测试不同版本的Windows系统
- 需要处理资源文件的相对路径
- 需要确保Python EXE能够在目标系统上运行