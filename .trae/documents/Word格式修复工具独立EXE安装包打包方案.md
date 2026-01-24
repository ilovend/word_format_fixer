# Word格式修复工具独立EXE打包计划

## 1. 安装PyInstaller
```bash
pip install pyinstaller
```

## 2. 打包Python后端
```bash
cd python-backend
pyinstaller --onefile --name=word_format_fixer_backend main.py
```

## 3. 修改Electron主进程
更新`electron/main.js`中的`startPythonServer`函数，将Python后端路径修改为打包后的可执行文件路径

## 4. 配置electron-builder
更新`electron/package.json`中的build配置，添加extraResources配置，将打包后的Python后端包含在安装包中

## 5. 打包Electron应用
```bash
cd electron
npm install
npm run build
```

## 6. 测试打包后的应用
运行生成的EXE安装包，验证应用能正常启动和运行

## 7. 优化用户体验
确保安装过程简单直观，添加桌面快捷方式，确保卸载干净

## 注意事项
- 使用PyInstaller的`--onefile`选项生成单个EXE文件
- 确保electron-builder配置正确包含Python后端
- 测试时注意防火墙和杀毒软件的影响
- 考虑添加应用图标和版本信息