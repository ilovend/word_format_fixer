"""Word文档格式修复工具 - PyQt-Fluent-Widgets主窗口"""

from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QTextEdit, QListWidget, QGroupBox, QLabel, QFileDialog
from PyQt5.QtCore import Qt
from qfluentwidgets import (
    FluentWindow, NavigationItemPosition, 
    CardWidget, 
    PrimaryPushButton, PushButton, 
    LineEdit, ComboBox, CheckBox, 
    setTheme, Theme
)

from ...core.fixer import RobustWordFixer
from ...core.config import get_preset_config
import os


class HomePage(QWidget):
    """主页面"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("HomePage")
        self.setup_ui()
        self.fixer = None
    
    def setup_ui(self):
        """设置UI"""
        # 主布局
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setSpacing(20)
        self.main_layout.setContentsMargins(24, 24, 24, 24)
        
        # 标题区域
        self.setup_title_section()
        
        # 内容区域
        self.setup_content_section()
        
        # 操作按钮区域
        self.setup_button_section()
        
        # 日志区域
        self.setup_log_section()
    
    def setup_title_section(self):
        """设置标题区域"""
        # 标题布局
        title_layout = QHBoxLayout()
        
        # 应用标题
        title_label = QLabel("Word文档格式修复工具")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #1F1F1F;")
        
        # 版本信息
        version_label = QLabel("v1.1.2")
        version_label.setStyleSheet("font-size: 12px; color: #666666;")
        
        title_layout.addWidget(title_label)
        title_layout.addStretch()
        title_layout.addWidget(version_label)
        
        self.main_layout.addLayout(title_layout)
    
    def setup_content_section(self):
        """设置内容区域"""
        # 内容区域布局
        content_layout = QVBoxLayout()
        
        # 文件选择卡片
        file_card = CardWidget()
        file_layout = QVBoxLayout(file_card)
        file_layout.setSpacing(12)
        
        # 批量模式切换
        batch_layout = QHBoxLayout()
        self.batch_mode = CheckBox("批量模式")
        batch_hint = QLabel("(支持选择多个文件或文件夹)")
        batch_hint.setStyleSheet("font-size: 11px; color: #999999;")
        
        batch_layout.addWidget(self.batch_mode)
        batch_layout.addWidget(batch_hint)
        batch_layout.addStretch()
        
        # 单文件模式
        single_layout = QHBoxLayout()
        single_layout.addWidget(QLabel("输入文件:"))
        self.input_edit = LineEdit()
        self.input_button = PushButton("浏览")
        self.input_button.clicked.connect(self.browse_input_file)
        
        single_layout.addWidget(self.input_edit, 1)
        single_layout.addWidget(self.input_button)
        
        # 批量模式
        batch_files_layout = QVBoxLayout()
        batch_select_layout = QHBoxLayout()
        batch_select_layout.addWidget(QLabel("批量选择:"))
        self.batch_files_button = PushButton("选择多个文件")
        self.batch_folder_button = PushButton("选择文件夹")
        
        self.batch_files_button.clicked.connect(self.browse_batch_files)
        self.batch_folder_button.clicked.connect(self.browse_batch_folder)
        
        batch_select_layout.addWidget(self.batch_files_button)
        batch_select_layout.addWidget(self.batch_folder_button)
        
        # 文件列表
        self.file_list = QListWidget()
        self.file_list.setMinimumHeight(120)
        
        batch_files_layout.addLayout(batch_select_layout)
        batch_files_layout.addWidget(self.file_list)
        
        # 输出设置
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("输出路径:"))
        self.output_edit = LineEdit()
        self.output_button = PushButton("浏览")
        self.output_button.clicked.connect(self.browse_output_file)
        
        output_layout.addWidget(self.output_edit, 1)
        output_layout.addWidget(self.output_button)
        
        # 添加到文件卡片
        file_layout.addLayout(batch_layout)
        file_layout.addLayout(single_layout)
        file_layout.addLayout(batch_files_layout)
        file_layout.addLayout(output_layout)
        
        # 配置卡片
        config_card = CardWidget()
        config_layout = QVBoxLayout(config_card)
        config_layout.setSpacing(12)
        
        # 预设配置
        preset_layout = QHBoxLayout()
        preset_layout.addWidget(QLabel("预设配置:"))
        self.preset_combo = ComboBox()
        self.preset_combo.addItems(["默认", "标书专用", "紧凑", "打印就绪", "学术论文", "简历", "报告", "演示"])
        
        preset_layout.addWidget(self.preset_combo)
        preset_layout.addStretch()
        
        # 添加到配置卡片
        config_layout.addLayout(preset_layout)
        
        # 添加到内容区域
        content_layout.addWidget(file_card)
        content_layout.addWidget(config_card)
        
        self.main_layout.addLayout(content_layout)
    
    def setup_button_section(self):
        """设置按钮区域"""
        # 按钮布局
        button_layout = QHBoxLayout()
        
        # 开始修复按钮
        self.fix_button = PrimaryPushButton("开始修复")
        self.fix_button.setMinimumWidth(120)
        self.fix_button.clicked.connect(self.fix_document)
        
        # 重置按钮
        self.reset_button = PushButton("重置")
        self.reset_button.clicked.connect(self.reset_settings)
        
        # 配置按钮
        self.config_button = PushButton("配置")
        
        # 退出按钮
        self.exit_button = PushButton("退出")
        self.exit_button.clicked.connect(QApplication.instance().quit)
        
        button_layout.addWidget(self.fix_button)
        button_layout.addWidget(self.reset_button)
        button_layout.addWidget(self.config_button)
        button_layout.addStretch()
        button_layout.addWidget(self.exit_button)
        
        self.main_layout.addLayout(button_layout)
    
    def setup_log_section(self):
        """设置日志区域"""
        # 日志卡片
        log_card = CardWidget()
        log_layout = QVBoxLayout(log_card)
        
        # 日志标题
        log_title = QLabel("操作日志")
        log_title.setStyleSheet("font-size: 12px; font-weight: bold; color: #333333;")
        
        # 日志文本框
        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setStyleSheet("font-family: Consolas; font-size: 11px; background-color: #F5F5F5;")
        self.log_edit.setMinimumHeight(120)
        
        log_layout.addWidget(log_title)
        log_layout.addWidget(self.log_edit)
        
        self.main_layout.addWidget(log_card)
    
    def browse_input_file(self):
        """浏览输入文件"""
        dialog = QFileDialog(self, "选择输入文件", ".", "Word文档 (*.docx);;所有文件 (*.*)")
        dialog.setFileMode(QFileDialog.ExistingFile)
        if dialog.exec_():
            filename = dialog.selectedFiles()[0]
            self.input_edit.setText(filename)
            # 自动设置输出路径
            if not self.output_edit.text():
                base, ext = os.path.splitext(filename)
                output_filename = f"{base}_fixed{ext}"
                self.output_edit.setText(output_filename)
    
    def browse_output_file(self):
        """浏览输出文件"""
        dialog = QFileDialog(self, "选择输出文件", ".", "Word文档 (*.docx);;所有文件 (*.*)")
        dialog.setFileMode(QFileDialog.AnyFile)
        dialog.setAcceptMode(QFileDialog.AcceptSave)
        if dialog.exec_():
            filename = dialog.selectedFiles()[0]
            self.output_edit.setText(filename)
    
    def browse_batch_files(self):
        """浏览多个文件"""
        dialog = QFileDialog(self, "选择多个文件", ".", "Word文档 (*.docx);;所有文件 (*.*)")
        dialog.setFileMode(QFileDialog.ExistingFiles)
        if dialog.exec_():
            files = dialog.selectedFiles()
            self.file_list.clear()
            for file in files:
                self.file_list.addItem(file)
    
    def browse_batch_folder(self):
        """浏览文件夹"""
        dialog = QFileDialog(self, "选择文件夹", ".", "文件夹")
        dialog.setFileMode(QFileDialog.Directory)
        dialog.setOption(QFileDialog.ShowDirsOnly, True)
        if dialog.exec_():
            folder = dialog.selectedFiles()[0]
            # 遍历文件夹中的所有docx文件
            docx_files = []
            for root, _, files in os.walk(folder):
                for file in files:
                    if file.lower().endswith('.docx'):
                        docx_files.append(os.path.join(root, file))
            
            if docx_files:
                self.file_list.clear()
                for file in docx_files:
                    self.file_list.addItem(file)
    
    def fix_document(self):
        """开始修复文档"""
        try:
            # 构建配置
            preset_map = {
                "默认": "default",
                "标书专用": "bid_document",
                "紧凑": "compact",
                "打印就绪": "print_ready",
                "学术论文": "academic_paper",
                "简历": "resume",
                "报告": "report",
                "演示": "presentation"
            }
            
            preset_name = self.preset_combo.currentText()
            preset_key = preset_map.get(preset_name, "default")
            config = get_preset_config(preset_key)
            
            # 创建修复器
            self.fixer = RobustWordFixer(config)
            
            # 清空日志
            self.log_edit.clear()
            
            # 执行修复
            batch_mode = self.batch_mode.isChecked()
            if batch_mode:
                # 批量模式
                files = [self.file_list.item(i).text() for i in range(self.file_list.count())]
                if not files:
                    self.log("错误: 请选择要处理的文件")
                    return
                
                output_dir = self.output_edit.text() if self.output_edit.text() else None
                results = self.fixer.fix_batch(files, output_dir)
                
                # 统计结果
                success_count = sum(1 for v in results.values() if v is not None)
                failed_count = len(results) - success_count
                
                self.log(f"批量修复完成！")
                self.log(f"成功: {success_count} 个文件")
                self.log(f"失败: {failed_count} 个文件")
            else:
                # 单文件模式
                input_file = self.input_edit.text()
                output_file = self.output_edit.text()
                
                # 验证输入
                if not input_file:
                    self.log("错误: 请选择输入文件")
                    return
                
                if not output_file:
                    self.log("错误: 请选择输出文件")
                    return
                
                if not os.path.exists(input_file):
                    self.log("错误: 输入文件不存在")
                    return
                
                result = self.fixer.fix_all(input_file, output_file)
                
                if result:
                    self.log(f"文档修复完成！")
                    self.log(f"输出文件: {result}")
                else:
                    self.log("文档修复失败，请查看详细日志")
        
        except Exception as e:
            self.log(f"错误: {str(e)}")
    
    def reset_settings(self):
        """重置设置"""
        self.input_edit.clear()
        self.output_edit.clear()
        self.file_list.clear()
        self.batch_mode.setChecked(False)
        self.preset_combo.setCurrentIndex(0)
        self.log_edit.clear()
    
    def log(self, message):
        """记录日志"""
        self.log_edit.append(message)


class WordFixerWindow(FluentWindow):
    """Word文档修复工具主窗口"""
    
    def __init__(self):
        super().__init__()
        self.setup_ui()
    
    def setup_ui(self):
        """设置UI"""
        # 设置主题
        setTheme(Theme.LIGHT)
        
        # 创建主页面
        self.home_page = HomePage(self)
        
        # 添加页面
        self.addSubInterface(self.home_page, "home", "Word文档格式修复工具", NavigationItemPosition.TOP)
        
        # 设置窗口属性
        self.setWindowTitle("Word文档格式修复工具")
        self.resize(900, 700)
        self.setMinimumSize(700, 500)
        
        # 添加快捷键
        self.setup_shortcuts()
        
        # 显示窗口
        self.show()
    
    def setup_shortcuts(self):
        """设置快捷键"""
        from PyQt5.QtWidgets import QShortcut
        from PyQt5.QtGui import QKeySequence
        
        # 开始修复
        QShortcut(QKeySequence('Ctrl+R'), self, self.home_page.fix_document)
        
        # 退出程序
        QShortcut(QKeySequence('Ctrl+Q'), self, QApplication.instance().quit)


def run_app():
    """运行应用"""
    app = QApplication([])
    window = WordFixerWindow()
    app.exec_()


if __name__ == "__main__":
    run_app()
