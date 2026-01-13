"""Word文档格式修复工具 - PyQt-Fluent-Widgets主窗口"""

from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QTextEdit, QListWidget, QGroupBox, QLabel, QFileDialog, QDialog
from PyQt5.QtCore import Qt
from qfluentwidgets import (
    FluentWindow, NavigationItemPosition, 
    CardWidget, 
    PrimaryPushButton, PushButton, 
    LineEdit, ComboBox, CheckBox, 
    setTheme, Theme
)

# 统一字体设置
FONT_STYLES = {
    "title": "font-size: 18px; font-weight: bold; color: #1F1F1F; font-family: 'Microsoft YaHei', 'SimHei', sans-serif;",
    "subtitle": "font-size: 14px; font-weight: bold; color: #333333; font-family: 'Microsoft YaHei', 'SimHei', sans-serif;",
    "label": "font-size: 13px; color: #333333; font-family: 'Microsoft YaHei', 'SimSun', sans-serif;",
    "small_label": "font-size: 11px; color: #666666; font-family: 'Microsoft YaHei', 'SimSun', sans-serif;",
    "hint": "font-size: 11px; color: #999999; font-family: 'Microsoft YaHei', 'SimSun', sans-serif;",
    "button": "font-size: 13px; font-family: 'Microsoft YaHei', 'SimHei', sans-serif;",
    "input": "font-size: 13px; font-family: 'Microsoft YaHei', 'SimSun', sans-serif;",
    "log": "font-family: Consolas, 'Courier New', monospace; font-size: 11px; background-color: #F5F5F5;"
}

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
        
        # 连接窗口显示事件，打印UI元素位置和大小
        self.showEvent = self.on_show_event
    
    def setup_title_section(self):
        """设置标题区域"""
        # 标题布局
        title_layout = QHBoxLayout()
        
        # 应用标题
        title_label = QLabel("Word文档格式修复工具")
        title_label.setStyleSheet(FONT_STYLES["title"])
        
        # 版本信息
        version_label = QLabel("v1.1.2")
        version_label.setStyleSheet(FONT_STYLES["small_label"])
        
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
        
        # 添加调试信息，打印布局结构
        print("=== UI布局结构 ===")
        print("文件选择卡片:", file_card)
        print("文件选择卡片布局:", file_layout)
        
        # 批量模式切换
        batch_layout = QHBoxLayout()
        self.batch_mode = CheckBox("批量模式")
        self.batch_mode.setStyleSheet(FONT_STYLES["label"])
        batch_hint = QLabel("(支持选择多个文件或文件夹)")
        batch_hint.setStyleSheet(FONT_STYLES["hint"])
        
        batch_layout.addWidget(self.batch_mode)
        batch_layout.addWidget(batch_hint)
        batch_layout.addStretch()
        
        # 单文件模式
        single_layout = QHBoxLayout()
        input_label = QLabel("输入文件:")
        input_label.setStyleSheet(FONT_STYLES["label"])
        single_layout.addWidget(input_label)
        self.input_edit = LineEdit()
        self.input_edit.setStyleSheet(FONT_STYLES["input"])
        self.input_button = PushButton("浏览")
        self.input_button.setStyleSheet(FONT_STYLES["button"])
        self.input_button.clicked.connect(self.browse_input_file)
        
        single_layout.addWidget(self.input_edit, 1)
        single_layout.addWidget(self.input_button)
        
        # 批量模式
        batch_select_layout = QHBoxLayout()
        batch_label = QLabel("批量选择:")
        batch_label.setStyleSheet(FONT_STYLES["label"])
        batch_select_layout.addWidget(batch_label)
        self.batch_files_button = PushButton("选择多个文件")
        self.batch_files_button.setStyleSheet(FONT_STYLES["button"])
        self.batch_folder_button = PushButton("选择文件夹")
        self.batch_folder_button.setStyleSheet(FONT_STYLES["button"])
        
        self.batch_files_button.clicked.connect(self.browse_batch_files)
        self.batch_folder_button.clicked.connect(self.browse_batch_folder)
        
        batch_select_layout.addWidget(self.batch_files_button)
        batch_select_layout.addWidget(self.batch_folder_button)
        
        # 文件列表
        self.file_list = QListWidget()
        self.file_list.setMinimumHeight(200)  # 增加最小高度，确保能显示更多文件
        # 移除黑色框线
        self.file_list.setStyleSheet("border: none; background-color: #FFFFFF;")
        
        # 输出设置
        output_layout = QHBoxLayout()
        output_label = QLabel("输出路径:")
        output_label.setStyleSheet(FONT_STYLES["label"])
        output_layout.addWidget(output_label)
        self.output_edit = LineEdit()
        self.output_edit.setStyleSheet(FONT_STYLES["input"])
        self.output_button = PushButton("浏览")
        self.output_button.setStyleSheet(FONT_STYLES["button"])
        self.output_button.clicked.connect(self.browse_output_file)
        
        output_layout.addWidget(self.output_edit, 1)
        output_layout.addWidget(self.output_button)
        
        # 添加调试信息，打印布局顺序
        print("=== 布局顺序 ===")
        print("1. 批量模式切换:", batch_layout)
        print("2. 单文件模式:", single_layout)
        print("3. 批量选择按钮:", batch_select_layout)
        print("4. 文件列表:", self.file_list)
        print("5. 输出设置:", output_layout)
        
        # 重新组织布局结构，避免嵌套布局的复杂性
        # 直接在 file_layout 中添加所有元素
        file_layout.addLayout(batch_layout)
        file_layout.addLayout(single_layout)
        file_layout.addLayout(batch_select_layout)
        file_layout.addWidget(self.file_list)
        
        # 配置卡片
        config_card = CardWidget()
        config_layout = QVBoxLayout(config_card)
        config_layout.setSpacing(12)
        
        # 预设配置
        preset_layout = QHBoxLayout()
        preset_label = QLabel("预设配置:")
        preset_label.setStyleSheet(FONT_STYLES["label"])
        preset_layout.addWidget(preset_label)
        self.preset_combo = ComboBox()
        self.preset_combo.setStyleSheet(FONT_STYLES["input"])
        self.preset_combo.addItems(["默认", "标书专用", "紧凑", "打印就绪", "学术论文", "简历", "报告", "演示"])
        
        preset_layout.addWidget(self.preset_combo)
        preset_layout.addStretch()
        
        # 添加到配置卡片
        config_layout.addLayout(preset_layout)
        
        # 输出设置卡片
        output_card = CardWidget()
        output_card_layout = QVBoxLayout(output_card)
        output_card_layout.setSpacing(12)
        
        # 输出设置
        output_layout = QHBoxLayout()
        output_label = QLabel("输出路径:")
        output_label.setStyleSheet(FONT_STYLES["label"])
        output_layout.addWidget(output_label)
        self.output_edit = LineEdit()
        self.output_edit.setStyleSheet(FONT_STYLES["input"])
        self.output_button = PushButton("浏览")
        self.output_button.setStyleSheet(FONT_STYLES["button"])
        self.output_button.clicked.connect(self.browse_output_file)
        
        output_layout.addWidget(self.output_edit, 1)
        output_layout.addWidget(self.output_button)
        
        output_card_layout.addLayout(output_layout)
        
        # 添加到内容区域
        content_layout.addWidget(file_card)
        content_layout.addWidget(config_card)
        # 将输出设置卡片添加到配置卡片的下方，这样它就会在文件列表的下方
        content_layout.addWidget(output_card)
        
        self.main_layout.addLayout(content_layout)
    
    def setup_button_section(self):
        """设置按钮区域"""
        # 按钮布局
        button_layout = QHBoxLayout()
        
        # 开始修复按钮
        self.fix_button = PrimaryPushButton("开始修复")
        self.fix_button.setMinimumWidth(120)
        self.fix_button.setStyleSheet(FONT_STYLES["button"])
        self.fix_button.clicked.connect(self.fix_document)
        
        # 重置按钮
        self.reset_button = PushButton("重置")
        self.reset_button.setStyleSheet(FONT_STYLES["button"])
        self.reset_button.clicked.connect(self.reset_settings)
        
        # 配置按钮
        self.config_button = PushButton("配置")
        self.config_button.setStyleSheet(FONT_STYLES["button"])
        self.config_button.clicked.connect(self.open_config_dialog)
        
        # 退出按钮
        self.exit_button = PushButton("退出")
        self.exit_button.setStyleSheet(FONT_STYLES["button"])
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
        log_title.setStyleSheet(FONT_STYLES["subtitle"])
        
        # 日志文本框
        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setStyleSheet(FONT_STYLES["log"])
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
    
    def open_config_dialog(self):
        """打开配置对话框"""
        from PyQt5.QtWidgets import QVBoxLayout, QHBoxLayout, QPushButton, QGridLayout
        from qfluentwidgets import CardWidget, LineEdit, ComboBox, SpinBox, DoubleSpinBox, CheckBox
        
        # 创建配置对话框
        dialog = QDialog(self)
        dialog.setWindowTitle("配置选项")
        dialog.resize(600, 400)
        
        # 主布局
        main_layout = QVBoxLayout(dialog)
        
        # 字体设置卡片
        font_card = CardWidget()
        font_layout = QGridLayout(font_card)
        
        # 中文字体
        chinese_font_label = QLabel("中文字体:")
        chinese_font_label.setStyleSheet(FONT_STYLES["label"])
        font_layout.addWidget(chinese_font_label, 0, 0)
        self.chinese_font_edit = LineEdit()
        self.chinese_font_edit.setStyleSheet(FONT_STYLES["input"])
        self.chinese_font_edit.setText("宋体")
        font_layout.addWidget(self.chinese_font_edit, 0, 1)
        
        # 西文字体
        western_font_label = QLabel("西文字体:")
        western_font_label.setStyleSheet(FONT_STYLES["label"])
        font_layout.addWidget(western_font_label, 0, 2)
        self.western_font_edit = LineEdit()
        self.western_font_edit.setStyleSheet(FONT_STYLES["input"])
        self.western_font_edit.setText("Arial")
        font_layout.addWidget(self.western_font_edit, 0, 3)
        
        # 标题字体
        title_font_label = QLabel("标题字体:")
        title_font_label.setStyleSheet(FONT_STYLES["label"])
        font_layout.addWidget(title_font_label, 1, 0)
        self.title_font_edit = LineEdit()
        self.title_font_edit.setStyleSheet(FONT_STYLES["input"])
        self.title_font_edit.setText("黑体")
        font_layout.addWidget(self.title_font_edit, 1, 1)
        
        # 字号设置卡片
        size_card = CardWidget()
        size_layout = QGridLayout(size_card)
        
        # 正文字号
        body_size_label = QLabel("正文字号:")
        body_size_label.setStyleSheet(FONT_STYLES["label"])
        size_layout.addWidget(body_size_label, 0, 0)
        self.body_size_spin = SpinBox()
        self.body_size_spin.setStyleSheet(FONT_STYLES["input"])
        self.body_size_spin.setRange(8, 24)
        self.body_size_spin.setValue(12)
        size_layout.addWidget(self.body_size_spin, 0, 1)
        
        # 一级标题字号
        title1_size_label = QLabel("一级标题:")
        title1_size_label.setStyleSheet(FONT_STYLES["label"])
        size_layout.addWidget(title1_size_label, 0, 2)
        self.title1_size_spin = SpinBox()
        self.title1_size_spin.setStyleSheet(FONT_STYLES["input"])
        self.title1_size_spin.setRange(16, 32)
        self.title1_size_spin.setValue(22)
        size_layout.addWidget(self.title1_size_spin, 0, 3)
        
        # 二级标题字号
        title2_size_label = QLabel("二级标题:")
        title2_size_label.setStyleSheet(FONT_STYLES["label"])
        size_layout.addWidget(title2_size_label, 1, 0)
        self.title2_size_spin = SpinBox()
        self.title2_size_spin.setStyleSheet(FONT_STYLES["input"])
        self.title2_size_spin.setRange(14, 28)
        self.title2_size_spin.setValue(18)
        size_layout.addWidget(self.title2_size_spin, 1, 1)
        
        # 三级标题字号
        title3_size_label = QLabel("三级标题:")
        title3_size_label.setStyleSheet(FONT_STYLES["label"])
        size_layout.addWidget(title3_size_label, 1, 2)
        self.title3_size_spin = SpinBox()
        self.title3_size_spin.setStyleSheet(FONT_STYLES["input"])
        self.title3_size_spin.setRange(12, 24)
        self.title3_size_spin.setValue(16)
        size_layout.addWidget(self.title3_size_spin, 1, 3)
        
        # 页面设置卡片
        page_card = CardWidget()
        page_layout = QGridLayout(page_card)
        
        # 上边距
        margin_top_label = QLabel("上边距 (cm):")
        margin_top_label.setStyleSheet(FONT_STYLES["label"])
        page_layout.addWidget(margin_top_label, 0, 0)
        self.margin_top_spin = DoubleSpinBox()
        self.margin_top_spin.setStyleSheet(FONT_STYLES["input"])
        self.margin_top_spin.setRange(0.5, 5.0)
        self.margin_top_spin.setSingleStep(0.1)
        self.margin_top_spin.setValue(2.54)
        page_layout.addWidget(self.margin_top_spin, 0, 1)
        
        # 下边距
        margin_bottom_label = QLabel("下边距 (cm):")
        margin_bottom_label.setStyleSheet(FONT_STYLES["label"])
        page_layout.addWidget(margin_bottom_label, 0, 2)
        self.margin_bottom_spin = DoubleSpinBox()
        self.margin_bottom_spin.setStyleSheet(FONT_STYLES["input"])
        self.margin_bottom_spin.setRange(0.5, 5.0)
        self.margin_bottom_spin.setSingleStep(0.1)
        self.margin_bottom_spin.setValue(2.54)
        page_layout.addWidget(self.margin_bottom_spin, 0, 3)
        
        # 左边距
        margin_left_label = QLabel("左边距 (cm):")
        margin_left_label.setStyleSheet(FONT_STYLES["label"])
        page_layout.addWidget(margin_left_label, 1, 0)
        self.margin_left_spin = DoubleSpinBox()
        self.margin_left_spin.setStyleSheet(FONT_STYLES["input"])
        self.margin_left_spin.setRange(0.5, 5.0)
        self.margin_left_spin.setSingleStep(0.1)
        self.margin_left_spin.setValue(2.54)
        page_layout.addWidget(self.margin_left_spin, 1, 1)
        
        # 右边距
        margin_right_label = QLabel("右边距 (cm):")
        margin_right_label.setStyleSheet(FONT_STYLES["label"])
        page_layout.addWidget(margin_right_label, 1, 2)
        self.margin_right_spin = DoubleSpinBox()
        self.margin_right_spin.setStyleSheet(FONT_STYLES["input"])
        self.margin_right_spin.setRange(0.5, 5.0)
        self.margin_right_spin.setSingleStep(0.1)
        self.margin_right_spin.setValue(2.54)
        page_layout.addWidget(self.margin_right_spin, 1, 3)
        
        # 表格设置卡片
        table_card = CardWidget()
        table_layout = QGridLayout(table_card)
        
        # 表格宽度百分比
        table_width_label = QLabel("表格宽度 (%):")
        table_width_label.setStyleSheet(FONT_STYLES["label"])
        table_layout.addWidget(table_width_label, 0, 0)
        self.table_width_spin = SpinBox()
        self.table_width_spin.setStyleSheet(FONT_STYLES["input"])
        self.table_width_spin.setRange(50, 100)
        self.table_width_spin.setValue(95)
        table_layout.addWidget(self.table_width_spin, 0, 1)
        
        # 自动调整列宽
        auto_adjust_label = QLabel("自动调整列宽:")
        auto_adjust_label.setStyleSheet(FONT_STYLES["label"])
        table_layout.addWidget(auto_adjust_label, 0, 2)
        self.auto_adjust_check = CheckBox()
        self.auto_adjust_check.setStyleSheet(FONT_STYLES["label"])
        self.auto_adjust_check.setChecked(True)
        table_layout.addWidget(self.auto_adjust_check, 0, 3)
        
        # 预设管理卡片
        preset_card = CardWidget()
        preset_layout = QVBoxLayout(preset_card)
        
        # 预设管理标题
        preset_title = QLabel("预设管理")
        preset_title.setStyleSheet(FONT_STYLES["subtitle"])
        preset_layout.addWidget(preset_title)
        
        # 预设选择
        preset_select_layout = QHBoxLayout()
        preset_select_label = QLabel("选择预设:")
        preset_select_label.setStyleSheet(FONT_STYLES["label"])
        preset_select_layout.addWidget(preset_select_label)
        self.preset_manage_combo = ComboBox()
        self.preset_manage_combo.setStyleSheet(FONT_STYLES["input"])
        preset_names = ["默认", "标书专用", "紧凑", "打印就绪", "学术论文", "简历", "报告", "演示"]
        self.preset_manage_combo.addItems(preset_names)
        self.preset_manage_combo.currentIndexChanged.connect(self.on_preset_manage_change)
        preset_select_layout.addWidget(self.preset_manage_combo)
        preset_layout.addLayout(preset_select_layout)
        
        # 预设描述
        self.preset_desc_label = QLabel("选择一个预设进行编辑")
        self.preset_desc_label.setStyleSheet(FONT_STYLES["small_label"])
        preset_layout.addWidget(self.preset_desc_label)
        
        # 按钮布局
        button_layout = QHBoxLayout()
        self.save_button = PushButton("保存")
        self.save_button.setStyleSheet(FONT_STYLES["button"])
        self.save_button.clicked.connect(lambda: self.save_config(dialog))
        self.cancel_button = PushButton("取消")
        self.cancel_button.setStyleSheet(FONT_STYLES["button"])
        self.cancel_button.clicked.connect(dialog.reject)
        
        button_layout.addStretch()
        button_layout.addWidget(self.save_button)
        button_layout.addWidget(self.cancel_button)
        
        # 添加到主布局
        main_layout.addWidget(font_card)
        main_layout.addWidget(size_card)
        main_layout.addWidget(page_card)
        main_layout.addWidget(table_card)
        main_layout.addWidget(preset_card)
        main_layout.addLayout(button_layout)
        
        # 显示对话框
        dialog.exec_()
    
    def on_preset_manage_change(self, index):
        """处理预设管理选择变化"""
        from ...core.config import get_preset_config
        
        # 预设名称映射
        preset_map = {
            0: "default",
            1: "bid_document",
            2: "compact",
            3: "print_ready",
            4: "academic_paper",
            5: "resume",
            6: "report",
            7: "presentation"
        }
        
        # 获取选中的预设
        preset_key = preset_map.get(index, "default")
        config = get_preset_config(preset_key)
        
        # 更新配置值
        if 'chinese_font' in config:
            self.chinese_font_edit.setText(config['chinese_font'])
        if 'western_font' in config:
            self.western_font_edit.setText(config['western_font'])
        if 'title_font' in config:
            self.title_font_edit.setText(config['title_font'])
        if 'font_size_body' in config:
            self.body_size_spin.setValue(config['font_size_body'])
        if 'font_size_title1' in config:
            self.title1_size_spin.setValue(config['font_size_title1'])
        if 'font_size_title2' in config:
            self.title2_size_spin.setValue(config['font_size_title2'])
        if 'font_size_title3' in config:
            self.title3_size_spin.setValue(config['font_size_title3'])
        if 'page_margin_top_cm' in config:
            self.margin_top_spin.setValue(config['page_margin_top_cm'])
        if 'page_margin_bottom_cm' in config:
            self.margin_bottom_spin.setValue(config['page_margin_bottom_cm'])
        if 'page_margin_left_cm' in config:
            self.margin_left_spin.setValue(config['page_margin_left_cm'])
        if 'page_margin_right_cm' in config:
            self.margin_right_spin.setValue(config['page_margin_right_cm'])
        if 'table_width_percent' in config:
            self.table_width_spin.setValue(config['table_width_percent'])
        
        # 更新预设描述
        preset_descriptions = {
            0: "默认配置，适合大多数文档",
            1: "标书专用配置，格式更加规范",
            2: "紧凑配置，节省空间",
            3: "打印就绪配置，优化打印效果",
            4: "学术论文配置，符合学术规范",
            5: "简历配置，简洁专业",
            6: "报告配置，结构清晰",
            7: "演示配置，突出重点"
        }
        
        self.preset_desc_label.setText(preset_descriptions.get(index, "选择一个预设进行编辑"))
    
    def save_config(self, dialog):
        """保存配置"""
        # 这里可以添加保存配置的逻辑
        # 目前只是关闭对话框
        dialog.accept()
        self.log("配置已保存")
    
    def log(self, message):
        """记录日志"""
        self.log_edit.append(message)
    
    def on_show_event(self, event):
        """窗口显示事件，打印UI元素位置和大小"""
        print("\n=== UI元素位置和大小 ===")
        # 打印文件列表信息
        if hasattr(self, 'file_list'):
            file_list_geo = self.file_list.geometry()
            print(f"文件列表位置: ({file_list_geo.x()}, {file_list_geo.y()})")
            print(f"文件列表大小: {file_list_geo.width()}x{file_list_geo.height()}")
        
        # 打印输出编辑框信息
        if hasattr(self, 'output_edit'):
            output_edit_geo = self.output_edit.geometry()
            print(f"输出编辑框位置: ({output_edit_geo.x()}, {output_edit_geo.y()})")
            print(f"输出编辑框大小: {output_edit_geo.width()}x{output_edit_geo.height()}")
        
        # 打印输出按钮信息
        if hasattr(self, 'output_button'):
            output_button_geo = self.output_button.geometry()
            print(f"输出按钮位置: ({output_button_geo.x()}, {output_button_geo.y()})")
            print(f"输出按钮大小: {output_button_geo.width()}x{output_button_geo.height()}")
        
        # 打印批量模式复选框信息
        if hasattr(self, 'batch_mode'):
            batch_mode_geo = self.batch_mode.geometry()
            print(f"批量模式复选框位置: ({batch_mode_geo.x()}, {batch_mode_geo.y()})")
            print(f"批量模式复选框大小: {batch_mode_geo.width()}x{batch_mode_geo.height()}")
        
        # 打印输入编辑框信息
        if hasattr(self, 'input_edit'):
            input_edit_geo = self.input_edit.geometry()
            print(f"输入编辑框位置: ({input_edit_geo.x()}, {input_edit_geo.y()})")
            print(f"输入编辑框大小: {input_edit_geo.width()}x{input_edit_geo.height()}")
    
    def mousePressEvent(self, event):
        """鼠标点击事件，打印点击位置和点击的UI元素信息"""
        # 获取点击位置
        pos = event.pos()
        print(f"\n=== 鼠标点击信息 ===")
        print(f"点击位置: ({pos.x()}, {pos.y()})")
        
        # 检查点击是否在文件列表上
        if hasattr(self, 'file_list'):
            file_list_geo = self.file_list.geometry()
            if file_list_geo.contains(pos):
                print(f"点击了文件列表:")
                print(f"  位置: ({file_list_geo.x()}, {file_list_geo.y()})")
                print(f"  大小: {file_list_geo.width()}x{file_list_geo.height()}")
                return
        
        # 检查点击是否在输出编辑框上
        if hasattr(self, 'output_edit'):
            output_edit_geo = self.output_edit.geometry()
            if output_edit_geo.contains(pos):
                print(f"点击了输出编辑框:")
                print(f"  位置: ({output_edit_geo.x()}, {output_edit_geo.y()})")
                print(f"  大小: {output_edit_geo.width()}x{output_edit_geo.height()}")
                return
        
        # 检查点击是否在输出按钮上
        if hasattr(self, 'output_button'):
            output_button_geo = self.output_button.geometry()
            if output_button_geo.contains(pos):
                print(f"点击了输出按钮:")
                print(f"  位置: ({output_button_geo.x()}, {output_button_geo.y()})")
                print(f"  大小: {output_button_geo.width()}x{output_button_geo.height()}")
                return
        
        # 检查点击是否在批量模式复选框上
        if hasattr(self, 'batch_mode'):
            batch_mode_geo = self.batch_mode.geometry()
            if batch_mode_geo.contains(pos):
                print(f"点击了批量模式复选框:")
                print(f"  位置: ({batch_mode_geo.x()}, {batch_mode_geo.y()})")
                print(f"  大小: {batch_mode_geo.width()}x{batch_mode_geo.height()}")
                return
        
        # 检查点击是否在输入编辑框上
        if hasattr(self, 'input_edit'):
            input_edit_geo = self.input_edit.geometry()
            if input_edit_geo.contains(pos):
                print(f"点击了输入编辑框:")
                print(f"  位置: ({input_edit_geo.x()}, {input_edit_geo.y()})")
                print(f"  大小: {input_edit_geo.width()}x{input_edit_geo.height()}")
                return
        
        # 检查点击是否在输入按钮上
        if hasattr(self, 'input_button'):
            input_button_geo = self.input_button.geometry()
            if input_button_geo.contains(pos):
                print(f"点击了输入按钮:")
                print(f"  位置: ({input_button_geo.x()}, {input_button_geo.y()})")
                print(f"  大小: {input_button_geo.width()}x{input_button_geo.height()}")
                return
        
        # 检查点击是否在批量文件按钮上
        if hasattr(self, 'batch_files_button'):
            batch_files_button_geo = self.batch_files_button.geometry()
            if batch_files_button_geo.contains(pos):
                print(f"点击了批量文件按钮:")
                print(f"  位置: ({batch_files_button_geo.x()}, {batch_files_button_geo.y()})")
                print(f"  大小: {batch_files_button_geo.width()}x{batch_files_button_geo.height()}")
                return
        
        # 检查点击是否在批量文件夹按钮上
        if hasattr(self, 'batch_folder_button'):
            batch_folder_button_geo = self.batch_folder_button.geometry()
            if batch_folder_button_geo.contains(pos):
                print(f"点击了批量文件夹按钮:")
                print(f"  位置: ({batch_folder_button_geo.x()}, {batch_folder_button_geo.y()})")
                print(f"  大小: {batch_folder_button_geo.width()}x{batch_folder_button_geo.height()}")
                return
        
        # 检查点击是否在预设组合框上
        if hasattr(self, 'preset_combo'):
            preset_combo_geo = self.preset_combo.geometry()
            if preset_combo_geo.contains(pos):
                print(f"点击了预设组合框:")
                print(f"  位置: ({preset_combo_geo.x()}, {preset_combo_geo.y()})")
                print(f"  大小: {preset_combo_geo.width()}x{preset_combo_geo.height()}")
                return
        
        # 如果没有点击任何UI元素，打印点击位置
        print("未点击任何UI元素")


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
