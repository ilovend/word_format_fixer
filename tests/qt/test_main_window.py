"""Word文档格式修复工具 - PyQt-Fluent-Widgets界面单元测试"""

import unittest
import sys
import os
from unittest.mock import Mock, patch

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))

from word_format_fixer.ui.qt.main_window import HomePage, WordFixerWindow
from PyQt5.QtWidgets import QApplication


class TestWordFixerWindow(unittest.TestCase):
    """测试WordFixerWindow类"""
    
    @classmethod
    def setUpClass(cls):
        """在所有测试方法执行前创建QApplication实例"""
        # 确保只创建一个QApplication实例
        if not QApplication.instance():
            cls.app = QApplication(sys.argv)
        else:
            cls.app = QApplication.instance()
    
    def setUp(self):
        """在每个测试方法执行前创建测试对象"""
        self.window = WordFixerWindow()
    
    def tearDown(self):
        """在每个测试方法执行后清理测试对象"""
        if self.window:
            self.window.close()
    
    def test_window_initialization(self):
        """测试窗口初始化"""
        self.assertIsNotNone(self.window)
        self.assertEqual(self.window.windowTitle(), "Word文档格式修复工具")
        self.assertEqual(self.window.size().width(), 900)
        self.assertEqual(self.window.size().height(), 700)
    
    def test_home_page_creation(self):
        """测试主页面创建"""
        self.assertIsNotNone(self.window.home_page)
        self.assertEqual(self.window.home_page.objectName(), "HomePage")
    
    def test_shortcuts_setup(self):
        """测试快捷键设置"""
        # 快捷键设置在setup_ui中完成，这里主要测试窗口是否能正常初始化
        self.assertIsNotNone(self.window)


class TestHomePage(unittest.TestCase):
    """测试HomePage类"""
    
    @classmethod
    def setUpClass(cls):
        """在所有测试方法执行前创建QApplication实例"""
        if not QApplication.instance():
            cls.app = QApplication(sys.argv)
        else:
            cls.app = QApplication.instance()
    
    def setUp(self):
        """在每个测试方法执行前创建测试对象"""
        self.page = HomePage()
    
    def tearDown(self):
        """在每个测试方法执行后清理测试对象"""
        if self.page:
            self.page.deleteLater()
    
    def test_page_initialization(self):
        """测试页面初始化"""
        self.assertIsNotNone(self.page)
        self.assertEqual(self.page.objectName(), "HomePage")
    
    def test_batch_mode_initialization(self):
        """测试批量模式初始化"""
        self.assertIsNotNone(self.page.batch_mode)
        self.assertFalse(self.page.batch_mode.isChecked())
    
    def test_preset_combo_initialization(self):
        """测试预设配置下拉框初始化"""
        self.assertIsNotNone(self.page.preset_combo)
        self.assertEqual(self.page.preset_combo.currentText(), "默认")
        self.assertEqual(self.page.preset_combo.count(), 8)
    
    def test_file_list_initialization(self):
        """测试文件列表初始化"""
        self.assertIsNotNone(self.page.file_list)
        self.assertEqual(self.page.file_list.count(), 0)
    
    @patch('word_format_fixer.ui.qt.main_window.QFileDialog')
    def test_browse_input_file(self, mock_file_dialog):
        """测试浏览输入文件功能"""
        # 模拟文件对话框返回值
        mock_dialog = Mock()
        mock_dialog.exec_.return_value = True
        mock_dialog.selectedFiles.return_value = ["test_input.docx"]
        mock_file_dialog.return_value = mock_dialog
        
        # 执行浏览输入文件方法
        self.page.browse_input_file()
        
        # 验证文件对话框被调用
        mock_file_dialog.assert_called_once()
        # 验证输入文件路径被设置
        self.assertEqual(self.page.input_edit.text(), "test_input.docx")
        # 验证输出文件路径被自动设置
        self.assertEqual(self.page.output_edit.text(), "test_input_fixed.docx")
    
    @patch('word_format_fixer.ui.qt.main_window.QFileDialog')
    def test_browse_output_file(self, mock_file_dialog):
        """测试浏览输出文件功能"""
        # 模拟文件对话框返回值
        mock_dialog = Mock()
        mock_dialog.exec_.return_value = True
        mock_dialog.selectedFiles.return_value = ["test_output.docx"]
        mock_file_dialog.return_value = mock_dialog
        
        # 执行浏览输出文件方法
        self.page.browse_output_file()
        
        # 验证文件对话框被调用
        mock_file_dialog.assert_called_once()
        # 验证输出文件路径被设置
        self.assertEqual(self.page.output_edit.text(), "test_output.docx")
    
    @patch('word_format_fixer.ui.qt.main_window.QFileDialog')
    def test_browse_batch_files(self, mock_file_dialog):
        """测试浏览批量文件功能"""
        # 模拟文件对话框返回值
        mock_dialog = Mock()
        mock_dialog.exec_.return_value = True
        mock_dialog.selectedFiles.return_value = ["file1.docx", "file2.docx"]
        mock_file_dialog.return_value = mock_dialog
        
        # 执行浏览批量文件方法
        self.page.browse_batch_files()
        
        # 验证文件对话框被调用
        mock_file_dialog.assert_called_once()
        # 验证文件列表被填充
        self.assertEqual(self.page.file_list.count(), 2)
        self.assertEqual(self.page.file_list.item(0).text(), "file1.docx")
        self.assertEqual(self.page.file_list.item(1).text(), "file2.docx")
    
    def test_reset_settings(self):
        """测试重置设置功能"""
        # 先设置一些值
        self.page.input_edit.setText("test_input.docx")
        self.page.output_edit.setText("test_output.docx")
        self.page.batch_mode.setChecked(True)
        self.page.preset_combo.setCurrentIndex(1)
        
        # 执行重置设置方法
        self.page.reset_settings()
        
        # 验证所有设置被重置
        self.assertEqual(self.page.input_edit.text(), "")
        self.assertEqual(self.page.output_edit.text(), "")
        self.assertFalse(self.page.batch_mode.isChecked())
        self.assertEqual(self.page.preset_combo.currentIndex(), 0)
        self.assertEqual(self.page.file_list.count(), 0)
    
    def test_config_button_connection(self):
        """测试配置按钮的点击事件连接"""
        # 验证配置按钮存在
        self.assertIsNotNone(self.page.config_button)
        # 验证配置按钮文本
        self.assertEqual(self.page.config_button.text(), "配置")
    
    def test_open_config_dialog(self):
        """测试打开配置对话框功能"""
        # 模拟模块级别的QDialog和方法内部导入的其他组件
        with patch('word_format_fixer.ui.qt.main_window.QDialog') as mock_dialog, \
             patch('PyQt5.QtWidgets.QVBoxLayout') as mock_vbox, \
             patch('PyQt5.QtWidgets.QHBoxLayout') as mock_hbox, \
             patch('PyQt5.QtWidgets.QGridLayout') as mock_grid, \
             patch('PyQt5.QtWidgets.QPushButton') as mock_button, \
             patch('qfluentwidgets.CardWidget') as mock_card, \
             patch('qfluentwidgets.LineEdit') as mock_line, \
             patch('qfluentwidgets.ComboBox') as mock_combo, \
             patch('qfluentwidgets.SpinBox') as mock_spin, \
             patch('qfluentwidgets.DoubleSpinBox') as mock_double, \
             patch('qfluentwidgets.CheckBox') as mock_check:
            
            # 模拟对话框
            mock_instance = Mock()
            mock_dialog.return_value = mock_instance
            
            # 模拟布局
            mock_layout_instance = Mock()
            mock_vbox.return_value = mock_layout_instance
            mock_hbox.return_value = mock_layout_instance
            mock_grid.return_value = mock_layout_instance
            
            # 执行打开配置对话框方法
            self.page.open_config_dialog()
            
            # 验证对话框被创建
            mock_dialog.assert_called_once()
            # 验证对话框的exec_方法被调用
            mock_instance.exec_.assert_called_once()
    
    def test_save_config(self):
        """测试保存配置功能"""
        # 模拟对话框
        mock_dialog = Mock()
        
        # 执行保存配置方法
        self.page.save_config(mock_dialog)
        
        # 验证对话框的accept方法被调用
        mock_dialog.accept.assert_called_once()
    
    @patch('word_format_fixer.ui.qt.main_window.RobustWordFixer')
    def test_fix_document_single_file(self, mock_fixer):
        """测试单文件模式下的修复文档功能"""
        # 模拟修复器
        mock_instance = Mock()
        mock_instance.fix_all.return_value = "test_output_fixed.docx"
        mock_fixer.return_value = mock_instance
        
        # 设置输入输出文件路径
        self.page.input_edit.setText("test_input.docx")
        self.page.output_edit.setText("test_output.docx")
        
        # 模拟文件存在
        with patch('os.path.exists', return_value=True):
            # 执行修复文档方法
            self.page.fix_document()
            
            # 验证修复器被调用
            mock_fixer.assert_called_once()
            # 验证fix_all方法被调用
            mock_instance.fix_all.assert_called_once_with("test_input.docx", "test_output.docx")
    
    @patch('word_format_fixer.ui.qt.main_window.RobustWordFixer')
    def test_fix_document_batch_mode(self, mock_fixer):
        """测试批量模式下的修复文档功能"""
        # 模拟修复器
        mock_instance = Mock()
        mock_instance.fix_batch.return_value = {
            "file1.docx": "file1_fixed.docx",
            "file2.docx": "file2_fixed.docx"
        }
        mock_fixer.return_value = mock_instance
        
        # 设置批量模式和文件列表
        self.page.batch_mode.setChecked(True)
        self.page.file_list.addItem("file1.docx")
        self.page.file_list.addItem("file2.docx")
        
        # 执行修复文档方法
        self.page.fix_document()
        
        # 验证修复器被调用
        mock_fixer.assert_called_once()
        # 验证fix_batch方法被调用
        mock_instance.fix_batch.assert_called_once()


if __name__ == '__main__':
    unittest.main()
