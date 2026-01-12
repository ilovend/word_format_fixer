"""Word文档修复工具GUI界面"""

import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
import sys
from ..core.fixer import RobustWordFixer
from ..core.config import get_preset_config


class WordFixerApp:
    """Word文档修复工具GUI应用"""
    
    def __init__(self, root):
        """
        初始化应用
        
        Args:
            root: tkinter根窗口
        """
        self.root = root
        self.root.title("Word文档格式修复工具")
        self.root.geometry("800x800")
        self.root.resizable(True, True)
        
        # 设置最小窗口大小
        self.root.minsize(600, 600)
        
        # 设置主题和样式
        self.style = ttk.Style()
        if sys.platform == 'darwin':  # macOS
            self.style.theme_use('aqua')
        else:
            self.style.theme_use('clam')
        
        # 自定义样式
        self.style.configure('TLabelFrame',
                           borderwidth=2,
                           relief='groove',
                           padding=10)
        
        self.style.configure('TButton',
                           padding=6,
                           font=('微软雅黑', 10))
        
        self.style.configure('TLabelframe.Label',
                           font=('微软雅黑', 11, 'bold'),
                           foreground='#333333')
        
        self.style.configure('TLabel',
                           font=('微软雅黑', 10))
        
        self.style.configure('TEntry',
                           padding=4,
                           font=('微软雅黑', 10))
        
        # 变量
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.preset = tk.StringVar(value='default')
        
        # 字体设置
        self.chinese_font = tk.StringVar(value='宋体')
        self.western_font = tk.StringVar(value='Arial')
        self.title_font = tk.StringVar(value='黑体')
        
        # 字号设置
        self.font_size_body = tk.IntVar(value=12)
        self.font_size_title1 = tk.IntVar(value=22)
        self.font_size_title2 = tk.IntVar(value=18)
        self.font_size_title3 = tk.IntVar(value=16)
        
        # 页面设置
        self.margin_top = tk.DoubleVar(value=2.54)
        self.margin_bottom = tk.DoubleVar(value=2.54)
        self.margin_left = tk.DoubleVar(value=2.54)
        self.margin_right = tk.DoubleVar(value=2.54)
        
        # 表格设置
        self.table_width_percent = tk.IntVar(value=95)
        self.auto_adjust_columns = tk.BooleanVar(value=True)
        
        # 创建主框架
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 添加标题
        title_frame = ttk.Frame(self.main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 20))
        
        title_label = ttk.Label(title_frame, text="Word文档格式修复工具", font=('微软雅黑', 16, 'bold'), foreground='#0066cc')
        title_label.pack(side=tk.LEFT)
        
        version_label = ttk.Label(title_frame, text="v1.0.0", font=('微软雅黑', 10, 'italic'), foreground='#666666')
        version_label.pack(side=tk.RIGHT, padx=10)
        
        # 创建滚动窗口
        self.create_scrollable_window()
        
        # 创建UI组件
        self.create_widgets()

        # 快捷键（Ctrl+R 开始修复，Ctrl+O 选择输入文件，Ctrl+Q 退出）
        self.root.bind('<Control-r>', lambda e: self.fix_document())
        self.root.bind('<Control-o>', lambda e: self.browse_input_file())
        self.root.bind('<Control-q>', lambda e: self.root.quit())
    
    def create_scrollable_window(self):
        """创建可滚动窗口"""
        # 创建滚动条容器
        self.scroll_frame = ttk.Frame(self.main_frame)
        self.scroll_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建画布和滚动条
        self.canvas = tk.Canvas(self.scroll_frame)
        self.scrollbar = ttk.Scrollbar(self.scroll_frame, orient="vertical", command=self.canvas.yview)
        
        # 创建内部框架
        self.inner_frame = ttk.Frame(self.canvas)
        
        # 绑定滚动事件
        self.inner_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )
        
        # 创建窗口（保存 window id 以便响应式调整）
        self.window_id = self.canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")

        # 当 canvas 大小改变时，调整内部窗口宽度并尽量居中（避免右侧空白或左侧挤压）
        def _on_canvas_configure(e):
            # 保留左右边距 20px，总最大宽度 900px
            max_content_width = 900
            padding = 20
            available = max(padding, e.width - padding * 2)
            content_width = min(available, max_content_width)

            # 设置内部窗口宽度
            self.canvas.itemconfig(self.window_id, width=content_width)

            # 计算左上角 x，使内容水平居中
            x = max(padding, (e.width - content_width) // 2)
            # 更新窗口坐标（anchor=NW）
            self.canvas.coords(self.window_id, x, 0)

            # 更新状态条上的尺寸显示（若存在）
            if hasattr(self, 'size_label'):
                self.size_label.config(
                    text=f"窗口: {e.width}×{e.height}，内容宽度: {content_width}"
                )

        self.canvas.bind("<Configure>", _on_canvas_configure)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # 布局
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 当鼠标进入/离开内部区域时，绑定/解绑鼠标滚轮以支持滚轮滚动（Windows/macOS/Linux）
        self.inner_frame.bind("<Enter>", lambda e: self._bind_mousewheel())
        self.inner_frame.bind("<Leave>", lambda e: self._unbind_mousewheel())
    
    def _on_mousewheel(self, event):
        """统一处理鼠标滚轮事件，支持 Windows/macOS/Linux"""
        # Linux 鼠标滚动使用 Button-4/5
        if hasattr(event, 'num'):
            if event.num == 4:
                self.canvas.yview_scroll(-1, 'units')
                return
            elif event.num == 5:
                self.canvas.yview_scroll(1, 'units')
                return

        # Windows 和 macOS 使用 delta（Windows delta 为 120 的倍数）
        if hasattr(event, 'delta'):
            try:
                # 以 120 为单位缩放到通用步进
                step = int(-1 * (event.delta / 120))
            except Exception:
                step = -1 if event.delta > 0 else 1
            if step == 0:
                step = -1 if event.delta > 0 else 1
            self.canvas.yview_scroll(step, 'units')

    def _bind_mousewheel(self):
        """在鼠标进入区域时绑定滚轮事件"""
        # 使用 bind_all 可以确保在光标在画布上时捕获滚轮事件
        if sys.platform.startswith('linux'):
            self.canvas.bind_all('<Button-4>', self._on_mousewheel)
            self.canvas.bind_all('<Button-5>', self._on_mousewheel)
        else:
            self.canvas.bind_all('<MouseWheel>', self._on_mousewheel)

    def _unbind_mousewheel(self):
        """在鼠标离开区域时解绑滚轮事件"""
        if sys.platform.startswith('linux'):
            self.canvas.unbind_all('<Button-4>')
            self.canvas.unbind_all('<Button-5>')
        else:
            self.canvas.unbind_all('<MouseWheel>')

    def create_widgets(self):
        """创建UI组件"""
        # 文件选择部分
        file_frame = ttk.LabelFrame(self.inner_frame, text="文件选择", padding="10")
        file_frame.pack(fill=tk.X, pady=5)
        
        # 输入文件
        input_frame = ttk.Frame(file_frame)
        input_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(input_frame, text="输入文件:", width=10).pack(side=tk.LEFT, padx=5)
        ttk.Entry(input_frame, textvariable=self.input_file).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(input_frame, text="浏览", command=self.browse_input_file, width=10).pack(side=tk.LEFT, padx=5)
        
        # 输出文件
        output_frame = ttk.Frame(file_frame)
        output_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(output_frame, text="输出文件:", width=10).pack(side=tk.LEFT, padx=5)
        ttk.Entry(output_frame, textvariable=self.output_file).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(output_frame, text="浏览", command=self.browse_output_file, width=10).pack(side=tk.LEFT, padx=5)
        
        # 提示信息
        hint_frame = ttk.Frame(file_frame)
        hint_frame.pack(fill=tk.X, pady=5)
        
        hint_label = ttk.Label(hint_frame, text="提示: 选择输入文件后，输出文件路径会自动填充", font=('微软雅黑', 9, 'italic'), foreground='#666666')
        hint_label.pack(side=tk.LEFT, padx=5)
        
        # 配置部分
        config_frame = ttk.LabelFrame(self.inner_frame, text="配置选项", padding="10")
        config_frame.pack(fill=tk.X, pady=5)
        
        # 预设配置
        preset_frame = ttk.Frame(config_frame)
        preset_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(preset_frame, text="预设配置:", width=12).pack(side=tk.LEFT, padx=5)
        preset_combo = ttk.Combobox(preset_frame, textvariable=self.preset, values=['default', 'bid_document', 'compact', 'print_ready'], state='readonly', width=20)
        preset_combo.pack(side=tk.LEFT, padx=5)
        preset_combo.bind('<<ComboboxSelected>>', self.on_preset_change)
        
        # 预设配置说明
        preset_hint = ttk.Label(preset_frame, text="(默认/标书专用/紧凑/打印就绪)", font=('微软雅黑', 9, 'italic'), foreground='#666666')
        preset_hint.pack(side=tk.LEFT, padx=5)
        
        # 字体设置
        font_frame = ttk.LabelFrame(config_frame, text="字体设置", padding="10")
        font_frame.pack(fill=tk.X, pady=5)
        
        font_row1 = ttk.Frame(font_frame)
        font_row1.pack(fill=tk.X, pady=5)
        
        ttk.Label(font_row1, text="中文字体:", width=12).pack(side=tk.LEFT, padx=5)
        ttk.Entry(font_row1, textvariable=self.chinese_font, width=15).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(font_row1, text="西文字体:", width=12).pack(side=tk.LEFT, padx=5)
        ttk.Entry(font_row1, textvariable=self.western_font, width=15).pack(side=tk.LEFT, padx=5)
        
        font_row2 = ttk.Frame(font_frame)
        font_row2.pack(fill=tk.X, pady=5)
        
        ttk.Label(font_row2, text="标题字体:", width=12).pack(side=tk.LEFT, padx=5)
        ttk.Entry(font_row2, textvariable=self.title_font, width=15).pack(side=tk.LEFT, padx=5)
        
        # 字号设置
        size_frame = ttk.LabelFrame(config_frame, text="字号设置 (磅)", padding="10")
        size_frame.pack(fill=tk.X, pady=5)
        
        size_row1 = ttk.Frame(size_frame)
        size_row1.pack(fill=tk.X, pady=5)
        
        ttk.Label(size_row1, text="正文字号:", width=12).pack(side=tk.LEFT, padx=5)
        ttk.Entry(size_row1, textvariable=self.font_size_body, width=10).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(size_row1, text="一级标题:", width=12).pack(side=tk.LEFT, padx=5)
        ttk.Entry(size_row1, textvariable=self.font_size_title1, width=10).pack(side=tk.LEFT, padx=5)
        
        size_row2 = ttk.Frame(size_frame)
        size_row2.pack(fill=tk.X, pady=5)
        
        ttk.Label(size_row2, text="二级标题:", width=12).pack(side=tk.LEFT, padx=5)
        ttk.Entry(size_row2, textvariable=self.font_size_title2, width=10).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(size_row2, text="三级标题:", width=12).pack(side=tk.LEFT, padx=5)
        ttk.Entry(size_row2, textvariable=self.font_size_title3, width=10).pack(side=tk.LEFT, padx=5)
        
        # 页面设置
        page_frame = ttk.LabelFrame(config_frame, text="页面设置 (厘米)", padding="10")
        page_frame.pack(fill=tk.X, pady=5)
        
        page_row1 = ttk.Frame(page_frame)
        page_row1.pack(fill=tk.X, pady=5)
        
        ttk.Label(page_row1, text="上边距:", width=12).pack(side=tk.LEFT, padx=5)
        ttk.Entry(page_row1, textvariable=self.margin_top, width=10).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(page_row1, text="下边距:", width=12).pack(side=tk.LEFT, padx=5)
        ttk.Entry(page_row1, textvariable=self.margin_bottom, width=10).pack(side=tk.LEFT, padx=5)
        
        page_row2 = ttk.Frame(page_frame)
        page_row2.pack(fill=tk.X, pady=5)
        
        ttk.Label(page_row2, text="左边距:", width=12).pack(side=tk.LEFT, padx=5)
        ttk.Entry(page_row2, textvariable=self.margin_left, width=10).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(page_row2, text="右边距:", width=12).pack(side=tk.LEFT, padx=5)
        ttk.Entry(page_row2, textvariable=self.margin_right, width=10).pack(side=tk.LEFT, padx=5)
        
        # 表格设置
        table_frame = ttk.LabelFrame(config_frame, text="表格设置", padding="10")
        table_frame.pack(fill=tk.X, pady=5)
        
        table_row1 = ttk.Frame(table_frame)
        table_row1.pack(fill=tk.X, pady=5)
        
        ttk.Label(table_row1, text="表格宽度 (%):", width=12).pack(side=tk.LEFT, padx=5)
        ttk.Entry(table_row1, textvariable=self.table_width_percent, width=10).pack(side=tk.LEFT, padx=5)
        
        table_row2 = ttk.Frame(table_frame)
        table_row2.pack(fill=tk.X, pady=5)
        
        ttk.Label(table_row2, text="自动调整列宽:", width=12).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(table_row2, variable=self.auto_adjust_columns).pack(side=tk.LEFT, padx=5)
        
        # 操作按钮
        button_frame = ttk.Frame(self.inner_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        # 居中按钮容器
        button_center_frame = ttk.Frame(button_frame)
        button_center_frame.pack(fill=tk.X)
        
        # 开始修复按钮（突出显示）
        self.style.configure('Primary.TButton',
                           background='#0066cc',
                           foreground='white',
                           padding=8,
                           font=('微软雅黑', 11, 'bold'))
        
        self.fix_button = ttk.Button(button_center_frame, text="开始修复", command=self.fix_document, style='Primary.TButton')
        self.fix_button.pack(side=tk.LEFT, padx=10, ipady=5)

        # 重置按钮（恢复为当前预设）
        self.reset_button = ttk.Button(button_center_frame, text="重置", command=self.reset_to_preset)
        self.reset_button.pack(side=tk.LEFT, padx=10, ipady=5)
        
        # 退出按钮
        self.cancel_button = ttk.Button(button_center_frame, text="退出", command=self.root.quit)
        self.cancel_button.pack(side=tk.LEFT, padx=10, ipady=5) 
        
        # 提示信息
        tip_frame = ttk.Frame(self.inner_frame)
        tip_frame.pack(fill=tk.X, pady=5)
        
        tip_label = ttk.Label(tip_frame, text="提示: 修复过程中请耐心等待，不要关闭窗口", font=('微软雅黑', 9, 'italic'), foreground='#666666')
        tip_label.pack(anchor=tk.CENTER)
        
        # 状态显示
        self.status_var = tk.StringVar(value="就绪")
        status_frame = ttk.LabelFrame(self.inner_frame, text="状态", padding="10")
        status_frame.pack(fill=tk.X, pady=5)
        
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var, font=('微软雅黑', 10))
        self.status_label.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 窗口尺寸显示（用于调试响应式布局）
        self.size_label = ttk.Label(status_frame, text="窗口: --×--，内容宽度: --", font=('微软雅黑', 9), foreground='#666666')
        self.size_label.pack(anchor='e', padx=10, pady=(0,6))
        
        # 日志文本框
        log_frame = ttk.LabelFrame(self.inner_frame, text="操作日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 日志文本和滚动条（滚动条放在 log_frame 中，避免将其附着于 Text 本身）
        self.log_text = tk.Text(
            log_frame, 
            height=10, 
            wrap=tk.WORD, 
            font=('微软雅黑', 10)
        )
        self.log_scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5, padx=(0,5))
        self.log_text.config(yscrollcommand=self.log_scrollbar.set)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=(5,0), pady=5)
    
    def browse_input_file(self):
        """浏览输入文件"""
        filetypes = [('Word文档', '*.docx'), ('所有文件', '*.*')]
        filename = filedialog.askopenfilename(title="选择输入文件", filetypes=filetypes)
        if filename:
            self.input_file.set(filename)
            # 自动设置输出文件路径
            if not self.output_file.get():
                base, ext = os.path.splitext(filename)
                output_filename = f"{base}_fixed{ext}"
                self.output_file.set(output_filename)
    
    def browse_output_file(self):
        """浏览输出文件"""
        filetypes = [('Word文档', '*.docx'), ('所有文件', '*.*')]
        filename = filedialog.asksaveasfilename(title="选择输出文件", filetypes=filetypes, defaultextension=".docx")
        if filename:
            self.output_file.set(filename)
    
    def on_preset_change(self, event):
        """预设配置改变时的处理"""
        preset = self.preset.get()
        config = get_preset_config(preset)
        
        # 更新字体设置
        if 'chinese_font' in config:
            self.chinese_font.set(config['chinese_font'])
        if 'western_font' in config:
            self.western_font.set(config['western_font'])
        if 'title_font' in config:
            self.title_font.set(config['title_font'])
        
        # 更新字号设置
        if 'font_size_body' in config:
            self.font_size_body.set(config['font_size_body'])
        if 'font_size_title1' in config:
            self.font_size_title1.set(config['font_size_title1'])
        if 'font_size_title2' in config:
            self.font_size_title2.set(config['font_size_title2'])
        if 'font_size_title3' in config:
            self.font_size_title3.set(config['font_size_title3'])
        
        # 更新页面设置
        if 'page_margin_top_cm' in config:
            self.margin_top.set(config['page_margin_top_cm'])
        if 'page_margin_bottom_cm' in config:
            self.margin_bottom.set(config['page_margin_bottom_cm'])
        if 'page_margin_left_cm' in config:
            self.margin_left.set(config['page_margin_left_cm'])
        if 'page_margin_right_cm' in config:
            self.margin_right.set(config['page_margin_right_cm'])
        
        # 更新表格设置
        if 'table_width_percent' in config:
            self.table_width_percent.set(config['table_width_percent'])
        
        # 显示预设信息
        self.log(f"已选择预设配置: {preset}")

    def reset_to_preset(self):
        """将界面设置恢复到当前选择的预设配置（或默认）"""
        preset = self.preset.get()
        config = get_preset_config(preset)

        # 更新字体设置
        if 'chinese_font' in config:
            self.chinese_font.set(config['chinese_font'])
        if 'western_font' in config:
            self.western_font.set(config['western_font'])
        if 'title_font' in config:
            self.title_font.set(config['title_font'])

        # 更新字号设置
        if 'font_size_body' in config:
            self.font_size_body.set(config['font_size_body'])
        if 'font_size_title1' in config:
            self.font_size_title1.set(config['font_size_title1'])
        if 'font_size_title2' in config:
            self.font_size_title2.set(config['font_size_title2'])
        if 'font_size_title3' in config:
            self.font_size_title3.set(config['font_size_title3'])

        # 更新页面设置
        if 'page_margin_top_cm' in config:
            self.margin_top.set(config['page_margin_top_cm'])
        if 'page_margin_bottom_cm' in config:
            self.margin_bottom.set(config['page_margin_bottom_cm'])
        if 'page_margin_left_cm' in config:
            self.margin_left.set(config['page_margin_left_cm'])
        if 'page_margin_right_cm' in config:
            self.margin_right.set(config['page_margin_right_cm'])

        # 更新表格设置
        if 'table_width_percent' in config:
            self.table_width_percent.set(config['table_width_percent'])
        if 'auto_adjust_columns' in config:
            self.auto_adjust_columns.set(config.get('auto_adjust_columns', True))

        self.log(f"已恢复预设配置: {preset}")    
    def fix_document(self):
        """开始修复文档"""
        input_file = self.input_file.get()
        output_file = self.output_file.get()
        
        # 验证输入
        if not input_file:
            messagebox.showerror("错误", "请选择输入文件")
            return
        
        if not output_file:
            messagebox.showerror("错误", "请选择输出文件")
            return
        
        if not os.path.exists(input_file):
            messagebox.showerror("错误", "输入文件不存在")
            return
        
        try:
            # 更新状态
            self.status_var.set("正在修复...")
            self.fix_button.config(state=tk.DISABLED)
            self.root.update()
            
            # 清空日志
            self.log_text.delete(1.0, tk.END)
            
            # 构建配置
            config = get_preset_config(self.preset.get())
            config.update({
                # 字体设置
                'chinese_font': self.chinese_font.get(),
                'western_font': self.western_font.get(),
                'title_font': self.title_font.get(),
                # 字号设置
                'font_size_body': self.font_size_body.get(),
                'font_size_title1': self.font_size_title1.get(),
                'font_size_title2': self.font_size_title2.get(),
                'font_size_title3': self.font_size_title3.get(),
                # 页面设置
                'page_margin_top_cm': self.margin_top.get(),
                'page_margin_bottom_cm': self.margin_bottom.get(),
                'page_margin_left_cm': self.margin_left.get(),
                'page_margin_right_cm': self.margin_right.get(),
                # 表格设置
                'table_width_percent': self.table_width_percent.get(),
                'auto_adjust_columns': self.auto_adjust_columns.get(),
            })
            
            # 创建修复器并执行修复
            fixer = RobustWordFixer(config)
            
            # 重定向标准输出到日志
            old_stdout = sys.stdout
            class LogRedirector:
                def __init__(self, text_widget):
                    self.text_widget = text_widget
                def write(self, text):
                    self.text_widget.insert(tk.END, text)
                    self.text_widget.see(tk.END)
                def flush(self):
                    pass
            
            sys.stdout = LogRedirector(self.log_text)
            
            # 执行修复
            result = fixer.fix_all(input_file, output_file)
            
            # 恢复标准输出
            sys.stdout = old_stdout
            
            if result:
                self.status_var.set("修复完成")
                messagebox.showinfo("成功", f"文档修复完成！\n输出文件: {result}")
            else:
                self.status_var.set("修复失败")
                messagebox.showerror("错误", "文档修复失败，请查看日志")
                
        except Exception as e:
            self.status_var.set("错误")
            self.log(f"错误: {str(e)}")
            messagebox.showerror("错误", f"发生错误: {str(e)}")
        finally:
            self.fix_button.config(state=tk.NORMAL)
            self.root.update()
    
    def log(self, message):
        """记录日志"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)


def run_app():
    """运行应用"""
    root = tk.Tk()
    WordFixerApp(root)
    root.mainloop()


if __name__ == "__main__":
    run_app()
