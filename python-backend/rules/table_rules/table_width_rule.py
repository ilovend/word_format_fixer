from rules.base_rule import BaseRule, RuleResult
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Cm
import re

class TableWidthRule(BaseRule):
    """表格宽度规则 - 优化表格宽度，支持复杂表格结构"""

    display_name = "表格宽度优化"
    category = "表格规则"
    
    def __init__(self, config=None):
        # 搬运原 default_config 中的表格相关配置
        default_params = {
            'table_width_percent': 95,  # 表格宽度占页面宽度的百分比
            'table_alignment': WD_TABLE_ALIGNMENT.CENTER,
            'column_widths': [],  # 自定义列宽百分比
            'auto_adjust_columns': True,
        }
        super().__init__({**default_params, **(config or {})})
    
    def apply(self, doc_context) -> RuleResult:
        """
        核心执行逻辑
        :param doc_context: 文档上下文对象
        """
        document = doc_context.get_document()
        fixed_count = 0
        details = []
        
        if not document.tables:
            details.append("文档中没有表格，跳过表格宽度优化")
            return RuleResult(
                rule_id=self.rule_id,
                success=True,
                fixed_count=fixed_count,
                details=details
            )
        
        details.append(f"开始优化表格（共{len(document.tables)}个）...")

        # 获取第一个分区的页面信息（所有分区应该相同）
        section = document.sections[0]
        available_width_cm = (section.page_width.cm -
                             section.left_margin.cm -
                             section.right_margin.cm)

        for table_idx, table in enumerate(document.tables):
            details.append(f"处理表格 {table_idx + 1}: {len(table.columns)}列")

            # 检查是否有合并单元格
            has_merged_cells = self._check_merged_cells(table)
            if has_merged_cells:
                details.append(f"注意: 表格包含合并单元格")

            # 检查是否有嵌套表格
            has_nested_tables = self._check_nested_tables(table)
            if has_nested_tables:
                details.append(f"注意: 表格包含嵌套表格")

            # 计算表格宽度
            table_width_cm = available_width_cm * self.config['table_width_percent'] / 100
            table.width = Cm(table_width_cm)
            table.alignment = self.config['table_alignment']
            
            details.append(f"表格宽度: {table_width_cm:.1f}cm ({self.config['table_width_percent']}%)")
            
            # 处理列宽
            col_count = len(table.columns)
            
            if self.config['column_widths'] and len(self.config['column_widths']) == col_count:
                # 使用自定义列宽
                total_percent = sum(self.config['column_widths'])
                for i, col in enumerate(table.columns):
                    col_width_cm = (table_width_cm * 
                                   self.config['column_widths'][i] / 
                                   total_percent)
                    col.width = Cm(col_width_cm)
                    details.append(f"列{i+1}宽度: {col_width_cm:.1f}cm ({self.config['column_widths'][i]}%)")
            elif self.config['auto_adjust_columns']:
                # 自动调整列宽
                self._auto_adjust_column_widths(table, table_width_cm, details)
            else:
                # 平均分配列宽
                col_width_cm = table_width_cm / col_count
                for col in table.columns:
                    col.width = Cm(col_width_cm)
                details.append(f"平均列宽: {col_width_cm:.1f}cm")
            
            # 处理嵌套表格
            if has_nested_tables:
                self._process_nested_tables(table, document, details)
            
            fixed_count += 1
        
        details.append(f"总共优化了 {fixed_count} 个表格的宽度")
        
        return RuleResult(
            rule_id=self.rule_id,
            success=True,
            fixed_count=fixed_count,
            details=details
        )
    
    def _check_merged_cells(self, table):
        """检查表格是否有合并单元格"""
        for row in table.rows:
            for cell in row.cells:
                # 检查是否有vMerge或hMerge属性（使用本地名称查询）
                tcPr = cell._tc.get_or_add_tcPr()
                for child in tcPr.iter():
                    local_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    if local_name in ['vMerge', 'hMerge']:
                        return True
        return False
    
    def _check_nested_tables(self, table):
        """检查表格是否有嵌套表格"""
        for row in table.rows:
            for cell in row.cells:
                if len(cell.tables) > 0:
                    return True
        return False
    
    def _process_nested_tables(self, table, document, details):
        """处理嵌套表格"""
        # 获取页面信息
        section = document.sections[0]
        available_width_cm = (section.page_width.cm -
                             section.left_margin.cm -
                             section.right_margin.cm)

        for row in table.rows:
            for cell in row.cells:
                for nested_table in cell.tables:
                    # 为嵌套表格设置较小的宽度
                    nested_table_width_cm = cell.width.cm * 0.9 if cell.width else available_width_cm * 0.7
                    nested_table.width = Cm(nested_table_width_cm)

                    # 为嵌套表格添加边框
                    for nested_row in nested_table.rows:
                        for nested_cell in nested_row.cells:
                            # 这里只是设置宽度，边框会由TableBorderRule处理
                            pass

                    details.append(f"处理了一个嵌套表格，宽度: {nested_table_width_cm:.1f}cm")
    
    def _auto_adjust_column_widths(self, table, table_width_cm, details):
        """自动调整列宽 - 支持复杂表格"""
        col_count = len(table.columns)
        max_lengths = [0] * col_count
        
        for row in table.rows:
            for j, cell in enumerate(row.cells):
                # 处理合并单元格的情况
                # 检查单元格是否跨列
                tcPr = cell._tc.get_or_add_tcPr()
                # 使用本地名称查询替代带有命名空间前缀的XPath查询
                gridSpan = None
                for child in tcPr.iter():
                    local_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    if local_name == 'gridSpan':
                        gridSpan = child
                        break
                span = 1
                if gridSpan is not None:
                    # 查找val属性
                    val = None
                    for attr in gridSpan.attrib:
                        local_attr_name = attr.split('}')[-1] if '}' in attr else attr
                        if local_attr_name == 'val':
                            val = gridSpan.attrib[attr]
                            break
                    if val:
                        span = int(val)
                
                text = cell.text.strip()
                if text:
                    # 估算文本长度（中文占2个字符宽度）
                    chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', text))
                    english_chars = len(text) - chinese_chars
                    length = (english_chars + chinese_chars * 2) / span  # 平均分配到跨列
                    
                    # 更新跨列的最大长度
                    for k in range(j, min(j + span, col_count)):
                        max_lengths[k] = max(max_lengths[k], length)
        
        total_length = sum(max_lengths) if sum(max_lengths) > 0 else 1
        
        for j, col in enumerate(table.columns):
            width_ratio = max_lengths[j] / total_length
            col_width_cm = table_width_cm * width_ratio
            col.width = Cm(col_width_cm)
            details.append(f"列{j+1}宽度: {col_width_cm:.1f}cm (比例: {width_ratio:.2f})")
