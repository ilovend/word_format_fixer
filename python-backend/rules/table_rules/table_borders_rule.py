from rules.base_rule import BaseRule, RuleResult
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

class TableBordersRule(BaseRule):
    """表格边框统一规则"""

    display_name = "表格边框统一"
    category = "表格规则"
    
    def __init__(self, config=None):
        default_params = {
            'border_size': 4,  # 边框大小，默认4磅
            'border_color': '000000',  # 边框颜色，默认黑色
            'vertical_alignment': 'center',  # 垂直对齐方式
        }
        super().__init__({**default_params, **(config or {})})
    
    def apply(self, context):
        """应用表格边框规则"""
        document = context.get_document()
        fixed_count = 0
        details = []
        
        # 对齐方式映射
        align_map = {
            'center': WD_ALIGN_VERTICAL.CENTER,
            'top': WD_ALIGN_VERTICAL.TOP,
            'bottom': WD_ALIGN_VERTICAL.BOTTOM,
        }
        vertical_alignment = align_map.get(self.config['vertical_alignment'], WD_ALIGN_VERTICAL.CENTER)
        
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    # 设置单元格垂直居中
                    cell.vertical_alignment = vertical_alignment
                    # 设置单元格边框
                    self._set_cell_border(cell)
                    fixed_count += 1
        
        details.append(f"统一了{fixed_count}个表格单元格的边框格式")
        details.append(f"边框大小: {self.config['border_size']}磅")
        details.append(f"边框颜色: {self.config['border_color']}")
        details.append(f"垂直对齐: {self.config['vertical_alignment']}")
        return RuleResult(
            rule_id=self.rule_id,
            success=True,
            fixed_count=fixed_count,
            details=details
        )
    
    def _set_cell_border(self, cell):
        """设置单元格边框"""
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        
        borders = ['top', 'left', 'bottom', 'right']
        for border in borders:
            border_elem = OxmlElement(f'w:{border}')
            border_elem.set(qn('w:val'), 'single')
            border_elem.set(qn('w:sz'), str(self.config['border_size']))
            border_elem.set(qn('w:space'), '0')
            border_elem.set(qn('w:color'), self.config['border_color'])
            tcPr.append(border_elem)
    
    def explain(self) -> str:
        """解释规则"""
        return "为文档中所有表格的单元格添加统一的边框，并设置垂直居中对齐"
