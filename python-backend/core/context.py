from docx import Document
from typing import Optional, Dict, Any

class RuleContext:
    """规则执行上下文"""
    
    def __init__(self, document_path: str):
        self.document_path = document_path
        self.document: Optional[Document] = None
        self._load_document()
        self._cache: Dict[str, Any] = {}
        self.available_width_cm = 15.92  # 默认值，页面布局规则会更新它
        self.runtime_data = {}  # 用于规则间传递临时数据
    
    def _load_document(self):
        """加载文档"""
        try:
            self.document = Document(self.document_path)
        except Exception as e:
            raise Exception(f"加载文档失败: {str(e)}")
    
    def get_document(self) -> Document:
        """获取文档对象"""
        if not self.document:
            self._load_document()
        return self.document
    
    def set_cache(self, key: str, value: Any):
        """设置缓存"""
        self._cache[key] = value
    
    def get_cache(self, key: str, default: Any = None) -> Any:
        """获取缓存"""
        return self._cache.get(key, default)
    
    def clear_cache(self):
        """清除缓存"""
        self._cache.clear()
    
    def save_document(self, output_path: Optional[str] = None):
        """保存文档"""
        save_path = output_path or self.document_path
        if self.document:
            self.document.save(save_path)
            return True
        return False