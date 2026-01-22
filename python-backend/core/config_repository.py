from abc import ABC, abstractmethod
from typing import Dict, Any, Optional

class IConfigRepository(ABC):
    """配置仓库接口 - 抽象持久化操作"""

    @abstractmethod
    def load_config(self) -> Dict[str, Any]:
        """加载配置"""
        pass

    @abstractmethod
    def save_config(self, config: Dict[str, Any]) -> None:
        """保存配置"""
        pass

    @abstractmethod
    def get_preset(self, preset_name: str) -> Optional[Dict[str, Any]]:
        """获取指定预设"""
        pass

    @abstractmethod
    def delete_preset(self, preset_id: str) -> None:
        """删除预设"""
        pass

    @abstractmethod
    def save_preset(self, preset_id: str, preset_data: Dict[str, Any]) -> None:
        """保存预设"""
        pass
