"""Word文档格式修复工具"""

__version__ = "1.0.0"

from .core.fixer import RobustWordFixer
from .core.config import load_config, save_config, get_preset_config
from .ui.main import run_app
from .cli.main import main

__all__ = [
    'RobustWordFixer',
    'load_config',
    'save_config',
    'get_preset_config',
    'run_app',
    'main',
]
