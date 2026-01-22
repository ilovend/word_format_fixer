import socket
import time
import os
from typing import Optional

class PortManager:
    """端口管理器 - 统一管理端口分配和发现"""

    # 默认端口配置
    DEFAULT_PORT = 7777
    MAX_ATTEMPTS = 10
    TIMEOUT = 0.5  # 秒

    # 端口文件路径
    # __file__ 是 utils/port_manager.py
    # os.path.dirname(__file__) = utils/
    # os.path.dirname(utils/) = python-backend/
    # os.path.dirname(python-backend/) = word_format_fixer/ (项目根目录)
    PORT_FILE_PATH = os.path.join(
        os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
        '.port'
    )

    def __init__(self, host: str = "127.0.0.1", start_port: int = None):
        self.host = host
        self.start_port = start_port or self.DEFAULT_PORT
        self.current_port = None

    def check_port_availability(self, port: int) -> bool:
        """
        检查端口是否可用
        :param port: 端口号
        :return: True表示可用，False表示被占用
        """
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.settimeout(self.TIMEOUT)
                result = s.connect_ex((self.host, port))
                return result != 0  # 0表示被占用，非0表示可用
        except Exception:
            return False

    def find_available_port(self) -> int:
        """
        寻找可用端口
        :return: 可用的端口号
        """
        for attempt in range(self.MAX_ATTEMPTS):
            port = self.start_port + attempt

            if self.check_port_availability(port):
                self.current_port = port
                self._write_port_to_file(port)
                return port

            # 短暂休眠避免检测过快
            time.sleep(0.1)

        # 如果都不可用，使用默认端口
        self.current_port = self.start_port
        self._write_port_to_file(self.start_port)
        return self.start_port

    def get_port(self) -> Optional[int]:
        """
        获取当前端口
        :return: 端口号，如果未设置则返回None
        """
        if self.current_port is None:
            self.current_port = self._read_port_from_file()
        return self.current_port

    def _write_port_to_file(self, port: int) -> None:
        """
        将端口号写入文件
        :param port: 端口号
        """
        try:
            with open(self.PORT_FILE_PATH, 'w') as f:
                f.write(str(port))
        except Exception as e:
            print(f"Warning: Failed to write port file: {e}")

    def _read_port_from_file(self) -> Optional[int]:
        """
        从文件读取端口号
        :return: 端口号，如果文件不存在或读取失败则返回None
        """
        try:
            if os.path.exists(self.PORT_FILE_PATH):
                with open(self.PORT_FILE_PATH, 'r') as f:
                    port_str = f.read().strip()
                    if port_str:
                        return int(port_str)
        except Exception as e:
            print(f"Warning: Failed to read port file: {e}")
        return None

    def cleanup_port_file(self) -> None:
        """清理端口文件"""
        try:
            if os.path.exists(self.PORT_FILE_PATH):
                os.remove(self.PORT_FILE_PATH)
        except Exception as e:
            print(f"Warning: Failed to cleanup port file: {e}")

# 全局端口管理器实例
port_manager = PortManager()
