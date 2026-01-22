import http.server
import socketserver
import json
from typing import Dict, Any

from services import (
    DocumentProcessingService,
    ConfigManagementService,
    RuleManagementService
)
from utils import port_manager

# 全局服务实例
doc_service = DocumentProcessingService()
config_service = ConfigManagementService()
rule_service = RuleManagementService()

class IPCRequestHandler(http.server.BaseHTTPRequestHandler):
    """IPC请求处理器 - 只负责HTTP协议转换"""

    def _send_response(self, status_code: int, data: Dict[str, Any]):
        """发送JSON响应"""
        self.send_response(status_code)
        self.send_header('Content-type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps(data).encode('utf-8'))

    def do_POST(self):
        """处理POST请求"""
        if self.path == '/api/process':
            self._handle_process_request()
        elif self.path == '/api/presets/save':
            self._handle_save_preset_request()
        else:
            self._send_response(404, {"error": "Not found"})

    def do_DELETE(self):
        """处理DELETE请求"""
        if self.path == '/api/presets/delete':
            self._handle_delete_preset_request()
        else:
            self._send_response(404, {"error": "Not found"})

    def do_GET(self):
        """处理GET请求"""
        if self.path == '/api/rules':
            self._handle_get_rules_request()
        elif self.path == '/api/presets':
            self._handle_get_presets_request()
        elif self.path == '/api/health':
            self._send_response(200, {"status": "healthy"})
        else:
            self._send_response(404, {"error": "Not found"})

    def _handle_process_request(self):
        """处理文档处理请求"""
        try:
            # 读取请求体
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            request_data = json.loads(post_data.decode('utf-8'))

            # 提取参数
            file_path = request_data.get('file_path')
            active_rules = request_data.get('active_rules', [])

            # 调用服务层处理
            result = doc_service.process_document(file_path, active_rules)

            # 构建响应
            response = {
                "status": result.get("status", "success"),
                "summary": result.get("summary", {}),
                "results": result.get("results", []),
                "save_success": result.get("save_success", False),
                "saved_to": result.get("saved_to", file_path)
            }
            self._send_response(200, response)
        except ValueError as e:
            self._send_response(400, {"error": str(e)})
        except Exception as e:
            self._send_response(500, {"error": str(e)})

    def _handle_save_preset_request(self):
        """处理保存预设请求"""
        try:
            # 读取请求体
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            request_data = json.loads(post_data.decode('utf-8'))

            # 提取参数
            preset_id = request_data.get('preset_id')
            preset_data = request_data.get('preset_data')

            # 调用服务层处理
            result = config_service.save_preset(preset_id, preset_data)
            self._send_response(200, result)
        except ValueError as e:
            self._send_response(400, {"error": str(e)})
        except Exception as e:
            self._send_response(500, {"error": str(e)})

    def _handle_delete_preset_request(self):
        """处理删除预设请求"""
        try:
            # 读取请求体
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            request_data = json.loads(post_data.decode('utf-8'))

            # 提取参数
            preset_id = request_data.get('preset_id')

            # 调用服务层处理
            result = config_service.delete_preset(preset_id)
            self._send_response(200, result)
        except ValueError as e:
            self._send_response(400, {"error": str(e)})
        except Exception as e:
            self._send_response(500, {"error": str(e)})

    def _handle_get_rules_request(self):
        """处理获取规则请求"""
        try:
            rules_info = rule_service.get_all_rules()
            self._send_response(200, rules_info)
        except Exception as e:
            self._send_response(500, {"error": str(e)})

    def _handle_get_presets_request(self):
        """处理获取预设请求"""
        try:
            presets = config_service.get_all_presets()
            self._send_response(200, presets)
        except Exception as e:
            self._send_response(500, {"error": str(e)})

def start_server(host: str = "127.0.0.1", port: int = 7777):
    """
    启动服务器
    :param host: 主机地址
    :param port: 端口号（将被端口管理器覆盖）
    """
    # 使用端口管理器获取可用端口
    available_port = port_manager.find_available_port()

    try:
        with socketserver.TCPServer((host, available_port), IPCRequestHandler) as httpd:
            print(f"Server running at http://{host}:{available_port}")
            httpd.serve_forever()
    except Exception as e:
        print(f"Error starting server: {e}")
        raise

if __name__ == "__main__":
    start_server()
