from typing import List, Dict, Any
import importlib
import os
import time
from rules.base_rule import BaseRule, RuleResult
from core.context import RuleContext

class RuleEngine:
    """规则执行引擎"""
    
    def __init__(self):
        self.rules: Dict[str, BaseRule] = {}
        self._load_rules()
    
    def _load_rules(self):
        """动态加载所有规则"""
        # 扫描rules目录下的所有规则
        rules_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'rules')
        
        # 遍历所有规则目录
        for root, dirs, files in os.walk(rules_dir):
            for file in files:
                if file.endswith('.py') and not file.startswith('__init__') and not file.startswith('base_rule'):
                    # 计算模块路径
                    relative_path = os.path.relpath(os.path.join(root, file), rules_dir)
                    module_path = 'rules.' + relative_path.replace(os.path.sep, '.').replace('.py', '')
                    
                    try:
                        # 导入模块
                        module = importlib.import_module(module_path)
                        # 查找模块中的规则类
                        for name, obj in module.__dict__.items():
                            if isinstance(obj, type) and issubclass(obj, BaseRule) and obj != BaseRule:
                                # 创建规则实例
                                rule_instance = obj()
                                self.rules[rule_instance.rule_id] = rule_instance
                    except Exception as e:
                        print(f"加载规则失败 {module_path}: {e}")
    
    def register_rule(self, rule: BaseRule):
        """注册规则"""
        self.rules[rule.rule_id] = rule
    
    def execute(self, document_path: str, active_rules: List[Dict[str, Any]] = None) -> Dict[str, Any]:
        """
        执行规则
        :param document_path: 文档路径
        :param active_rules: 前端传来的激活规则列表
        """
        start_time = time.time()
        context = RuleContext(document_path)
        results = []
        total_fixed = 0
        
        # 确定要执行的规则
        if active_rules is not None:
            # 执行前端指定的规则
            for rule_info in active_rules:
                rule_id = rule_info['rule_id']
                params = rule_info.get('params', {})
                
                if rule_id in self.rules:
                    rule = self.rules[rule_id]
                    try:
                        # 更新规则配置
                        rule.config.update(params)
                        # 执行规则
                        result = rule.apply(context)
                        # 直接构建结果字典，避免使用dict()方法
                        result_dict = {
                            "rule_id": result.rule_id,
                            "success": result.success,
                            "fixed_count": result.fixed_count,
                            "details": result.details
                        }
                        results.append(result_dict)
                        total_fixed += result.fixed_count
                    except Exception as e:
                        # 构建失败结果
                        results.append({
                            "rule_id": rule_id,
                            "success": False,
                            "fixed_count": 0,
                            "details": [f"执行失败: {str(e)}"]
                        })
                else:
                    # 无效规则ID，添加失败结果
                    results.append({
                        "rule_id": rule_id,
                        "success": False,
                        "fixed_count": 0,
                        "details": [f"规则ID不存在: {rule_id}"]
                    })
        else:
            # 执行所有启用的规则
            for rule_id, rule in self.rules.items():
                if rule.enabled:
                    try:
                        result = rule.apply(context)
                        # 直接构建结果字典，避免使用dict()方法
                        result_dict = {
                            "rule_id": result.rule_id,
                            "success": result.success,
                            "fixed_count": result.fixed_count,
                            "details": result.details
                        }
                        results.append(result_dict)
                        total_fixed += result.fixed_count
                    except Exception as e:
                        results.append({
                            "rule_id": rule_id,
                            "success": False,
                            "fixed_count": 0,
                            "details": [f"执行失败: {str(e)}"]
                        })
        
        # 保存修改后的文档
        save_success = context.save_document()
        time_taken = f"{time.time() - start_time:.2f}s"
        
        return {
            "status": "success" if save_success else "error",
            "summary": {
                "total_fixed": total_fixed,
                "time_taken": time_taken
            },
            "results": results,
            "save_success": save_success,
            "saved_to": document_path
        }
    
    def get_rules_info(self) -> List[Dict[str, Any]]:
        """获取所有规则信息"""
        return [rule.get_metadata() for rule in self.rules.values()]
