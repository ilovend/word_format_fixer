from typing import Dict, List, Any
from rules.base_rule import RuleResult

class FixReporter:
    """修复报告生成器"""
    
    def __init__(self):
        self.results: Dict[str, RuleResult] = {}
    
    def add_result(self, rule_id: str, result: RuleResult):
        """添加规则执行结果"""
        self.results[rule_id] = result
    
    def generate_summary(self) -> Dict[str, Any]:
        """生成修复摘要"""
        total_fixed = 0
        success_count = 0
        failed_count = 0
        
        for result in self.results.values():
            if result.success:
                success_count += 1
                total_fixed += result.fixed_count
            else:
                failed_count += 1
        
        return {
            "total_rules": len(self.results),
            "success_count": success_count,
            "failed_count": failed_count,
            "total_fixed": total_fixed,
            "success_rate": f"{success_count / len(self.results) * 100:.1f}%" if self.results else "0%"
        }
    
    def generate_detailed_report(self) -> Dict[str, Any]:
        """生成详细报告"""
        return {
            "summary": self.generate_summary(),
            "detailed_results": {
                rule_id: result.dict()
                for rule_id, result in self.results.items()
            }
        }