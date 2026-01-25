"""
Visual Diff Service - 文档对比服务
提供文档修改前后的可视化对比功能
"""

import mammoth
import difflib
import tempfile
import shutil
import os
from typing import Dict, Any, List, Optional


class DiffService:
    """文档对比服务 - 提供修改前后的可视化对比"""
    
    def __init__(self):
        self._temp_dir = None
        self._original_html = None
        self._original_path = None
    
    def prepare_diff(self, document_path: str) -> Dict[str, Any]:
        """
        准备对比：在处理文档前调用，缓存原始文档的HTML
        
        Args:
            document_path: 原始文档路径
            
        Returns:
            包含状态和原始HTML的字典
        """
        try:
            # 创建临时目录存储原始文档副本
            self._temp_dir = tempfile.mkdtemp(prefix="word_diff_")
            self._original_path = document_path
            
            # 复制原始文档到临时目录
            temp_doc_path = os.path.join(self._temp_dir, "original.docx")
            shutil.copy2(document_path, temp_doc_path)
            
            # 转换原始文档为HTML
            self._original_html = self._docx_to_html(temp_doc_path)
            
            return {
                "status": "success",
                "message": "Original document cached for comparison",
                "original_html": self._original_html
            }
        except Exception as e:
            return {
                "status": "error",
                "message": f"Failed to prepare diff: {str(e)}"
            }
    
    def generate_diff(self, modified_path: str) -> Dict[str, Any]:
        """
        生成对比：在处理文档后调用，生成差异数据
        
        Args:
            modified_path: 修改后的文档路径
            
        Returns:
            包含差异数据的字典
        """
        try:
            if not self._original_html:
                return {
                    "status": "error",
                    "message": "No original document cached. Call prepare_diff first."
                }
            
            # 转换修改后的文档为HTML
            modified_html = self._docx_to_html(modified_path)
            
            # 生成差异
            diff_result = self._generate_html_diff(self._original_html, modified_html)
            
            return {
                "status": "success",
                "original_html": self._original_html,
                "modified_html": modified_html,
                "diff_html": diff_result["diff_html"],
                "changes": diff_result["changes"],
                "stats": diff_result["stats"]
            }
        except Exception as e:
            return {
                "status": "error",
                "message": f"Failed to generate diff: {str(e)}"
            }
        finally:
            self._cleanup()
    
    def get_document_preview(self, document_path: str) -> Dict[str, Any]:
        """
        获取文档的HTML预览
        
        Args:
            document_path: 文档路径
            
        Returns:
            包含HTML预览的字典
        """
        try:
            html = self._docx_to_html(document_path)
            return {
                "status": "success",
                "html": html
            }
        except Exception as e:
            return {
                "status": "error",
                "message": f"Failed to generate preview: {str(e)}"
            }
    
    def _docx_to_html(self, docx_path: str) -> str:
        """
        将Docx文档转换为HTML
        
        Args:
            docx_path: Docx文件路径
            
        Returns:
            HTML字符串
        """
        with open(docx_path, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file)
            return result.value
    
    def _generate_html_diff(self, original: str, modified: str) -> Dict[str, Any]:
        """
        生成HTML差异对比
        
        Args:
            original: 原始HTML
            modified: 修改后的HTML
            
        Returns:
            包含差异HTML和统计信息的字典
        """
        # 将HTML按行分割进行对比
        original_lines = original.splitlines(keepends=True)
        modified_lines = modified.splitlines(keepends=True)
        
        # 使用difflib生成差异
        differ = difflib.HtmlDiff(wrapcolumn=80)
        diff_table = differ.make_table(
            original_lines, 
            modified_lines,
            fromdesc="修改前",
            todesc="修改后",
            context=True,
            numlines=3
        )
        
        # 生成unified diff用于统计
        unified_diff = list(difflib.unified_diff(
            original_lines, 
            modified_lines,
            lineterm=''
        ))
        
        # 统计变更
        additions = sum(1 for line in unified_diff if line.startswith('+') and not line.startswith('+++'))
        deletions = sum(1 for line in unified_diff if line.startswith('-') and not line.startswith('---'))
        
        # 提取具体变更列表
        changes = self._extract_changes(unified_diff)
        
        return {
            "diff_html": diff_table,
            "changes": changes,
            "stats": {
                "additions": additions,
                "deletions": deletions,
                "total_changes": additions + deletions
            }
        }
    
    def _extract_changes(self, unified_diff: List[str]) -> List[Dict[str, Any]]:
        """
        从unified diff中提取变更列表
        
        Args:
            unified_diff: unified diff行列表
            
        Returns:
            变更列表
        """
        changes = []
        current_change = None
        
        for line in unified_diff:
            if line.startswith('@@'):
                if current_change:
                    changes.append(current_change)
                current_change = {
                    "header": line,
                    "removed": [],
                    "added": []
                }
            elif current_change:
                if line.startswith('-') and not line.startswith('---'):
                    current_change["removed"].append(line[1:])
                elif line.startswith('+') and not line.startswith('+++'):
                    current_change["added"].append(line[1:])
        
        if current_change:
            changes.append(current_change)
        
        return changes
    
    def _cleanup(self):
        """清理临时文件"""
        if self._temp_dir and os.path.exists(self._temp_dir):
            try:
                shutil.rmtree(self._temp_dir)
            except Exception:
                pass
        self._temp_dir = None
        self._original_html = None
        self._original_path = None
