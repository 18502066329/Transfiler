import os
import sys
from pathlib import Path

class DocConverter:
    """老版 .doc 转 .docx 兼容转换器"""

    @staticmethod
    def convert_doc_to_docx(doc_path: str) -> str:
        """
        将 .doc 转换为 .docx 格式并返回新文件路径
        """
        doc_path_obj = Path(doc_path).resolve()
        if not doc_path_obj.exists():
            raise FileNotFoundError(f"未找到文件: {doc_path}")

        if doc_path_obj.suffix.lower() == ".docx":
            return str(doc_path_obj)

        docx_path_obj = doc_path_obj.with_suffix(".docx")

        # 尝试使用 Windows 本地 Word/WPS COM 组件转换
        try:
            import win32com.client as win32
            word = win32.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            
            doc = word.Documents.Open(str(doc_path_obj))
            # 16 = wdFormatXMLDocument (.docx 格式代码)
            doc.SaveAs2(str(docx_path_obj), FileFormat=16)
            doc.Close()
            word.Quit()
            return str(docx_path_obj)
        except Exception as e:
            # 如果没有安装 Word COM 或运行在非 COM 环境
            raise RuntimeError(
                f"转换老版 .doc 文件失败。建议直接在 Office 中将其另存为 .docx 格式再导入。\n详细错误: {str(e)}"
            )
