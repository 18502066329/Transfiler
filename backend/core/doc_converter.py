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

    @staticmethod
    def convert_ppt_to_pptx(ppt_path: str) -> str:
        """
        将 .ppt 转换为 .pptx 格式并返回新文件路径
        """
        ppt_path_obj = Path(ppt_path).resolve()
        if not ppt_path_obj.exists():
            raise FileNotFoundError(f"未找到文件: {ppt_path}")

        if ppt_path_obj.suffix.lower() == ".pptx":
            return str(ppt_path_obj)

        pptx_path_obj = ppt_path_obj.with_suffix(".pptx")

        # 尝试使用 Windows 本地 PowerPoint/WPS COM 组件转换
        try:
            import win32com.client as win32
            powerpoint = win32.DispatchEx("PowerPoint.Application")
            # 24 = ppSaveAsOpenXMLPresentation (.pptx 格式代码)
            pres = powerpoint.Presentations.Open(str(ppt_path_obj), WithWindow=False)
            pres.SaveAs(str(pptx_path_obj), 24)
            pres.Close()
            powerpoint.Quit()
            return str(pptx_path_obj)
        except Exception as e:
            raise RuntimeError(
                f"转换老版 .ppt 文件失败。建议直接在 Office/WPS 中将其另存为 .pptx 格式再导入。\n详细错误: {str(e)}"
            )
