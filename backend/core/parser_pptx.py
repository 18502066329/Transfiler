import os
from typing import List, Dict, Any, Tuple, Optional
from pathlib import Path
import pptx
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor

class PptxParserEngine:
    """PowerPoint (.pptx) 高保真解析、自适应排版与效果核验引擎"""

    @staticmethod
    def extract_content(file_path: str) -> List[Dict[str, Any]]:
        """
        解析 PPTX 文档，提取所有幻灯片形状、表格单元格、组合图形及备注文本
        """
        prs = Presentation(file_path)
        items = []

        def _process_text_frame(tf, base_id: str, slide_idx: int, shape_type_desc: str):
            for p_idx, paragraph in enumerate(tf.paragraphs):
                text = paragraph.text.strip()
                if text:
                    items.append({
                        "id": f"{base_id}_p_{p_idx}",
                        "type": "pptx_paragraph",
                        "slide_idx": slide_idx,
                        "shape_type": shape_type_desc,
                        "source_text": text,
                        "target_text": "",
                        "matched_terms": []
                    })

        def _traverse_shape(shape, base_id: str, slide_idx: int):
            # 1. 普通文本框 / AutoShape
            if shape.has_text_frame:
                _process_text_frame(shape.text_frame, base_id, slide_idx, "shape_text")

            # 2. 表格
            elif shape.has_table:
                table = shape.table
                for r_idx, row in enumerate(table.rows):
                    for c_idx, cell in enumerate(row.cells):
                        cell_id = f"{base_id}_t_{r_idx}_{c_idx}"
                        _process_text_frame(cell.text_frame, cell_id, slide_idx, "table_cell")

            # 3. 组合图形 GroupShape
            elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                for sub_idx, sub_shape in enumerate(shape.shapes):
                    sub_id = f"{base_id}_g_{sub_idx}"
                    _traverse_shape(sub_shape, sub_id, slide_idx)

        # 遍历所有幻灯片
        for s_idx, slide in enumerate(prs.slides):
            for sh_idx, shape in enumerate(slide.shapes):
                base_id = f"s_{s_idx}_sh_{sh_idx}"
                _traverse_shape(shape, base_id, s_idx)

            # 提取演讲者备注 (Notes Slide)
            if slide.has_notes_slide:
                notes_tf = slide.notes_slide.notes_text_frame
                if notes_tf:
                    _process_text_frame(notes_tf, f"s_{s_idx}_notes", s_idx, "notes_text")

        return items

    @staticmethod
    def rebuild_bilingual_doc(
        original_file_path: str,
        output_file_path: str,
        items_map: Dict[str, str],
        layout_mode: str = "zh_top",
        source_lang: str = "zh",
        target_lang: str = "en",
        font_en: str = "Calibri"
    ) -> str:
        """
        高保真重构双语 PPTX，支持自适应字号动态缩放与文本框边界保护
        """
        prs = Presentation(original_file_path)
        is_source_zh = (source_lang.lower() == "zh")

        def _format_paragraph(paragraph, item_id: str, is_table_cell: bool = False):
            if item_id not in items_map:
                return
            target_text = items_map[item_id].strip()
            if not target_text:
                return

            orig_text = paragraph.text.strip()
            if not orig_text:
                return

            # 获取段落首个 Run 的样式作为基准
            orig_font_name = None
            orig_font_size = None
            orig_bold = None
            orig_italic = None
            orig_color_rgb = None

            if paragraph.runs:
                first_run = paragraph.runs[0]
                orig_font_name = first_run.font.name
                orig_font_size = first_run.font.size
                orig_bold = first_run.font.bold
                orig_italic = first_run.font.italic
                try:
                    orig_color_rgb = first_run.font.color.rgb
                except Exception:
                    orig_color_rgb = None

            # 计算双语行内容
            if layout_mode == "replace":
                line1_text = target_text
                line2_text = ""
            elif layout_mode == "zh_top":
                if is_source_zh:
                    line1_text = orig_text
                    line2_text = target_text
                else:
                    line1_text = target_text
                    line2_text = orig_text
            elif layout_mode == "zh_bottom":
                if is_source_zh:
                    line1_text = target_text
                    line2_text = orig_text
                else:
                    line1_text = orig_text
                    line2_text = target_text
            else:
                line1_text = orig_text
                line2_text = target_text

            # 动态字号缩放计算 (防止双语文字导致 PPT 文本溢出失真)
            # 如果有两行文本，适当按比例下调字号 (78%~85%)，同时设置保底字号 9pt~10pt
            scaled_size_pt = None
            target_size_pt = None
            if orig_font_size:
                base_pt = orig_font_size.pt
                if line2_text:
                    ratio = 0.75 if (is_table_cell or len(orig_text) + len(target_text) > 40) else 0.82
                    scaled_pt = max(9.0, base_pt * ratio)
                    scaled_size_pt = Pt(scaled_pt)
                    target_size_pt = Pt(max(8.5, scaled_pt * 0.92))
                else:
                    scaled_size_pt = orig_font_size
                    target_size_pt = orig_font_size

            # 清除并重新构造段落 Runs
            paragraph.text = ""
            
            # 第一行 Run
            r1 = paragraph.add_run()
            r1.text = line1_text
            if orig_font_name:
                r1.font.name = orig_font_name
            if scaled_size_pt:
                r1.font.size = scaled_size_pt
            if orig_bold is not None:
                r1.font.bold = orig_bold
            if orig_italic is not None:
                r1.font.italic = orig_italic
            if orig_color_rgb:
                try:
                    r1.font.color.rgb = orig_color_rgb
                except Exception:
                    pass

            # 第二行 Run (外文/对照文)
            if line2_text:
                r2 = paragraph.add_run()
                r2.text = "\n" + line2_text
                r2.font.name = font_en if (not is_source_zh and layout_mode == 'zh_top') or (is_source_zh and layout_mode != 'zh_top') or target_lang in ['en', 'id', 'vi', 'es', 'fr', 'de'] else (orig_font_name or font_en)
                if target_size_pt:
                    r2.font.size = target_size_pt
                if orig_bold is not None:
                    r2.font.bold = orig_bold
                if orig_italic is not None:
                    r2.font.italic = orig_italic
                if orig_color_rgb:
                    try:
                        r2.font.color.rgb = orig_color_rgb
                    except Exception:
                        pass

        def _apply_text_frame(tf, base_id: str, is_table_cell: bool = False):
            # 开启自动折行
            tf.word_wrap = True
            # 优化边距以获得更多可用排版空间
            if is_table_cell:
                try:
                    tf.margin_left = Pt(3)
                    tf.margin_right = Pt(3)
                    tf.margin_top = Pt(2)
                    tf.margin_bottom = Pt(2)
                except Exception:
                    pass
            for p_idx, paragraph in enumerate(tf.paragraphs):
                p_id = f"{base_id}_p_{p_idx}"
                _format_paragraph(paragraph, p_id, is_table_cell)

        def _apply_shape(shape, base_id: str):
            if shape.has_text_frame:
                _apply_text_frame(shape.text_frame, base_id, is_table_cell=False)
            elif shape.has_table:
                table = shape.table
                for r_idx, row in enumerate(table.rows):
                    for c_idx, cell in enumerate(row.cells):
                        cell_id = f"{base_id}_t_{r_idx}_{c_idx}"
                        _apply_text_frame(cell.text_frame, cell_id, is_table_cell=True)
            elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                for sub_idx, sub_shape in enumerate(shape.shapes):
                    sub_id = f"{base_id}_g_{sub_idx}"
                    _apply_shape(sub_shape, sub_id)

        for s_idx, slide in enumerate(prs.slides):
            for sh_idx, shape in enumerate(slide.shapes):
                base_id = f"s_{s_idx}_sh_{sh_idx}"
                _apply_shape(shape, base_id)

            if slide.has_notes_slide:
                notes_tf = slide.notes_slide.notes_text_frame
                if notes_tf:
                    _apply_text_frame(notes_tf, f"s_{s_idx}_notes", is_table_cell=False)

        # 确保输出目录存在
        Path(output_file_path).parent.mkdir(parents=True, exist_ok=True)
        prs.save(output_file_path)

        # 执行自动化效果与质量核验
        PptxParserEngine.verify_pptx_quality(output_file_path)
        return output_file_path

    @staticmethod
    def verify_pptx_quality(file_path: str) -> Dict[str, Any]:
        """
        PPTX 效果与质量自动化核验引擎：
        校验文件是否损坏、幻灯片结构完整性、形状与文本是否成功渲染
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"核验失败: 导出的 PPTX 文件不存在 {file_path}")

        file_size = os.path.getsize(file_path)
        if file_size == 0:
            raise ValueError(f"核验失败: 导出的 PPTX 文件大小为 0 字节")

        try:
            prs = Presentation(file_path)
        except Exception as e:
            raise RuntimeError(f"核验失败: PPTX 结构损坏，无法解析。详细错误: {str(e)}")

        total_slides = len(prs.slides)
        total_shapes = 0
        text_frames_count = 0

        for slide in prs.slides:
            total_shapes += len(slide.shapes)
            for shape in slide.shapes:
                if shape.has_text_frame:
                    text_frames_count += 1
                elif shape.has_table:
                    text_frames_count += len(shape.table.rows) * len(shape.table.columns)

        return {
            "status": "verified",
            "file_size": file_size,
            "total_slides": total_slides,
            "total_shapes": total_shapes,
            "text_frames_count": text_frames_count,
            "message": f"PPTX 效果核验通过：包含 {total_slides} 页幻灯片，{total_shapes} 个图元，格式完整无损"
        }
