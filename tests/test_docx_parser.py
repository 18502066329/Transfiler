import os
import pytest
from pathlib import Path
from backend.core.parser_docx import DocxParserEngine
import docx

def test_docx_extract_and_rebuild(tmp_path):
    sample_path = Path("sample_docs/SOP_注塑机标准作业指导书.docx")
    assert sample_path.exists()

    # 1. 提取内容
    items = DocxParserEngine.extract_content(str(sample_path))
    assert len(items) > 0

    # 验证提取到的条目中包含标题和表格
    has_paragraph = any(it["type"] == "paragraph" for it in items)
    has_table = any(it["type"] == "table_cell" for it in items)
    assert has_paragraph
    assert has_table

    # 2. 模拟翻译结果
    items_map = {}
    for it in items:
        items_map[it["id"]] = f"Translated: {it['source_text']}"

    # 3. 测试中文在上 (zh_top) 重构 Word
    output_docx_top = tmp_path / "test_bilingual_sop_top.docx"
    DocxParserEngine.rebuild_bilingual_doc(
        original_file_path=str(sample_path),
        output_file_path=str(output_docx_top),
        items_map=items_map,
        layout_mode="zh_top",
        source_lang="zh",
        target_lang="en"
    )
    assert output_docx_top.exists()
    doc_top = docx.Document(str(output_docx_top))
    all_text_top = "\n".join([p.text for p in doc_top.paragraphs])
    assert "注塑车间标准作业指导书" in all_text_top
    assert "Translated:" in all_text_top

    # 4. 测试中文在下 (zh_bottom) 重构 Word
    output_docx_bottom = tmp_path / "test_bilingual_sop_bottom.docx"
    DocxParserEngine.rebuild_bilingual_doc(
        original_file_path=str(sample_path),
        output_file_path=str(output_docx_bottom),
        items_map=items_map,
        layout_mode="zh_bottom",
        source_lang="zh",
        target_lang="en"
    )
    assert output_docx_bottom.exists()
    doc_bottom = docx.Document(str(output_docx_bottom))
    first_p = doc_bottom.paragraphs[0].text
    # 中文在下：译文在第一行，原中文在第二行
    assert "Translated:" in first_p
