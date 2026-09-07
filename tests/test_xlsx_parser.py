import os
import pytest
from pathlib import Path
import openpyxl
from backend.core.parser_xlsx import XlsxParserEngine

def test_xlsx_extract_and_rebuild(tmp_path):
    sample_path = Path("sample_docs/工艺参数表_模具点检.xlsx")
    assert sample_path.exists()

    # 1. 提取 Excel 文本内容
    items = XlsxParserEngine.extract_content(str(sample_path))
    assert len(items) > 0

    # 2. 模拟翻译结果
    items_map = {}
    for it in items:
        items_map[it["id"]] = f"EN_{it['source_text']}"

    # 3. 测试中文在上 (zh_top) 重构 Excel
    output_xlsx_top = tmp_path / "test_bilingual_top.xlsx"
    XlsxParserEngine.rebuild_bilingual_doc(
        original_file_path=str(sample_path),
        output_file_path=str(output_xlsx_top),
        items_map=items_map,
        layout_mode="zh_top",
        source_lang="zh",
        target_lang="en"
    )
    assert output_xlsx_top.exists()
    wb_top = openpyxl.load_workbook(str(output_xlsx_top))
    ws_top = wb_top.active
    first_cell_top = str(ws_top["A1"].value)
    assert "EN_" in first_cell_top

    # 4. 测试中文在下 (zh_bottom) 重构 Excel
    output_xlsx_bottom = tmp_path / "test_bilingual_bottom.xlsx"
    XlsxParserEngine.rebuild_bilingual_doc(
        original_file_path=str(sample_path),
        output_file_path=str(output_xlsx_bottom),
        items_map=items_map,
        layout_mode="zh_bottom",
        source_lang="zh",
        target_lang="en"
    )
    assert output_xlsx_bottom.exists()
    wb_bottom = openpyxl.load_workbook(str(output_xlsx_bottom))
    ws_bottom = wb_bottom.active
    first_cell_bottom = str(ws_bottom["A1"].value)
    assert first_cell_bottom.startswith("EN_")
