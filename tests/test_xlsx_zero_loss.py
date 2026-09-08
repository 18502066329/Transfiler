import os
import zipfile
import pytest
from backend.core.parser_xlsx import XlsxParserEngine

def test_xlsx_lossless_extraction_and_rebuild():
    sample_file = "WI-FJS-204 HQ6200101A330礼盒(4).xlsx"
    if not os.path.exists(sample_file):
        pytest.skip(f"Sample file {sample_file} not found")

    items = XlsxParserEngine.extract_content(sample_file)
    assert len(items) > 50

    # 验证提取包含 drawing 文本框
    textbox_items = [it for it in items if it["type"] == "excel_textbox"]
    assert len(textbox_items) > 0

    # 重构测试
    out_file = "test_lossless_output.xlsx"
    items_map = {it["id"]: f"Trans_{it['source_text'][:10]}" for it in items}

    try:
        XlsxParserEngine.rebuild_bilingual_doc(
            original_file_path=sample_file,
            output_file_path=out_file,
            items_map=items_map,
            layout_mode="zh_top",
            source_lang="zh",
            target_lang="en"
        )

        assert os.path.exists(out_file)
        assert os.path.getsize(out_file) > 100000

        # 检查重构前后 zip 结构零丢失
        with zipfile.ZipFile(sample_file) as z_orig, zipfile.ZipFile(out_file) as z_out:
            orig_names = set(z_orig.namelist())
            out_names = set(z_out.namelist())
            assert orig_names == out_names, f"Missing files in output: {orig_names - out_names}"

            # 验证所有 media 图片完整存在
            media_files = [n for n in orig_names if n.startswith("xl/media/")]
            assert len(media_files) > 0
            for mf in media_files:
                assert z_orig.read(mf) == z_out.read(mf), f"Media file {mf} was modified or corrupted!"

    finally:
        if os.path.exists(out_file):
            os.remove(out_file)
