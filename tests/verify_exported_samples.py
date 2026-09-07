import os
from pathlib import Path
from backend.core.parser_docx import DocxParserEngine
from backend.core.parser_xlsx import XlsxParserEngine
from backend.core.glossary_matcher import GlossaryMatcher
from backend.db.database import init_db, get_enabled_glossary_dict

def run_verification():
    init_db()
    glossary = get_enabled_glossary_dict("en")
    matcher = GlossaryMatcher(glossary)

    # 1. 验证 Word 真实文件
    docx_in = "sample_docs/SOP_注塑机标准作业指导书.docx"
    docx_out = "exported_docs/SOP_注塑机标准作业指导书_双语对照版_EN.docx"
    
    docx_items = DocxParserEngine.extract_content(docx_in)
    print(f"Word 提取出 {len(docx_items)} 个段落与表格单元格")

    docx_map = {}
    for it in docx_items:
        src = it["source_text"]
        matched = matcher.match_terms_in_text(src)
        if "标准作业指导书" in src:
            docx_map[it["id"]] = "Standard Operating Procedure (SOP) for Injection Molding"
        elif "安全确认" in src:
            docx_map[it["id"]] = "1. Pre-operation Safety Confirmation & Preparation"
        elif "耐高温手套" in src:
            docx_map[it["id"]] = "1. Operators must wear heat-resistant gloves, safety goggles, and anti-static safety shoes before entering the workshop."
        elif "急停开关" in src:
            docx_map[it["id"]] = "2. Check the machine emergency stop switch and safety light curtain to confirm normal interlock protection."
        elif "防错治具" in src:
            docx_map[it["id"]] = "3. Confirm that the poka-yoke fixtures and locating jigs are firmly installed without looseness or deviation."
        elif "工艺参数项目" in src:
            docx_map[it["id"]] = "Process Parameter Item"
        elif "设定标准值" in src:
            docx_map[it["id"]] = "Standard Set Value"
        elif "控制公差" in src:
            docx_map[it["id"]] = "Tolerance Requirement"
        elif "射胶压力" in src:
            docx_map[it["id"]] = "Injection Pressure"
        elif "保压时间" in src:
            docx_map[it["id"]] = "Holding Time"
        elif "模具预热温度" in src:
            docx_map[it["id"]] = "Mold Preheating Temperature"
        elif "周期节拍" in src:
            docx_map[it["id"]] = "Cycle Time (CT)"
        elif "品质管控" in src:
            docx_map[it["id"]] = "3. Quality Control & First Article Inspection (FAI)"
        elif "首件全尺寸检验" in src:
            docx_map[it["id"]] = "First article full dimension inspection must be performed for every batch; check for flash, burr, sink mark, or short shot defects."
        else:
            docx_map[it["id"]] = f"English Translation: {src}"

    DocxParserEngine.rebuild_bilingual_doc(
        original_file_path=docx_in,
        output_file_path=docx_out,
        items_map=docx_map,
        layout_mode="bilingual_stacked"
    )
    print(f"[OK] Word output created: {docx_out} (size: {os.path.getsize(docx_out)} bytes)")

    # 2. 验证 Excel 真实文件
    xlsx_in = "sample_docs/工艺参数表_模具点检.xlsx"
    xlsx_out = "exported_docs/工艺参数表_模具点检_双语对照版_EN.xlsx"

    xlsx_items = XlsxParserEngine.extract_content(xlsx_in)
    print(f"Excel 提取出 {len(xlsx_items)} 个数据单元格")

    xlsx_map = {}
    for it in xlsx_items:
        src = it["source_text"]
        if "精密注塑车间" in src:
            xlsx_map[it["id"]] = "Precision Injection Workshop - Daily Machine Parameter Checklist"
        elif "点检项目" in src:
            xlsx_map[it["id"]] = "Check Item"
        elif "标准工艺要求" in src:
            xlsx_map[it["id"]] = "Standard Process Requirement"
        elif "检查频率" in src:
            xlsx_map[it["id"]] = "Frequency"
        elif "责任人" in src:
            xlsx_map[it["id"]] = "Responsible Person & Criteria"
        elif "射胶压力与保压参数" in src:
            xlsx_map[it["id"]] = "Injection Pressure & Holding Time"
        elif "射胶压力保持" in src:
            xlsx_map[it["id"]] = "Maintain injection pressure at 125 MPa, holding time 8.5 seconds."
        elif "每班开机首检" in src:
            xlsx_map[it["id"]] = "First check at start of each shift"
        elif "读取控制柜" in src:
            xlsx_map[it["id"]] = "Operator reads control cabinet meter and logs data"
        elif "模具预热温度" in src:
            xlsx_map[it["id"]] = "Mold Preheating Temperature"
        elif "前模 90 ℃" in src:
            xlsx_map[it["id"]] = "Front mold 90 ℃, rear mold 85 ℃, temp difference ≤ 3 ℃."
        elif "防错治具及夹具定位" in src:
            xlsx_map[it["id"]] = "Poka-Yoke Fixture & Jig Locating"
        elif "治具定位销完好" in src:
            xlsx_map[it["id"]] = "Fixture pins intact, error-proofing sensor light ON."
        elif "外观品质检验" in src:
            xlsx_map[it["id"]] = "Visual Quality Inspection"
        elif "确认成型品无飞边" in src:
            xlsx_map[it["id"]] = "Confirm molded parts are free of flash, air marks, sink marks, or burrs."
        else:
            xlsx_map[it["id"]] = f"EN: {src}"

    XlsxParserEngine.rebuild_bilingual_doc(
        original_file_path=xlsx_in,
        output_file_path=xlsx_out,
        items_map=xlsx_map,
        layout_mode="bilingual_stacked"
    )
    print(f"[OK] Excel output created: {xlsx_out} (size: {os.path.getsize(xlsx_out)} bytes)")

if __name__ == "__main__":
    run_verification()
