import os
from pathlib import Path
import docx
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

SAMPLE_DIR = Path(__file__).resolve().parent
SAMPLE_DIR.mkdir(parents=True, exist_ok=True)

def create_sample_docx():
    file_path = SAMPLE_DIR / "SOP_注塑机标准作业指导书.docx"
    doc = docx.Document()

    # 1. 标题
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run("注塑车间标准作业指导书 (SOP)")
    title_run.font.name = "Microsoft YaHei"
    title_run.font.size = Pt(16)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(30, 41, 59)

    # 2. 章节段落
    p1 = doc.add_paragraph()
    r1 = p1.add_run("一、作业前安全确认与准备")
    r1.font.bold = True
    r1.font.size = Pt(12)

    p2 = doc.add_paragraph()
    p2.add_run("1. 操作员进入车间前必须按规定穿戴耐高温手套、护目镜和防静电劳保鞋。")

    p3 = doc.add_paragraph()
    p3.add_run("2. 检查机台急停开关与安全门光电保护器，确认联动停机保护功能灵敏有效。")

    p4 = doc.add_paragraph()
    p4.add_run("3. 确认防错治具和定位夹具安装牢固，无松动或偏移。")

    # 3. 工艺参数表格
    p_tbl = doc.add_paragraph()
    p_tbl_r = p_tbl.add_run("二、注塑核心工艺参数标准表")
    p_tbl_r.font.bold = True
    p_tbl_r.font.size = Pt(12)

    table = doc.add_table(rows=5, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'

    headers = ["工艺参数项目", "设定标准值", "控制公差要求"]
    for c_idx, h in enumerate(headers):
        cell = table.cell(0, c_idx)
        cell.text = h
        cell.paragraphs[0].runs[0].font.bold = True
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        # 表头底色
        shading = parse_xml(r'<w:shd {} w:fill="E2E8F0"/>'.format(nsdecls('w')))
        cell._tc.get_or_add_tcPr().append(shading)

    rows_data = [
        ("射胶压力", "125 MPa", "± 5 MPa"),
        ("保压时间", "8.5 秒", "± 0.5 秒"),
        ("模具预热温度", "90 ℃", "± 3 ℃"),
        ("周期节拍", "22 秒", "≤ 24 秒")
    ]

    for r_idx, (col1, col2, col3) in enumerate(rows_data, start=1):
        table.cell(r_idx, 0).text = col1
        table.cell(r_idx, 1).text = col2
        table.cell(r_idx, 2).text = col3

    # 4. 品质要求
    p_qc = doc.add_paragraph()
    p_qc_r = p_qc.add_run("三、品质管控与首件检验要求")
    p_qc_r.font.bold = True
    p_qc_r.font.size = Pt(12)

    p_qc_desc = doc.add_paragraph()
    p_qc_desc.add_run("每批次开机必须执行首件全尺寸检验，重点检查产品外观是否有飞边、毛刺、缩水或缺胶不良。")

    doc.save(str(file_path))
    print(f"已创建测试样本: {file_path}")

def create_sample_xlsx():
    file_path = SAMPLE_DIR / "工艺参数表_模具点检.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "注塑机工艺点检表"

    # 样式
    font_header = Font(name="Microsoft YaHei", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    # 标题合并单元格
    ws.merge_cells("A1:D1")
    ws["A1"] = "精密注塑车间 - 机台每日工艺点检标准表"
    ws["A1"].font = Font(name="Microsoft YaHei", size=14, bold=True, color="1E293B")
    ws["A1"].alignment = align_center
    ws.row_dimensions[1].height = 32

    # 表头
    headers = ["点检项目", "标准工艺要求", "检查频率", "责任人与判定方法"]
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border
    ws.row_dimensions[2].height = 24

    # 数据行
    data = [
        ("射胶压力与保压参数", "射胶压力保持 125 MPa，保压时间 8.5 秒", "每班开机首检", "操作员读取控制柜仪表并记录"),
        ("模具预热温度", "前模 90 ℃，后模 85 ℃，温差不超过 3 ℃", "每 2 小时点检", "使用红外测温仪测量模仁温度"),
        ("防错治具及夹具定位", "治具定位销完好，防错传感器指示灯点亮", "每次换模与开机", "目视检查并进行首件检验验证"),
        ("外观品质检验", "确认成型品无飞边、气纹、缩水或严重毛刺", "每 30 分钟抽检", "质检员依据标准样件比对判定")
    ]

    for r_idx, row_values in enumerate(data, start=3):
        ws.row_dimensions[r_idx].height = 28
        for c_idx, val in enumerate(row_values, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = Font(name="Microsoft YaHei", size=10)
            cell.alignment = align_left if c_idx in [1, 2, 4] else align_center
            cell.border = thin_border

    # 设置列宽
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 32

    wb.save(str(file_path))
    print(f"已创建测试样本: {file_path}")

def create_sample_glossary_xlsx():
    file_path = SAMPLE_DIR / "常用制造业术语表_导入测试.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "术语库"

    headers = ["中文源词条 (Source)", "标准外文译文 (Target)", "适用语种 (Lang)", "工艺分类 (Category)"]
    ws.append(headers)

    terms = [
        ("注塑机", "Injection Molding Machine", "en", "设备名称"),
        ("模仁", "Mold Core", "en", "模具零件"),
        ("温控箱", "Temperature Controller", "en", "辅助设备"),
        ("脱模剂", "Release Agent", "en", "耗材辅料"),
        ("防静电劳保鞋", "Anti-Static Safety Shoes", "en", "劳保防护"),
    ]
    for t in terms:
        ws.append(t)

    wb.save(str(file_path))
    print(f"已创建测试样本: {file_path}")

def create_sample_pptx():
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor

    file_path = SAMPLE_DIR / "注塑车间生产工艺与安全培训课件.pptx"
    prs = Presentation()
    # 16:9 宽屏尺寸
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    blank_layout = prs.slide_layouts[6]

    # Slide 1: 封面幻灯片
    slide1 = prs.slides.add_slide(blank_layout)
    tx_box = slide1.shapes.add_textbox(Inches(1.5), Inches(2.0), Inches(10.33), Inches(3.0))
    tf1 = tx_box.text_frame
    tf1.word_wrap = True

    p1 = tf1.paragraphs[0]
    p1.text = "注塑车间生产工艺与安全规程培训"
    p1.font.name = "Microsoft YaHei"
    p1.font.size = Pt(36)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(30, 41, 59)
    p1.alignment = PP_ALIGN.CENTER

    p2 = tf1.add_paragraph()
    p2.text = "2026年度新员工岗位技能与标准化作业 (SOP) 专项教材"
    p2.font.size = Pt(18)
    p2.font.color.rgb = RGBColor(100, 116, 139)
    p2.alignment = PP_ALIGN.CENTER

    p3 = tf1.add_paragraph()
    p3.text = "编制审核：罗成耿"
    p3.font.size = Pt(14)
    p3.font.color.rgb = RGBColor(71, 85, 105)
    p3.alignment = PP_ALIGN.CENTER

    # Slide 2: 工艺与安全规范条目
    slide2 = prs.slides.add_slide(blank_layout)
    title_box2 = slide2.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.3), Inches(1.0))
    tf2_title = title_box2.text_frame
    p2_t = tf2_title.paragraphs[0]
    p2_t.text = "一、车间核心作业安全防护规范"
    p2_t.font.name = "Microsoft YaHei"
    p2_t.font.size = Pt(24)
    p2_t.font.bold = True

    content_box2 = slide2.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.3), Inches(4.5))
    tf2_content = content_box2.text_frame
    tf2_content.word_wrap = True

    rules = [
        "1. 操作人员进入车间必须佩戴防静电劳保鞋、安全帽与防护目镜；",
        "2. 机台运行状态下，严禁将手伸入合模机构与射胶喷嘴区域；",
        "3. 开机作业前必须测试急停开关与安全门光电保护器，确认联动停机功能灵敏；",
        "4. 发生模具报警或卡料时，必须先按下急停按钮并通知机修技术员；",
        "现场责任人：张三"
    ]
    for idx, r in enumerate(rules):
        p = tf2_content.paragraphs[0] if idx == 0 else tf2_content.add_paragraph()
        p.text = r
        p.font.size = Pt(16)
        p.space_after = Pt(12)

    # Slide 3: 表格幻灯片
    slide3 = prs.slides.add_slide(blank_layout)
    title_box3 = slide3.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.3), Inches(1.0))
    p3_t = title_box3.text_frame.paragraphs[0]
    p3_t.text = "二、核心注塑参数控制标准表"
    p3_t.font.name = "Microsoft YaHei"
    p3_t.font.size = Pt(24)
    p3_t.font.bold = True

    table_shape = slide3.shapes.add_table(rows=4, cols=3, left=Inches(1.0), top=Inches(2.2), width=Inches(11.3), height=Inches(3.5))
    tbl = table_shape.table
    tbl_data = [
        ["工艺参数名称", "设定标准值", "控制公差要求"],
        ["射胶压力 (Injection Pressure)", "120 MPa", "± 5 MPa"],
        ["模具预热温度 (Mold Temp)", "85 ℃", "± 3 ℃"],
        ["保压时间 (Holding Time)", "6.5 秒", "± 0.2 秒"]
    ]
    for r_idx, row_vals in enumerate(tbl_data):
        for c_idx, val in enumerate(row_vals):
            cell = tbl.cell(r_idx, c_idx)
            cell.text = val
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(14)
            if r_idx == 0:
                p.font.bold = True

    # 演讲者备注
    slide3.notes_slide.notes_text_frame.text = "培训主讲人注意：重点向学员强调保压时间与产品缩水缺陷的对应关系。"

    prs.save(str(file_path))
    print(f"已创建测试样本: {file_path}")

if __name__ == "__main__":
    create_sample_docx()
    create_sample_xlsx()
    create_sample_glossary_xlsx()
    create_sample_pptx()
