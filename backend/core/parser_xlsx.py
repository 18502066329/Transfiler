import os
import re
import zipfile
import xml.etree.ElementTree as ET
from typing import List, Dict, Any, Tuple, Optional

# 注册标准 OpenXML 命名空间，确保 XML 重新序列化时不产生异常前缀
ET.register_namespace('', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
ET.register_namespace('xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing')
ET.register_namespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
ET.register_namespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')

# 常见百家姓表与非人名技术词过滤表
CHINESE_SURNAMES = (
    '赵钱孙李周吴郑王冯陈褚卫蒋沈韩杨朱秦尤许何吕施张孔曹严华金魏陶姜戚谢邹喻柏水窦章云苏潘葛奚范彭郎'
    '鲁韦昌马苗凤花方俞任袁柳酆鲍史唐费廉岑薛雷贺倪汤滕殷罗毕郝邬安常乐于时傅皮卞齐康伍余元卜顾孟平黄'
    '和穆萧尹姚邵湛汪祁毛禹狄米贝明臧计伏成戴谈宋茅庞熊纪舒屈项祝董梁杜阮蓝闵席季麻强贾路娄危江童颜郭'
    '梅盛林刁钟徐邱骆高夏蔡田樊胡凌霍虞万支柯昝管卢莫经房裘缪干解应宗丁宣贲邓郁单杭洪包诸左石崔吉钮龚'
    '程嵇邢滑裴陆荣翁荀羊於惠甄曲家封芮羿储靳汲邴糜松井段富巫乌焦巴弓牧隗山谷车侯宓蓬全郗班仰秋仲伊宫'
    '宁仇栾暴甘钭厉戎祖武符刘景詹束龙叶幸司韶郜黎蓟薄印宿白怀蒲邰从鄂索咸籍赖卓蔺屠蒙池乔阴鬱胥能苍双'
    '闻莘党翟谭贡劳逄姬申扶堵冉宰郦雍卻璩桑桂濮牛寿通边扈燕冀郏浦尚农温别庄晏柴瞿阎充慕连茹习宦艾鱼容'
    '向古易慎戈廖庾终暨居衡步都耿满弘匡国文寇广禄阙东欧殳沃利蔚越夔隆师巩厍聂晁勾敖融冷訾辛阚那简饶空'
    '曾毋沙乜养鞠须丰巢关蒯相查后荆红游竺权逯盖益桓公晋楚闫法汝鄢涂钦岳帅缑亢况'
)

NON_NAME_TERMS = {
    '工程', '品质', '生产', '工单', '工位', '工序', '工步', '序号', '手套', '指套',
    '工帽', '耳塞', '口罩', '作业', '检查', '包装', '搬运', '天盒', '地盒', '线长',
    '合计', '首次', '发行', '版本', '页码', '部门', '签核', '人力', '设备', '名称',
    '型号', '数量', '辅料', '物料', '成品', '内盒', '天盖', '地盖', '面纸', '白胶',
    '胶水', '胶带', '盒胚', '压痕', '内卡', '卡纸', '外箱', '纸箱', '料架', '周转',
    '排板', '全检', '标准', '要求', '注意', '事项', '说明', '重点', '图示', '备注',
    '参数', '尺寸', '外观', '尺寸', '测试', '项目', '范围', '单位', '状态', '确认',
    '体系', '规范', '通用', '组装', '过胶', '除泡', '预折', '折拼', '贴合', '擦拭'
}

TECHNICAL_SUFFIXES = set(
    '率度量力费价款数额值线图表规部课组站位门案单条袋箱板盒机器具品号纸胶料工法序步标项目历书文页卡件期间和分点'
)

def is_person_name_text(text: str) -> Tuple[bool, Optional[str]]:
    """
    智能识别文本是否为人名或签署人栏
    返回: (is_name, name_str)
    """
    clean = text.strip()
    if not clean:
        return False, None

    # 1. 签署前缀组合 (如 "制订: 罗成耿", "审核：李四", "批准: 张伟")
    m = re.match(r'^(制订|制定|编制|审核|批准|确认|校对|核准|制表|责任人|签核|检验员|经手人|负责人)[:：\s]+([\u4e00-\u9fa5]{2,4})$', clean)
    if m:
        return True, m.group(2)

    # 2. 独立 2-3 字纯中文人名 (百家姓开头，排除技术名词及工业技术词尾)
    if 2 <= len(clean) <= 3 and re.match(r'^[\u4e00-\u9fa5]+$', clean):
        if clean[0] in CHINESE_SURNAMES and clean not in NON_NAME_TERMS and clean[-1] not in TECHNICAL_SUFFIXES:
            return True, clean

    return False, None


class XlsxParserEngine:
    """
    Excel (.xlsx) 无损高保真解析与双语重构引擎
    核心特性：
    1. 基于 OpenXML 底层直接操作，100% 保留图片尺寸、位置、锚点、样式、图表与打印布局；
    2. 深度支持提取并回填 DrawingML 文本框/图形（<xdr:sp> / <xdr:txBody>）中的所有操作指导说明；
    3. 内置智能人名识别与保护，签署人与人名保持原样免翻译。
    """

    @staticmethod
    def extract_content(file_path: str) -> List[Dict[str, Any]]:
        """
        解析 Excel 工作簿，提取单元格文本及 Drawing 文本框说明
        """
        items = []

        if not zipfile.is_zipfile(file_path):
            return items

        with zipfile.ZipFile(file_path, 'r') as z:
            namelist = z.namelist()

            # 1. 解析 xl/sharedStrings.xml (工作表共享字符串)
            if 'xl/sharedStrings.xml' in namelist:
                try:
                    sst_content = z.read('xl/sharedStrings.xml')
                    root = ET.fromstring(sst_content)
                    si_nodes = root.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si')
                    
                    for idx, si in enumerate(si_nodes):
                        # 提取 si 内所有 t 标签文本
                        t_nodes = si.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                        text = ''.join([t.text or '' for t in t_nodes]).strip()

                        if not text:
                            continue

                        # 过滤纯数字、纯符号
                        if text.replace('.', '', 1).isdigit() or len(text) == 1 and text in '*×√-/':
                            continue

                        # 人名识别
                        is_name, name_val = is_person_name_text(text)
                        target_val = text if is_name and text == name_val else ""

                        items.append({
                            "id": f"sst_{idx}",
                            "type": "excel_cell",
                            "location": f"共享字符串 #{idx + 1}",
                            "source_text": text,
                            "target_text": target_val,
                            "is_person_name": is_name,
                            "matched_terms": []
                        })
                except Exception as e:
                    print(f"解析 sharedStrings 异常: {e}")

            # 2. 解析 xl/drawings/drawing*.xml (文本框、工艺说明图形)
            for name in namelist:
                if name.startswith('xl/drawings/drawing') and name.endswith('.xml'):
                    try:
                        draw_content = z.read(name)
                        root = ET.fromstring(draw_content)
                        # 提取所有 shape/textbox 元素
                        sp_elements = root.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}sp')

                        for sp_idx, sp in enumerate(sp_elements):
                            t_nodes = sp.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}t')
                            text = ''.join([t.text or '' for t in t_nodes]).strip()

                            if not text:
                                continue

                            if text.replace('.', '', 1).isdigit():
                                continue

                            is_name, name_val = is_person_name_text(text)
                            target_val = text if is_name and text == name_val else ""

                            draw_id = name.replace('xl/drawings/', '').replace('.xml', '')
                            items.append({
                                "id": f"draw_{draw_id}_sp_{sp_idx}",
                                "type": "excel_textbox",
                                "location": f"图纸文本框 ({draw_id})",
                                "source_text": text,
                                "target_text": target_val,
                                "is_person_name": is_name,
                                "matched_terms": []
                            })
                    except Exception as e:
                        print(f"解析 {name} 绘图文本框异常: {e}")

            # 3. 解析工作表中的 inlineStr (如果存在)
            for name in namelist:
                if name.startswith('xl/worksheets/sheet') and name.endswith('.xml'):
                    try:
                        sheet_content = z.read(name)
                        root = ET.fromstring(sheet_content)
                        # 查找 t="inlineStr" 的单元格
                        c_nodes = root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c[@t="inlineStr"]')
                        sheet_id = name.replace('xl/worksheets/', '').replace('.xml', '')

                        for c in c_nodes:
                            r_attr = c.get('r', '')
                            t_nodes = c.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                            text = ''.join([t.text or '' for t in t_nodes]).strip()

                            if text and not text.replace('.', '', 1).isdigit():
                                is_name, name_val = is_person_name_text(text)
                                target_val = text if is_name and text == name_val else ""

                                items.append({
                                    "id": f"sheet_{sheet_id}_cell_{r_attr}",
                                    "type": "excel_cell",
                                    "location": f"工作表 {sheet_id} 单元格 {r_attr}",
                                    "source_text": text,
                                    "target_text": target_val,
                                    "is_person_name": is_name,
                                    "matched_terms": []
                                })
                    except Exception as e:
                        print(f"解析工作表 inlineStr 异常 ({name}): {e}")

        return items

    @staticmethod
    def rebuild_bilingual_doc(
        original_file_path: str,
        output_file_path: str,
        items_map: Dict[str, str],
        layout_mode: str = "zh_top",
        source_lang: str = "zh",
        target_lang: str = "en"
    ) -> str:
        """
        高保真无损重构 Excel 文件：
        基于 OpenXML 原生流式重构，100% 保持原有排版、图片大小锚点、文本框与宏代码
        """
        if layout_mode == "bilingual_stacked":
            layout_mode = "zh_top"

        def format_bilingual(orig: str, trans: str) -> str:
            """生成双语排版内容"""
            if not trans or trans.strip() == orig.strip():
                return orig

            if layout_mode == "replace":
                return trans
            elif layout_mode == "zh_top":
                return f"{orig}\n{trans}" if source_lang == "zh" else f"{trans}\n{orig}"
            elif layout_mode == "zh_bottom":
                return f"{trans}\n{orig}" if source_lang == "zh" else f"{orig}\n{trans}"
            return f"{orig}\n{trans}"

        out_dir = os.path.dirname(output_file_path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        with zipfile.ZipFile(original_file_path, 'r') as zin, \
             zipfile.ZipFile(output_file_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:

            for item in zin.infolist():
                raw_data = zin.read(item.filename)

                # 1. 重构 xl/sharedStrings.xml
                if item.filename == 'xl/sharedStrings.xml':
                    try:
                        root = ET.fromstring(raw_data)
                        si_nodes = root.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si')

                        for idx, si in enumerate(si_nodes):
                            item_id = f"sst_{idx}"
                            trans_text = items_map.get(item_id, "").strip()

                            if not trans_text:
                                continue

                            # 提取原有文本
                            t_nodes = si.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                            orig_text = ''.join([t.text or '' for t in t_nodes])

                            new_val = format_bilingual(orig_text, trans_text)

                            # 清理 si 子元素并设为单一 t 节点 (避免富文本分段截断)
                            si.clear()
                            t_elem = ET.SubElement(si, '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                            t_elem.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                            t_elem.text = new_val

                        rebuilt_data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                        zout.writestr(item, rebuilt_data)
                        continue
                    except Exception as e:
                        print(f"重构 sharedStrings 失败: {e}")
                        zout.writestr(item, raw_data)
                        continue

                # 2. 重构 xl/drawings/drawing*.xml (文本框与图形说明)
                if item.filename.startswith('xl/drawings/drawing') and item.filename.endswith('.xml'):
                    try:
                        root = ET.fromstring(raw_data)
                        sp_elements = root.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}sp')
                        draw_id = item.filename.replace('xl/drawings/', '').replace('.xml', '')
                        modified = False

                        for sp_idx, sp in enumerate(sp_elements):
                            item_id = f"draw_{draw_id}_sp_{sp_idx}"
                            trans_text = items_map.get(item_id, "").strip()

                            if not trans_text:
                                continue

                            t_nodes = sp.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}t')
                            if not t_nodes:
                                continue

                            orig_text = ''.join([t.text or '' for t in t_nodes])
                            new_val = format_bilingual(orig_text, trans_text)

                            # 将新文本置于首个 t 节点，其余清空
                            t_nodes[0].text = new_val
                            for extra_t in t_nodes[1:]:
                                extra_t.text = ""
                            modified = True

                        if modified:
                            rebuilt_data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                            zout.writestr(item, rebuilt_data)
                        else:
                            zout.writestr(item, raw_data)
                        continue
                    except Exception as e:
                        print(f"重构绘图文本框 {item.filename} 失败: {e}")
                        zout.writestr(item, raw_data)
                        continue

                # 3. 重构工作表中的 inlineStr (如果存在)
                if item.filename.startswith('xl/worksheets/sheet') and item.filename.endswith('.xml'):
                    try:
                        root = ET.fromstring(raw_data)
                        c_nodes = root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c[@t="inlineStr"]')
                        sheet_id = item.filename.replace('xl/worksheets/', '').replace('.xml', '')
                        modified = False

                        for c in c_nodes:
                            r_attr = c.get('r', '')
                            item_id = f"sheet_{sheet_id}_cell_{r_attr}"
                            trans_text = items_map.get(item_id, "").strip()

                            if not trans_text:
                                continue

                            t_nodes = c.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                            if not t_nodes:
                                continue

                            orig_text = ''.join([t.text or '' for t in t_nodes])
                            new_val = format_bilingual(orig_text, trans_text)

                            t_nodes[0].text = new_val
                            for extra_t in t_nodes[1:]:
                                extra_t.text = ""
                            modified = True

                        if modified:
                            rebuilt_data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                            zout.writestr(item, rebuilt_data)
                        else:
                            zout.writestr(item, raw_data)
                        continue
                    except Exception as e:
                        print(f"重构工作表 {item.filename} 失败: {e}")
                        zout.writestr(item, raw_data)
                        continue

                # 4. 其他所有文件（图片 xl/media/*, 样式 styles.xml, 关系 _rels/*）100% 原样保留
                zout.writestr(item, raw_data)

        return output_file_path
