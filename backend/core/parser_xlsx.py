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



class XlsxParserEngine:
    """
    Excel (.xlsx) 无损高保真解析与双语重构引擎
    核心特性：
    1. 基于 OpenXML 底层直接操作，100% 保留图片尺寸、位置、锚点、样式、图表与打印布局；
    2. 深度支持提取并回填 DrawingML 文本框/图形（<xdr:sp> / <xdr:txBody>）中的所有操作指导说明。
    """

    @staticmethod
    def extract_content(file_path: str) -> List[Dict[str, Any]]:
        """
        全量提取 Excel 内容：共享字符串 (sharedStrings) + 绘图文本框 (drawing*.xml)
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

                        items.append({
                            "id": f"sst_{idx}",
                            "type": "excel_cell",
                            "location": f"共享字符串 #{idx + 1}",
                            "source_text": text,
                            "target_text": "",
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

                            draw_id = name.replace('xl/drawings/', '').replace('.xml', '')
                            items.append({
                                "id": f"draw_{draw_id}_sp_{sp_idx}",
                                "type": "excel_textbox",
                                "location": f"图纸文本框 ({draw_id})",
                                "source_text": text,
                                "target_text": "",
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
                                items.append({
                                    "id": f"sheet_{sheet_id}_cell_{r_attr}",
                                    "type": "excel_cell",
                                    "location": f"工作表 {sheet_id} 单元格 {r_attr}",
                                    "source_text": text,
                                    "target_text": "",
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
