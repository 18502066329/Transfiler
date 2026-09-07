import os
import csv
from typing import List, Dict, Any
from backend.core.parser_xlsx import is_person_name_text

class CsvParserEngine:
    """CSV 表格高保真解析与重构引擎"""

    @staticmethod
    def detect_encoding(file_path: str) -> str:
        """自动探测 CSV 编码 (utf-8-sig / gbk / utf-8 / latin-1)"""
        encodings = ['utf-8-sig', 'utf-8', 'gbk', 'gb2312', 'latin-1']
        with open(file_path, 'rb') as f:
            raw = f.read(10000)
            for enc in encodings:
                try:
                    raw.decode(enc)
                    return enc
                except UnicodeDecodeError:
                    continue
        return 'utf-8-sig'

    @staticmethod
    def extract_content(file_path: str) -> List[Dict[str, Any]]:
        enc = CsvParserEngine.detect_encoding(file_path)
        items = []
        with open(file_path, "r", encoding=enc, newline="") as f:
            reader = csv.reader(f)
            for r_idx, row in enumerate(reader):
                for c_idx, cell_text in enumerate(row):
                    text = cell_text.strip()
                    if text and not text.replace('.', '', 1).isdigit():
                        is_name, name_val = is_person_name_text(text)
                        item_id = f"csv_r_{r_idx}_c_{c_idx}"
                        items.append({
                            "id": item_id,
                            "type": "csv_cell",
                            "row": r_idx,
                            "col": c_idx,
                            "source_text": text,
                            "target_text": text if is_name and text == name_val else "",
                            "is_person_name": is_name,
                            "matched_terms": []
                        })
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
        if layout_mode == "bilingual_stacked":
            layout_mode = "zh_top"

        enc = CsvParserEngine.detect_encoding(original_file_path)
        rows_out = []
        with open(original_file_path, "r", encoding=enc, newline="") as f:
            reader = csv.reader(f)
            for r_idx, row in enumerate(reader):
                new_row = []
                for c_idx, cell_text in enumerate(row):
                    item_id = f"csv_r_{r_idx}_c_{c_idx}"
                    trans_text = items_map.get(item_id, "").strip()
                    if trans_text:
                        if layout_mode == "replace":
                            new_row.append(trans_text)
                        elif layout_mode == "zh_top":
                            if source_lang == "zh":
                                new_row.append(f"{cell_text}\n{trans_text}")
                            else:
                                new_row.append(f"{trans_text}\n{cell_text}")
                        elif layout_mode == "zh_bottom":
                            if source_lang == "zh":
                                new_row.append(f"{trans_text}\n{cell_text}")
                            else:
                                new_row.append(f"{cell_text}\n{trans_text}")
                    else:
                        new_row.append(cell_text)
                rows_out.append(new_row)

        os.makedirs(os.path.dirname(output_file_path), exist_ok=True)
        with open(output_file_path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerows(rows_out)

        return output_file_path
