import os
import uuid
import time
import shutil
from pathlib import Path
from typing import Optional, List, Dict, Any
from fastapi import APIRouter, HTTPException, UploadFile, File, Form
from pydantic import BaseModel

from backend.config import UPLOAD_DIR, EXPORT_DIR, SUPPORTED_LANGUAGES
from backend.db.database import (
    save_draft_task,
    get_draft_task,
    update_draft_item,
    get_all_settings,
    get_enabled_glossary_dict,
    add_history_task
)
from backend.core.parser_docx import DocxParserEngine
from backend.core.parser_xlsx import XlsxParserEngine
from backend.core.parser_csv import CsvParserEngine
from backend.core.parser_pptx import PptxParserEngine
from backend.core.doc_converter import DocConverter
from backend.core.glossary_matcher import GlossaryMatcher
from backend.core.translator import LLMTranslator

router = APIRouter(prefix="/api/task", tags=["translate_task"])

class ProcessTaskModel(BaseModel):
    task_id: str
    source_lang: Optional[str] = "zh"
    target_lang: str = "en"
    layout_mode: Optional[str] = "zh_top"
    use_glossary: bool = True

class UpdateItemModel(BaseModel):
    task_id: str
    item_id: str
    target_text: str

class ExportTaskModel(BaseModel):
    task_id: str
    custom_name: Optional[str] = None
    layout_mode: Optional[str] = None

@router.post("/upload")
async def upload_document(file: UploadFile = File(...)):
    """上传文件并执行第一阶段解析"""
    filename = file.filename
    ext = Path(filename).suffix.lower()
    
    if ext not in [".docx", ".xlsx", ".csv", ".doc", ".pptx", ".ppt"]:
        raise HTTPException(status_code=400, detail="不支持的文件格式，仅支持 .docx / .xlsx / .csv / .doc / .pptx / .ppt")
    
    task_id = str(uuid.uuid4())[:8]
    save_path = UPLOAD_DIR / f"{task_id}_{filename}"
    
    with open(save_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
        
    actual_file_path = str(save_path)
    # 处理老版 .doc
    if ext == ".doc":
        try:
            actual_file_path = DocConverter.convert_doc_to_docx(actual_file_path)
            ext = ".docx"
        except Exception as e:
            raise HTTPException(status_code=400, detail=str(e))
    # 处理老版 .ppt
    elif ext == ".ppt":
        try:
            actual_file_path = DocConverter.convert_ppt_to_pptx(actual_file_path)
            ext = ".pptx"
        except Exception as e:
            raise HTTPException(status_code=400, detail=str(e))

    # 解析提取文档结构
    items = []
    file_type = "docx"
    if ext == ".docx":
        file_type = "docx"
        items = DocxParserEngine.extract_content(actual_file_path)
    elif ext == ".xlsx":
        file_type = "xlsx"
        items = XlsxParserEngine.extract_content(actual_file_path)
    elif ext == ".csv":
        file_type = "csv"
        items = CsvParserEngine.extract_content(actual_file_path)
    elif ext == ".pptx":
        file_type = "pptx"
        items = PptxParserEngine.extract_content(actual_file_path)

    # 保存初始草稿数据
    file_size = os.path.getsize(actual_file_path)
    save_draft_task(
        task_id=task_id,
        file_name=filename,
        file_path=actual_file_path,
        source_lang="zh",
        target_lang="en",
        file_type=file_type,
        raw_items=items
    )

    return {
        "status": "success",
        "task_id": task_id,
        "file_name": filename,
        "file_size": file_size,
        "file_type": file_type,
        "total_items": len(items),
        "items": items
    }

@router.post("/process")
async def process_translation(data: ProcessTaskModel):
    """执行文档智能翻译与术语库匹配 (支持双向多语言及排版规则)"""
    task = get_draft_task(data.task_id)
    if not task:
        raise HTTPException(status_code=404, detail="未找到对应的翻译任务")

    items = task["raw_items"]
    source_lang = data.source_lang or "zh"
    target_lang = data.target_lang or "en"
    layout_mode = data.layout_mode or "zh_top"
    
    # 查找源语言与目标语言名称
    source_lang_name = "中文"
    target_lang_name = "英语"
    for lang in SUPPORTED_LANGUAGES:
        if lang["code"] == source_lang:
            source_lang_name = lang["name"]
        if lang["code"] == target_lang:
            target_lang_name = lang["name"]

    # 加载系统设置与术语库
    settings = get_all_settings()
    api_key = settings.get("api_key", "")
    base_url = settings.get("base_url", "https://api.deepseek.com/v1")
    model_name = settings.get("model_name", "deepseek-chat")

    glossary_dict = {}
    if data.use_glossary:
        glossary_dict = get_enabled_glossary_dict(source_lang=source_lang, target_lang=target_lang)

    matcher = GlossaryMatcher(glossary_dict)
    translator = LLMTranslator(api_key=api_key, base_url=base_url, model_name=model_name)

    start_t = time.time()
    translated_items = await translator.translate_batch(
        items=items,
        source_lang_name=source_lang_name,
        target_lang_name=target_lang_name,
        glossary_matcher=matcher
    )
    cost_time = round(time.time() - start_t, 2)

    # 更新草稿
    save_draft_task(
        task_id=data.task_id,
        file_name=task["file_name"],
        file_path=task["file_path"],
        source_lang=source_lang,
        target_lang=target_lang,
        file_type=task["file_type"],
        raw_items=translated_items
    )

    # 统计术语命中数
    glossary_hit_count = sum(1 for it in translated_items if it.get("matched_terms"))

    return {
        "status": "success",
        "task_id": data.task_id,
        "source_lang": source_lang,
        "target_lang": target_lang,
        "layout_mode": layout_mode,
        "cost_time": cost_time,
        "total_items": len(translated_items),
        "glossary_hit_count": glossary_hit_count,
        "items": translated_items
    }

@router.get("/{task_id}")
def get_task_detail(task_id: str):
    task = get_draft_task(task_id)
    if not task:
        raise HTTPException(status_code=404, detail="任务不存在")
    return {"status": "success", "data": task}

@router.post("/update-item")
def update_item_text(data: UpdateItemModel):
    """用户在对比校对界面手动修改某个单元格/段落文字"""
    ok = update_draft_item(task_id=data.task_id, item_id=data.item_id, target_text=data.target_text)
    if not ok:
        raise HTTPException(status_code=400, detail="更新条目失败")
    return {"status": "success", "message": "条目已更新"}

@router.post("/export")
def export_bilingual_document(data: ExportTaskModel):
    """将校对后的草稿重构并导出为最终文件 (支持中文在上/中文在下排版)"""
    task = get_draft_task(data.task_id)
    if not task:
        raise HTTPException(status_code=404, detail="未找到任务草稿数据")

    settings = get_all_settings()
    export_dir_str = settings.get("export_dir", str(EXPORT_DIR))
    export_dir = Path(export_dir_str)
    export_dir.mkdir(parents=True, exist_ok=True)

    layout_mode = data.layout_mode or settings.get("layout_mode", "zh_top")
    if layout_mode == "bilingual_stacked":
        layout_mode = "zh_top"
    font_en = settings.get("font_en", "Calibri")

    file_path = task["file_path"]
    file_type = task["file_type"]
    orig_name = task["file_name"]
    name_stem = Path(orig_name).stem
    ext = Path(orig_name).suffix.lower()
    if ext == ".doc":
        ext = ".docx"
    elif ext == ".ppt":
        ext = ".pptx"

    source_lang = task.get("source_lang", "zh")
    target_lang = task.get("target_lang", "en")
    
    export_filename = data.custom_name or f"{name_stem}_双语对照版_{source_lang.upper()}_TO_{target_lang.upper()}{ext}"
    output_path = str(export_dir / export_filename)

    # 构造 item_id -> target_text 映射字典
    items_map = {it["id"]: it.get("target_text", "") for it in task["raw_items"]}

    # 调用重构引擎
    if file_type == "docx":
        DocxParserEngine.rebuild_bilingual_doc(
            original_file_path=file_path,
            output_file_path=output_path,
            items_map=items_map,
            layout_mode=layout_mode,
            source_lang=source_lang,
            target_lang=target_lang,
            font_en=font_en
        )
    elif file_type == "xlsx":
        XlsxParserEngine.rebuild_bilingual_doc(
            original_file_path=file_path,
            output_file_path=output_path,
            items_map=items_map,
            layout_mode=layout_mode,
            source_lang=source_lang,
            target_lang=target_lang
        )
    elif file_type == "csv":
        CsvParserEngine.rebuild_bilingual_doc(
            original_file_path=file_path,
            output_file_path=output_path,
            items_map=items_map,
            layout_mode=layout_mode,
            source_lang=source_lang,
            target_lang=target_lang
        )
    elif file_type == "pptx":
        PptxParserEngine.rebuild_bilingual_doc(
            original_file_path=file_path,
            output_file_path=output_path,
            items_map=items_map,
            layout_mode=layout_mode,
            source_lang=source_lang,
            target_lang=target_lang,
            font_en=font_en
        )

    # 记录到历史任务
    file_size = os.path.getsize(output_path)
    add_history_task(
        task_id=data.task_id,
        file_name=export_filename,
        file_size=file_size,
        source_lang=source_lang,
        target_lang=target_lang,
        total_items=len(task["raw_items"]),
        export_path=output_path,
        status="completed",
        cost_time=0.0
    )

    return {
        "status": "success",
        "export_filename": export_filename,
        "export_path": output_path,
        "file_size": file_size,
        "message": "双语文件已成功导出！"
    }
