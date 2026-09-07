from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from typing import Dict, Any, Optional
from backend.config import SUPPORTED_LANGUAGES
from backend.db.database import get_all_settings, update_settings_batch, update_setting
from backend.core.translator import LLMTranslator

router = APIRouter(prefix="/api/settings", tags=["settings"])

class SettingsUpdateModel(BaseModel):
    provider: Optional[str] = None
    api_key: Optional[str] = None
    base_url: Optional[str] = None
    model_name: Optional[str] = None
    export_dir: Optional[str] = None
    font_zh: Optional[str] = None
    font_en: Optional[str] = None
    layout_mode: Optional[str] = None
    source_lang: Optional[str] = None
    target_lang: Optional[str] = None
    use_glossary: Optional[str] = None
    theme: Optional[str] = None

class TestConnectionModel(BaseModel):
    api_key: str
    base_url: str
    model_name: str

@router.get("")
def get_settings():
    settings = get_all_settings()
    # 隐藏部分 key 字符保护隐私 (仅在返回时用于展示，完整 key 保持在 settings['api_key'])
    key = settings.get("api_key", "")
    masked_key = key[:4] + "*" * max(0, len(key) - 8) + key[-4:] if len(key) >= 8 else key
    return {
        "status": "success",
        "data": {
            **settings,
            "masked_key": masked_key
        }
    }

@router.post("")
def save_settings(data: SettingsUpdateModel):
    update_dict = {k: v for k, v in data.model_dump(exclude_unset=True).items() if v is not None}
    update_settings_batch(update_dict)
    return {"status": "success", "message": "配置已成功保存", "data": update_dict}

@router.post("/test-connection")
async def test_connection(data: TestConnectionModel):
    translator = LLMTranslator(api_key=data.api_key, base_url=data.base_url, model_name=data.model_name)
    res = await translator.test_connection()
    return res

@router.get("/check-connection")
async def check_current_connection():
    """使用当前持久化保存的设置执行开机连通性检测"""
    settings = get_all_settings()
    api_key = settings.get("api_key", "").strip()
    base_url = settings.get("base_url", "https://api.deepseek.com/v1")
    model_name = settings.get("model_name", "deepseek-chat")
    provider = settings.get("provider", "deepseek")

    if not api_key and provider != "ollama":
        return {
            "success": False,
            "status": "unconfigured",
            "message": "未配置 API Key，当前为本地模拟模式",
            "provider": provider,
            "model_name": model_name
        }

    translator = LLMTranslator(api_key=api_key, base_url=base_url, model_name=model_name)
    res = await translator.test_connection()
    res["provider"] = provider
    res["model_name"] = model_name
    return res

@router.get("/languages")
def get_languages():
    return {"status": "success", "data": SUPPORTED_LANGUAGES}
