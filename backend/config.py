import os
import sys
import json
from pathlib import Path

# 判断是否在 PyInstaller 打包环境下运行
if getattr(sys, 'frozen', False):
    # 打包运行环境：内部打包资源在 _MEIPASS，用户数据在 exe 所在目录
    RESOURCE_DIR = Path(getattr(sys, '_MEIPASS', os.path.dirname(sys.executable)))
    APP_DIR = Path(os.path.dirname(sys.executable))
else:
    # 源码运行环境
    RESOURCE_DIR = Path(__file__).resolve().parent.parent
    APP_DIR = RESOURCE_DIR

BASE_DIR = RESOURCE_DIR
FRONTEND_DIR = RESOURCE_DIR / "frontend"
SAMPLE_DOCS_DIR = RESOURCE_DIR / "sample_docs"

DATA_DIR = APP_DIR / "data"
EXPORT_DIR = APP_DIR / "exported_docs"
UPLOAD_DIR = APP_DIR / "uploads"

DATA_DIR.mkdir(parents=True, exist_ok=True)
EXPORT_DIR.mkdir(parents=True, exist_ok=True)
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

DB_PATH = DATA_DIR / "transfiler.db"

# 默认支持的语种列表
SUPPORTED_LANGUAGES = [
    {"code": "zh", "name": "中文 (简体中文 / Chinese)", "flag": "🇨🇳"},
    {"code": "en", "name": "英语 (English)", "flag": "🇺🇸"},
    {"code": "vi", "name": "越南语 (Tiếng Việt)", "flag": "🇻🇳"},
    {"code": "th", "name": "泰语 (ภาษาไทย)", "flag": "🇹🇭"},
    {"code": "id", "name": "印尼语 (Bahasa Indonesia)", "flag": "🇮🇩"},
    {"code": "ja", "name": "日语 (日本語)", "flag": "🇯🇵"},
    {"code": "ko", "name": "韩语 (한국어)", "flag": "🇰🇷"},
    {"code": "es", "name": "西班牙语 (Español)", "flag": "🇪🇸"},
    {"code": "de", "name": "德语 (Deutsch)", "flag": "🇩🇪"},
    {"code": "fr", "name": "法语 (Français)", "flag": "🇫🇷"},
    {"code": "ru", "name": "俄语 (Русский)", "flag": "🇷🇺"},
]
