import sqlite3
import json
import time
from typing import List, Dict, Any, Optional
from backend.config import DB_PATH, EXPORT_DIR

def get_db():
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    cursor = conn.cursor()
    
    # 1. 系统设置表
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY,
        value TEXT
    )
    """)
    
    # 2. 制造业术语库表
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS glossary_terms (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        source_term TEXT NOT NULL,
        target_term TEXT NOT NULL,
        target_lang TEXT NOT NULL DEFAULT 'en',
        category TEXT DEFAULT '通用',
        is_enabled INTEGER DEFAULT 1,
        created_at REAL,
        updated_at REAL
    )
    """)
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_glossary_lang ON glossary_terms(target_lang, is_enabled)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_glossary_source ON glossary_terms(source_term)")

    # 3. 翻译任务历史表
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS history_tasks (
        id TEXT PRIMARY KEY,
        file_name TEXT NOT NULL,
        file_size INTEGER DEFAULT 0,
        source_lang TEXT DEFAULT 'zh',
        target_lang TEXT DEFAULT 'en',
        total_items INTEGER DEFAULT 0,
        export_path TEXT,
        status TEXT DEFAULT 'completed',
        cost_time REAL DEFAULT 0.0,
        created_at REAL
    )
    """)

    # 4. 任务实时数据暂存表 (用于校对工作台数据保存)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS draft_tasks (
        task_id TEXT PRIMARY KEY,
        file_name TEXT,
        file_path TEXT,
        source_lang TEXT,
        target_lang TEXT,
        file_type TEXT,
        raw_items TEXT,
        created_at REAL
    )
    """)

    # 插入默认设置项 (如果不存在)
    default_settings = {
        "provider": "deepseek",
        "api_key": "",
        "base_url": "https://api.deepseek.com/v1",
        "model_name": "deepseek-chat",
        "export_dir": str(EXPORT_DIR),
        "font_zh": "Microsoft YaHei",
        "font_en": "Calibri",
        "layout_mode": "zh_top",  # 中文在上，外文在下
        "source_lang": "zh",
        "target_lang": "en",
        "use_glossary": "1",
        "theme": "dark"
    }
    
    for k, v in default_settings.items():
        cursor.execute("INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)", (k, v))
        
    # 插入一批常见制造业默认术语 (如果术语库为空)
    cursor.execute("SELECT COUNT(*) as count FROM glossary_terms")
    if cursor.fetchone()["count"] == 0:
        now = time.time()
        initial_terms = [
            ("作业指导书", "Standard Operating Procedure (SOP)", "en", "体系规范"),
            ("射胶压力", "Injection Pressure", "en", "注塑工艺"),
            ("保压时间", "Holding Time", "en", "注塑工艺"),
            ("模具预热温度", "Mold Preheating Temperature", "en", "注塑工艺"),
            ("首件检验", "First Article Inspection (FAI)", "en", "品质管控"),
            ("全尺寸检验", "Full Dimension Inspection", "en", "品质管控"),
            ("防错治具", "Poka-Yoke Fixture", "en", "工装夹具"),
            ("夹具", "Fixture", "en", "工装夹具"),
            ("治具", "Jig", "en", "工装夹具"),
            ("合模机构", "Clamping Mechanism", "en", "机械结构"),
            ("顶针", "Ejector Pin", "en", "模具零件"),
            ("飞边", "Flash", "en", "缺陷术语"),
            ("毛刺", "Burr", "en", "缺陷术语"),
            ("缩水", "Sink Mark", "en", "缺陷术语"),
            ("缺胶", "Short Shot", "en", "缺陷术语"),
            ("气纹", "Air Mark", "en", "缺陷术语"),
            ("周期节拍", "Cycle Time (CT)", "en", "精益生产"),
            ("点检表", "Checklist", "en", "设备维护"),
            ("急停开关", "Emergency Stop Switch", "en", "安全防护"),
            ("光电保护器", "Safety Light Curtain", "en", "安全防护"),
        ]
        for src, tgt, lang, cat in initial_terms:
            cursor.execute(
                "INSERT INTO glossary_terms (source_term, target_term, target_lang, category, is_enabled, created_at, updated_at) VALUES (?, ?, ?, ?, 1, ?, ?)",
                (src, tgt, lang, cat, now, now)
            )

    conn.commit()
    conn.close()

# ----------------- Settings 操作 -----------------
def get_all_settings() -> Dict[str, str]:
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT key, value FROM settings")
    rows = cursor.fetchall()
    conn.close()
    return {row["key"]: row["value"] for row in rows}

def update_setting(key: str, value: str):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (key, value))
    conn.commit()
    conn.close()

def update_settings_batch(settings_dict: Dict[str, str]):
    conn = get_db()
    cursor = conn.cursor()
    for k, v in settings_dict.items():
        cursor.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (k, str(v)))
    conn.commit()
    conn.close()

# ----------------- 术语库 CRUD 操作 -----------------
def get_glossary_terms(target_lang: Optional[str] = None, category: Optional[str] = None, keyword: Optional[str] = None) -> List[Dict[str, Any]]:
    conn = get_db()
    cursor = conn.cursor()
    query = "SELECT * FROM glossary_terms WHERE 1=1"
    params = []
    
    if target_lang:
        query += " AND target_lang = ?"
        params.append(target_lang)
    if category and category != "全部":
        query += " AND category = ?"
        params.append(category)
    if keyword:
        query += " AND (source_term LIKE ? OR target_term LIKE ?)"
        params.extend([f"%{keyword}%", f"%{keyword}%"])
        
    query += " ORDER BY id DESC"
    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()
    return [dict(row) for row in rows]

def get_enabled_glossary_dict(source_lang: str = "zh", target_lang: str = "en") -> Dict[str, str]:
    conn = get_db()
    cursor = conn.cursor()
    if source_lang == "zh":
        cursor.execute(
            "SELECT source_term, target_term FROM glossary_terms WHERE is_enabled = 1 AND target_lang = ?",
            (target_lang,)
        )
        rows = cursor.fetchall()
        conn.close()
        return {row["source_term"]: row["target_term"] for row in rows}
    elif target_lang == "zh":
        # 外文翻译为中文时：将 target_term(外文) 映射为 source_term(中文)
        cursor.execute(
            "SELECT target_term, source_term FROM glossary_terms WHERE is_enabled = 1 AND target_lang = ?",
            (source_lang,)
        )
        rows = cursor.fetchall()
        conn.close()
        return {row["target_term"]: row["source_term"] for row in rows}
    else:
        # 其他语种互译时按目标语言匹配
        cursor.execute(
            "SELECT source_term, target_term FROM glossary_terms WHERE is_enabled = 1 AND target_lang = ?",
            (target_lang,)
        )
        rows = cursor.fetchall()
        conn.close()
        return {row["source_term"]: row["target_term"] for row in rows}

def add_glossary_term(source_term: str, target_term: str, target_lang: str = "en", category: str = "通用") -> int:
    conn = get_db()
    cursor = conn.cursor()
    now = time.time()
    cursor.execute(
        "INSERT INTO glossary_terms (source_term, target_term, target_lang, category, is_enabled, created_at, updated_at) VALUES (?, ?, ?, ?, 1, ?, ?)",
        (source_term.strip(), target_term.strip(), target_lang, category, now, now)
    )
    new_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return new_id

def update_glossary_term(term_id: int, source_term: str, target_term: str, target_lang: str, category: str, is_enabled: int = 1):
    conn = get_db()
    cursor = conn.cursor()
    now = time.time()
    cursor.execute(
        "UPDATE glossary_terms SET source_term = ?, target_term = ?, target_lang = ?, category = ?, is_enabled = ?, updated_at = ? WHERE id = ?",
        (source_term.strip(), target_term.strip(), target_lang, category, is_enabled, now, term_id)
    )
    conn.commit()
    conn.close()

def delete_glossary_term(term_id: int):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM glossary_terms WHERE id = ?", (term_id,))
    conn.commit()
    conn.close()

# ----------------- 任务与暂存数据操作 -----------------
def save_draft_task(task_id: str, file_name: str, file_path: str, source_lang: str, target_lang: str, file_type: str, raw_items: List[Dict[str, Any]]):
    conn = get_db()
    cursor = conn.cursor()
    now = time.time()
    cursor.execute(
        "INSERT OR REPLACE INTO draft_tasks (task_id, file_name, file_path, source_lang, target_lang, file_type, raw_items, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (task_id, file_name, file_path, source_lang, target_lang, file_type, json.dumps(raw_items, ensure_ascii=False), now)
    )
    conn.commit()
    conn.close()

def get_draft_task(task_id: str) -> Optional[Dict[str, Any]]:
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM draft_tasks WHERE task_id = ?", (task_id,))
    row = cursor.fetchone()
    conn.close()
    if not row:
        return None
    data = dict(row)
    data["raw_items"] = json.loads(data["raw_items"])
    return data

def update_draft_item(task_id: str, item_id: str, target_text: str):
    task = get_draft_task(task_id)
    if not task:
        return False
    items = task["raw_items"]
    updated = False
    for item in items:
        if item.get("id") == item_id:
            item["target_text"] = target_text
            item["is_modified"] = True
            updated = True
            break
    if updated:
        save_draft_task(task_id, task["file_name"], task["file_path"], task["source_lang"], task["target_lang"], task["file_type"], items)
    return updated

def add_history_task(task_id: str, file_name: str, file_size: int, source_lang: str, target_lang: str, total_items: int, export_path: str, status: str, cost_time: float):
    conn = get_db()
    cursor = conn.cursor()
    now = time.time()
    cursor.execute(
        "INSERT OR REPLACE INTO history_tasks (id, file_name, file_size, source_lang, target_lang, total_items, export_path, status, cost_time, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
        (task_id, file_name, file_size, source_lang, target_lang, total_items, export_path, status, cost_time, now)
    )
    conn.commit()
    conn.close()

def get_history_tasks(limit: int = 50) -> List[Dict[str, Any]]:
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM history_tasks ORDER BY created_at DESC LIMIT ?", (limit,))
    rows = cursor.fetchall()
    conn.close()
    return [dict(row) for row in rows]
