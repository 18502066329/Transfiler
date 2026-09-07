import os
import subprocess
from pathlib import Path
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from typing import Optional
from backend.db.database import get_history_tasks

router = APIRouter(prefix="/api/history", tags=["history"])

class OpenFolderModel(BaseModel):
    file_path: Optional[str] = None

@router.get("")
def list_history():
    tasks = get_history_tasks(limit=50)
    return {"status": "success", "data": tasks}

@router.post("/open-folder")
def open_folder(data: OpenFolderModel):
    """在 Windows 资源管理器中高亮打开文件或文件夹"""
    target = data.file_path
    if not target or not os.path.exists(target):
        raise HTTPException(status_code=404, detail="指定路径不存在")
    
    try:
        if os.path.isfile(target):
            # 打开文件夹并选中文件
            subprocess.run(f'explorer /select,"{os.path.abspath(target)}"', shell=True)
        else:
            subprocess.run(f'explorer "{os.path.abspath(target)}"', shell=True)
        return {"status": "success", "message": "已在资源管理器中打开"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"打开文件夹失败: {str(e)}")
