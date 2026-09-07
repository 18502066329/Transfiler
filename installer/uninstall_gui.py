# -*- coding: utf-8 -*-
"""
TransFiler 官方卸载程序 (Uninstaller)
用于清理桌面快捷方式、开始菜单目录、Windows 注册表信息及应用主程序。
"""

import os
import sys
import shutil
import winreg
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, messagebox

# 确定当前安装目录
if getattr(sys, 'frozen', False):
    INSTALL_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    INSTALL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

APP_NAME = "TransFiler"
APP_DISPLAY_NAME = "TransFiler 制造业文档双语智能转换系统"

def remove_shortcuts():
    """移除桌面和开始菜单快捷方式"""
    try:
        # 1. 桌面快捷方式
        desktop = os.path.join(os.environ.get('USERPROFILE', ''), 'Desktop')
        desktop_lnk = os.path.join(desktop, f"{APP_NAME}.lnk")
        if os.path.exists(desktop_lnk):
            os.remove(desktop_lnk)

        # 2. 开始菜单目录
        appdata = os.environ.get('APPDATA', '')
        start_menu_folder = os.path.join(appdata, r'Microsoft\Windows\Start Menu\Programs', APP_NAME)
        if os.path.exists(start_menu_folder):
            shutil.rmtree(start_menu_folder, ignore_errors=True)
    except Exception as e:
        print(f"删除快捷方式异常: {e}")

def remove_registry():
    """从 Windows 控制面板/应用列表中移除注册信息"""
    try:
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\TransFiler"
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path)
    except FileNotFoundError:
        pass
    except Exception as e:
        print(f"清理注册表异常: {e}")

def self_delete_and_exit(clean_data=False):
    """卸载并清理文件"""
    remove_shortcuts()
    remove_registry()

    # 清理主程序和附属文件
    files_to_remove = [
        "TransFiler.exe",
        "app_icon.ico",
        "更新日志.md",
        "sample_docs"
    ]
    for item in files_to_remove:
        path = os.path.join(INSTALL_DIR, item)
        if os.path.isfile(path):
            try:
                os.remove(path)
            except Exception:
                pass
        elif os.path.isdir(path):
            shutil.rmtree(path, ignore_errors=True)

    if clean_data:
        # 如果勾选了清理历史数据与术语库
        for folder in ["data", "uploads", "exported_docs"]:
            p = os.path.join(INSTALL_DIR, folder)
            if os.path.exists(p):
                shutil.rmtree(p, ignore_errors=True)

    # 通过自删除批处理清理残留文件与 uninstall.exe 自身
    bat_content = f"""@echo off
timeout /t 1 /nobreak >nul
del /f /q "{os.path.join(INSTALL_DIR, 'uninstall.exe')}" 2>nul
rmdir /s /q "{INSTALL_DIR}" 2>nul
del "%~f0"
"""
    bat_path = os.path.join(os.environ.get('TEMP', '.'), 'transfiler_uninst.bat')
    try:
        with open(bat_path, 'w', encoding='gbk') as f:
            f.write(bat_content)
        subprocess.Popen(bat_path, shell=True, creationflags=0x08000000) # CREATE_NO_WINDOW
    except Exception:
        pass

    sys.exit(0)

class UninstallerApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_NAME} 卸载向导")
        self.root.geometry("480x280")
        self.root.resizable(False, False)
        self.root.configure(bg="#0f172a")

        # 居中显示
        self.center_window()

        self.setup_ui()

    def center_window(self):
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (w // 2)
        y = (self.root.winfo_screenheight() // 2) - (h // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def setup_ui(self):
        # 顶栏
        header = tk.Frame(self.root, bg="#1e293b", height=70)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        title = tk.Label(header, text=f"卸载 {APP_DISPLAY_NAME}", font=("Microsoft YaHei UI", 12, "bold"), fg="#f8fafc", bg="#1e293b")
        title.pack(anchor="w", padx=20, pady=(15, 2))

        subtitle = tk.Label(header, text="从您的计算机中彻底移除该应用程序", font=("Microsoft YaHei UI", 9), fg="#94a3b8", bg="#1e293b")
        subtitle.pack(anchor="w", padx=20)

        # 内容区
        content = tk.Frame(self.root, bg="#0f172a")
        content.pack(fill="both", expand=True, padx=25, pady=20)

        self.prompt_label = tk.Label(content, text="确定要完全移除 TransFiler 及其所有组件吗？", font=("Microsoft YaHei UI", 10), fg="#e2e8f0", bg="#0f172a")
        self.prompt_label.pack(anchor="w", pady=(0, 15))

        self.clean_data_var = tk.BooleanVar(value=False)
        self.chk_data = tk.Checkbutton(
            content,
            text="同时删除本地专业术语库与交付文档 (data / exported_docs 目录)",
            variable=self.clean_data_var,
            font=("Microsoft YaHei UI", 9),
            fg="#cbd5e1",
            bg="#0f172a",
            selectcolor="#1e293b",
            activebackground="#0f172a",
            activeforeground="#38bdf8"
        )
        self.chk_data.pack(anchor="w")

        # 进度条 (初始隐藏)
        self.progress = ttk.Progressbar(content, mode="determinate", length=430)

        # 底部按钮区
        footer = tk.Frame(self.root, bg="#0f172a", height=50)
        footer.pack(fill="x", side="bottom", padx=20, pady=(0, 15))

        self.btn_cancel = tk.Button(
            footer,
            text="取消",
            command=self.root.destroy,
            font=("Microsoft YaHei UI", 9),
            bg="#334155",
            fg="#f8fafc",
            activebackground="#475569",
            activeforeground="#ffffff",
            relief="flat",
            padx=15,
            pady=4,
            cursor="hand2"
        )
        self.btn_cancel.pack(side="right", padx=(10, 0))

        self.btn_uninstall = tk.Button(
            footer,
            text="立即卸载",
            command=self.start_uninstall,
            font=("Microsoft YaHei UI", 9, "bold"),
            bg="#e11d48",
            fg="#ffffff",
            activebackground="#be123c",
            activeforeground="#ffffff",
            relief="flat",
            padx=18,
            pady=4,
            cursor="hand2"
        )
        self.btn_uninstall.pack(side="right")

    def start_uninstall(self):
        self.btn_uninstall.config(state="disabled", text="正在卸载...")
        self.btn_cancel.config(state="disabled")
        self.chk_data.config(state="disabled")
        self.prompt_label.config(text="正在清理应用文件、桌面快捷方式与注册表...")

        self.progress.pack(pady=(15, 0))
        self.progress.start(10)

        threading.Thread(target=self._run_uninstall_worker, daemon=True).start()

    def _run_uninstall_worker(self):
        import time
        time.sleep(1.0) # 模拟平滑进度
        clean_data = self.clean_data_var.get()
        self.root.after(0, lambda: self._on_uninstall_done(clean_data))

    def _on_uninstall_done(self, clean_data):
        self.progress.stop()
        messagebox.showinfo("卸载完成", f"{APP_DISPLAY_NAME} 已成功从您的计算机中移除。")
        self_delete_and_exit(clean_data)

if __name__ == "__main__":
    root = tk.Tk()
    app = UninstallerApp(root)
    root.mainloop()
