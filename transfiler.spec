# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from PyInstaller.utils.hooks import collect_all

block_cipher = None

# 打包静态资源目录与图标
datas = [
    ('frontend', 'frontend'),
    ('sample_docs', 'sample_docs'),
    ('app_icon.ico', '.'),
]

hiddenimports = [
    'uvicorn',
    'uvicorn.logging',
    'uvicorn.loops',
    'uvicorn.loops.auto',
    'uvicorn.protocols',
    'uvicorn.protocols.http',
    'uvicorn.protocols.http.auto',
    'uvicorn.protocols.http.h11_impl',
    'uvicorn.protocols.websockets',
    'uvicorn.protocols.websockets.auto',
    'uvicorn.lifespans',
    'uvicorn.lifespans.on',
    'uvicorn.lifespans.off',
    'fastapi',
    'starlette',
    'starlette.middleware',
    'starlette.middleware.cors',
    'starlette.staticfiles',
    'starlette.responses',
    'starlette.routing',
    'openpyxl',
    'docx',
    'pptx',
    'httpx',
    'pydantic',
    'pydantic_core',
    'webview',
    'webview.platforms.winforms',
    'clr_loader',
    'pythonnet',
    'clr',
    'backend',
    'backend.app',
    'backend.config',
    'backend.db',
    'backend.db.database',
    'backend.routers',
    'backend.routers.api_settings',
    'backend.routers.api_glossary',
    'backend.routers.api_translate',
    'backend.routers.api_history',
    'backend.core',
    'backend.core.parser_docx',
    'backend.core.parser_xlsx',
    'backend.core.parser_csv',
    'backend.core.parser_pptx',
    'backend.core.doc_converter',
    'backend.core.glossary_matcher',
    'backend.core.translator',
]

# 自动收集 uvicorn, fastapi, pywebview, pptx 数据与依赖
for pkg in ['uvicorn', 'fastapi', 'pywebview', 'pptx']:
    try:
        pkg_datas, pkg_binaries, pkg_hidden = collect_all(pkg)
        datas.extend(pkg_datas)
        hiddenimports.extend(pkg_hidden)
    except Exception as e:
        print(f"Warning collecting {pkg}: {e}")

a = Analysis(
    ['run.py'],
    pathex=['.'],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'numpy', 'scipy', 'pandas', 'IPython', 'notebook'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='TransFiler',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.ico',
)
