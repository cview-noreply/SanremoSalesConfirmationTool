# -*- mode: python ; coding: utf-8 -*-
# ==============================================================================
# PyInstaller build config (Tkinter Optimized)
# Project: sanremo (サンレモ成約捕捉)
# ==============================================================================

import sys
from pathlib import Path

block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=['.'],
    binaries=[],
    datas=[],
    hiddenimports=[
        'tkinter',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.scrolledtext',
        'xlwings',
        'create_sheets',
        'check_sheets',
        'create_sendlist',
        'store_documents',
        'create_alert',
        'utils',
        'yaml',
        'pandas',
        'openpyxl',
        'msoffcrypto',
        'msoffcrypto.method',
        'msoffcrypto.method.ecma376_agile',
        'msoffcrypto.method.ecma376_encrypted',
        'msoffcrypto.method.rc4',
        'msoffcrypto.method.xor',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'nicegui',
        'fastapi',
        'uvicorn',
        'starlette',
        'webview',
        'matplotlib',
        'scipy',
        'PIL',
        'cv2',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='サンレモ成約捕捉ツール',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='サンレモ成約捕捉ツール',
)