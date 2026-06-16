# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec 文件 - 将程序打包为独立 exe
使用方法: pyinstaller app.spec
"""

import os

block_cipher = None

# 程序所在目录
code_dir = SPECPATH

a = Analysis(
    [os.path.join(code_dir, 'app.py')],
    pathex=[code_dir],
    binaries=[],
    datas=[
        # 将模板PDF打包到exe同目录
        (os.path.join(code_dir, '勞工終止合約通知書-.pdf'), '.'),
    ],
    hiddenimports=[
        'flask',
        'openpyxl',
        'fitz',
        'pymupdf',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'numpy', 'scipy', 'pandas', 'PIL',
        'tkinter', 'IPython', 'jupyter', 'notebook',
    ],
    noarchive=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='勞工通知書生成器',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # 保留控制台窗口显示运行状态
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
