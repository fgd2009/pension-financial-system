# -*- coding: utf-8 -*-
"""
PyInstaller spec 文件 - 稳定版
养老金融通报数据自动化处理系统
"""

import sys
import os

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[os.path.dirname(SPEC)],
    binaries=[],
    datas=[],
    hiddenimports=[
        # 核心依赖
        'openpyxl',
        'pandas',
        'numpy',
        'et_xmlfile',
        'pandas._libs',
        'pandas._libs.tslibs.natypes',
        'numpy.core._multiarray_umath',
        'numpy.linalg._umath_linalg',
        # tkinter 完整导入
        'tkinter',
        'tkinter.ttk',
        'tkinter.scrolledtext',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.colorchooser',
        'tkinter.commondialog',
        'tkinter.constants',
        'tkinter.dialog',
        'tkinter.font',
        'tkinter.simpledialog',
        # 项目模块
        'gui.main_window',
        'core.config',
        'core.processor',
        'utils.excel_tool',
        'utils.init_template',
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PensionFinSystem',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    name='PensionFinSystem',
)
