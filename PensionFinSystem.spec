# -*- coding: utf-8 -*-
"""
PyInstaller spec 文件
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
        'numpy.core._multiarray_umath',
        # tkinter 模块
        'tkinter',
        'tkinter.ttk',
        'tkinter.scrolledtext',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.colorchooser',
        'tkinter.commondialog',
        'tkinter.constants',
        'tkinter.dialog',
        'tkinter.dnd',
        'tkinter.font',
        'tkinter.messagebox',
        'tkinter.simpledialog',
        'tkinter.tix',
        # 项目模块 - 关键！必须显式指定
        'gui',
        'gui.main_window',
        'gui.__init__',
        'core',
        'core.config',
        'core.processor',
        'core.__init__',
        'utils',
        'utils.excel_tool',
        'utils.init_template',
        'utils.__init__',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='PensionFinSystem',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    name='PensionFinSystem',
)
