# -*- mode: python ; coding: utf-8 -*-
# DataOrganizer v2.3 - Win7 兼容打包配置
#
# 兼容性要点：
# 1. 用 Python 3.7（官方最后一个支持 Win7 的版本，3.9 起放弃 Win7）构建。
# 2. 保留 x64 架构，与 v1.2 ~ v2.2 一致（目标机是 64 位 Win7）。
# 3. 去掉 UPX 压缩：老硬件上解压慢、且部分杀软误报。
# 4. 去掉未使用的 pytesseract，减小体积。
# 5. 显式声明 gui 无控制台、单文件模式。

block_cipher = None


a = Analysis(
    ['DataOrganizer.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'openpyxl',
        'openpyxl.utils',
        'win32com.client',
        'win32com.gen_py',
        'win32com',
        'pythoncom',
        'pywintypes',
        'win32api',
        'win32con',
        'PIL',
        'PIL.Image',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'pytesseract',
        'tkinter.test',
        'test',
        'unittest',
        'pydoc_data',
        'sqlite3',
        'numpy',
        'pandas',
        'matplotlib',
        'PyQt5',
        'PySide2',
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='DataOrganizerv2.3',
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
    # icon=None,
)
