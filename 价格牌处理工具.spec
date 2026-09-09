# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = ['fitz', 'PIL', 'PIL._tkinter_finder', 'openpyxl', 'xlrd']
hiddenimports += collect_submodules('core')


a = Analysis(
    ['C:\\Users\\csei\\Desktop\\价格牌处理软件\\第二版-Python运行版\\V1.6\\biaoqianchuli_kaifa\\main.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\csei\\Desktop\\价格牌处理软件\\第二版-Python运行版\\V1.6\\biaoqianchuli_kaifa\\assets', 'assets'), ('C:\\Users\\csei\\Desktop\\价格牌处理软件\\第二版-Python运行版\\V1.6\\biaoqianchuli_kaifa\\vendor', 'vendor')],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'numpy', 'scipy', 'pandas', 'IPython', 'pytest', 'notebook', 'jupyter'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='价格牌处理工具',
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
    icon=['C:\\Users\\csei\\Desktop\\价格牌处理软件\\第二版-Python运行版\\V1.6\\biaoqianchuli_kaifa\\assets\\logo.ico'],
)
