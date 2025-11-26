# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = [
    'playwright.sync_api',
    'playwright.__main__',
    'baseopensdk',
    'baseopensdk.api.base.v1.model.app_table_record',
    'baseopensdk.api.base.v1.model.create_app_table_record_request',
    'pystray',
    'PIL',
    'PIL.Image',
    'PIL.ImageDraw'
]

# Collect all necessary packages
packages = ['playwright', 'bs4', 'openpyxl', 'baseopensdk', 'pystray', 'PIL']
for package in packages:
    try:
        tmp_ret = collect_all(package)
        datas += tmp_ret[0]
        binaries += tmp_ret[1]
        hiddenimports += tmp_ret[2]
    except Exception as e:
        print(f"Warning: Failed to collect {package}: {e}")

a = Analysis(
    ['weibo_hot_gui.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='WeiboHotGUI',
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
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='WeiboHotGUI',
)
