# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['serve.py'],
    pathex=[],
    binaries=[],
    datas=[('index.html', '.'), ('template.xlsx', '.')],
    hiddenimports=['openpyxl', 'openpyxl.drawing.image', 'openpyxl.utils', 'webview', 'webview.platforms.cocoa', 'proxy_tools'],
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
    name='智能报销系统',
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
    name='智能报销系统',
)
app = BUNDLE(
    coll,
    name='智能报销系统.app',
    icon=None,
    bundle_identifier=None,
)
