# -*- mode: python ; coding: utf-8 -*-

import os

# The icon is optional. The original package did not include billing_tool.ico,
# so this spec builds successfully with or without that file.
icon_path = 'billing_tool.ico' if os.path.exists('billing_tool.ico') else None
datas = []
if icon_path:
    datas.append((icon_path, '.'))

binaries = []
hiddenimports = [
    'billing_core', 'customer_config', 'fuel_rates', 'sheet_merge', 'money', 'rate_store',
    'openpyxl', 'pandas', 'dateutil',
]

try:
    from PyInstaller.utils.hooks import collect_all
    _d, _b, _h = collect_all('tkinterdnd2')
    datas += _d
    binaries += _b
    hiddenimports += _h
except Exception:
    pass

a = Analysis(
    ['billing_tool.py'],
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
    a.binaries,
    a.datas,
    [],
    name='BillingTool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_path,
)
