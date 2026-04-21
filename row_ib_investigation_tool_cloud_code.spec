# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['row_ib_investigation_tool_cloud_code.py'],
    pathex=['C:\\Users\\Mukesh_Maruthi\\MFI_Tool'],
    binaries=[],
    datas=[],
    hiddenimports=['pandas', 'openpyxl', 'openpyxl.styles', 'tkinter', 'ttkthemes'],
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Row_IB_Investigation_Tool_v4.9',
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
    version='version_info.txt',
)
