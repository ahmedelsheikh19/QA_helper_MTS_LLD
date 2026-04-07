# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

a = Analysis(
    ['src/gui.py'],
    pathex=[os.path.abspath('src')],
    binaries=[],
    datas=[
        # bundle customtkinter's themes/assets
        *collect_data_files('customtkinter'),
    ],
    hiddenimports=[
        *collect_submodules('customtkinter'),
        'openpyxl',
        'req_loader',
        'generate_traceabilityLLD',
        'generate_ut_traceability',
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='TraceabilityGenerator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,        # no black terminal window
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='icon.ico',    # uncomment and add an .ico file if you have one
)
