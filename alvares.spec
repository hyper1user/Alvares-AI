# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec-файл для АЛЬВАРЕС AI"""

import os

block_cipher = None

a = Analysis(
    ['gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('templates/rozp_template.docx', 'templates'),
        ('emblem.png', '.'),
    ],
    hiddenimports=[
        'PIL',
        'PIL._tkinter_finder',
        'openpyxl',
        'docx',
        'data',
        'data.database',
        'core',
        'core.br_roles',
        'path_utils',
        'generate_reports',
        'br_calculator',
        'br_updater',
        'month_utils',
        'tabel_filler',
        'excel_processor',
        'excel_reports',
        'word_generator',
        'version',
        'updater',
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
    name='Alvares',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # windowed mode, без консолі
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Alvares',
)
