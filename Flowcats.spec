# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_data_files

datas = [('logo_flowcats.png', '.')]
if os.path.exists('TEMAS DEL DÍA.xlsx'):
    datas.append(('TEMAS DEL DÍA.xlsx', '.'))

datas += collect_data_files('customtkinter')

a = Analysis(
    ['automatizacion_santamaria.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[
        'openpyxl',
        'requests',
        'lxml',
        'lxml.etree',
        'bs4',
        'customtkinter',
        'PIL',
        'PIL.Image',
        'Pillow',
        'groq',
        'httpx',
        'pydantic',
        'anyio'
    ],
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
    name='Flowcats',
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
    icon=['logo_flowcats.png'],
)
