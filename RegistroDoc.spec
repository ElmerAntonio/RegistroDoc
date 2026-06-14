# -*- mode: python ; coding: utf-8 -*-
import os
import sys
import customtkinter

# Localizar la ruta del paquete customtkinter para incluir sus recursos gráficos y temas
ctk_dir = os.path.dirname(customtkinter.__file__)

block_cipher = None

a = Analysis(
    ['src/app.py'],
    pathex=['src'],
    binaries=[],
    datas=[
        ('assets', 'assets'),
        (os.path.join(ctk_dir, 'assets'), 'customtkinter/assets'),
    ],
    hiddenimports=[
        'PIL', 
        'openpyxl', 
        'cryptography', 
        'cryptography.fernet', 
        'win32com', 
        'win32com.client',
        'sqlite3'
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
    name='RegistroDoc',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/icon_fixed.ico'
)
