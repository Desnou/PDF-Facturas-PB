# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from pathlib import Path
import site

# Buscar tkinterdnd2 en las rutas de site-packages
def find_package(package_name):
    for p in site.getsitepackages() + [site.getusersitepackages()]:
        pkg_path = Path(p) / package_name
        if pkg_path.exists():
            return pkg_path
    # Fallback al venv local
    return Path('.venv/Lib/site-packages') / package_name

tkdnd_path = find_package('tkinterdnd2')

a = Analysis(
    ['app_facturas.py'],
    pathex=[],
    binaries=[],
    datas=[
        (str(tkdnd_path), 'tkinterdnd2')
    ],
    hiddenimports=[
        # tkinterdnd2
        'tkinterdnd2',
        
        # pdfplumber y pdfminer
        'pdfplumber',
        'pdfplumber.page',
        'pdfplumber.pdf',
        'pdfplumber.utils',
        'pdfminer',
        'pdfminer.pdfparser',
        'pdfminer.pdfdocument',
        'pdfminer.pdfpage',
        'pdfminer.pdfinterp',
        'pdfminer.converter',
        'pdfminer.layout',
        'pdfminer.high_level',
        
        # pywin32 - todos los módulos necesarios
        'win32api',
        'win32con',
        'win32clipboard',
        'win32gui',
        'pywintypes',
        'pythoncom',
        
        # PIL/Pillow
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        
        # tkinter
        'tkinter',
        'tkinter.font',
        'tkinter.colorchooser',
        'tkinter.filedialog',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.scrolledtext',
        
        # Stdlib para vista previa HTML
        'webbrowser',
        'tempfile',
        
        # Otros
        'queue',
        'html.parser',
        'platform',
        're',
        'os',
        
        # pypdfium2
        'pypdfium2',
        'pypdfium2._helpers',
        'pypdfium2.raw',
        
        # cryptography (dependencia de pdfminer)
        'cryptography',
        'cffi',
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
    name='Procesador_Facturas',
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
    icon='puntobasepdf.ico',
)
