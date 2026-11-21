# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = []

# Collecter tous les modules principaux avec leurs données, binaires et imports cachés
for pkg in ['customtkinter', 'openpyxl', 'et_xmlfile', 'fpdf', 'darkdetect', 'PIL']:
    try:
        tmp_ret = collect_all(pkg)
        datas += tmp_ret[0]
        binaries += tmp_ret[1]
        hiddenimports += tmp_ret[2]
    except Exception:
        pass

# Inclure l'icône de l'application
if os.path.exists('app_icon.ico'):
    datas.append(('app_icon.ico', '.'))

# Intégrer les bibliothèques C++ Runtime (MSVC) pour garantir le fonctionnement
# sur n'importe quel autre poste Windows vierge (sans Visual C++ Redistributable installé)
sys32 = os.path.join(os.environ.get('SystemRoot', r'C:\Windows'), 'System32')
for dll in ['msvcp140.dll', 'msvcp140_1.dll', 'vcruntime140.dll', 'vcruntime140_1.dll', 'vcomp140.dll']:
    dll_path = os.path.join(sys32, dll)
    if os.path.exists(dll_path):
        binaries.append((dll_path, '.'))

# Support Tcl/Tk 9 (Python 3.15) : inclure les archives zip de bibliothèques si présentes
py_dir = os.path.dirname(sys.executable)
py_tcl = os.path.join(py_dir, 'tcl')
if os.path.isdir(py_tcl):
    for f in os.listdir(py_tcl):
        full_f = os.path.join(py_tcl, f)
        if f.endswith('.zip'):
            datas.append((full_f, 'tcl'))
        elif os.path.isdir(full_f) and f.startswith(('tk', 'tcl', 'dde', 'reg')):
            datas.append((full_f, os.path.join('tcl', f)))

hiddenimports += [
    'tkinter',
    'tkinter.ttk',
    'tkinter.messagebox',
    'tkinter.filedialog',
    'tkinter.font',
    'openpyxl.cell._writer',
    'openpyxl.styles.stylesheet',
    'fpdf',
    'fpdf.fonts',
    'fpdf.data',
    'darkdetect',
    'darkdetect._windows_detect',
    'PIL',
    'PIL.Image',
    'PIL.ImageTk',
]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'scipy', 'numpy', 'pandas', 'IPython', 'notebook', 'pytest',
        'cryptography', 'cffi', 'lxml', 'sqlite3', 'unittest'
    ],
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
    name='Stock_Think_Power',
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
    icon='app_icon.ico' if os.path.exists('app_icon.ico') else None,
)
