# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all, collect_data_files, collect_dynamic_libs
import os

# ==================================================================================
# DATA FILES: Vosk model + any bundled assets
# ==================================================================================
datas = [('vosk-model-small-en-us-0.15', 'vosk-model-small-en-us-0.15')]
binaries = []

# ==================================================================================
# HIDDEN IMPORTS: Every submodule PyInstaller can't auto-detect
# ==================================================================================
hiddenimports = [
    # --- Python stdlib edge cases ---
    'email', 'email.mime', 'email.message', 'email.parser',
    'urllib', 'urllib.request',
    'xml.sax.saxutils',
    'ctypes', 'ctypes.wintypes',
    'multiprocessing', 'multiprocessing.freeze_support',
    '_cffi_backend',

    # --- pywin32 / COM automation ---
    'win32com', 'win32com.client', 'win32com.client.gencache',
    'win32com.shell', 'win32com.shell.shell',
    'win32api', 'win32gui', 'win32con', 'win32print',
    'pythoncom', 'pywintypes', 'winerror',

    # --- Core app dependencies ---
    'numpy', 'numpy.core', 'numpy.core._methods', 'numpy.lib', 'numpy.lib.format',
    'PIL', 'PIL.Image', 'PIL.ImageDraw', 'PIL.ImageFont', 'PIL.ImageTk',
    'sounddevice', '_sounddevice_data',
    'vosk', 'vosk.vosk_cffi',
    'speech_recognition',
    'screeninfo', 'screeninfo.enumerators', 'screeninfo.enumerators.windows',
    'customtkinter',
    'pptx', 'pptx.opc', 'pptx.oxml',

    # --- tkinter (for filedialog / messagebox) ---
    'tkinter', 'tkinter.filedialog', 'tkinter.messagebox',
    '_tkinter',

    # --- colorama (optional, used by live_speech_recognizer) ---
    'colorama',
]

# ==================================================================================
# COLLECT ALL: Let PyInstaller hooks grab data files, binaries, and submodules
# ==================================================================================
_packages_to_collect = [
    'customtkinter',    # theme JSON files + assets
    'vosk',             # native .dll/.so files
    'pptx',             # XML templates + schemas
    'speech_recognition',
    'sounddevice',      # _sounddevice_data portaudio DLL
    'screeninfo',
    'PIL',              # Pillow plugins
    'numpy',
]

for package in _packages_to_collect:
    try:
        pkg_datas, pkg_binaries, pkg_hiddenimports = collect_all(package)
        datas.extend(pkg_datas)
        binaries.extend(pkg_binaries)
        hiddenimports.extend(pkg_hiddenimports)
    except Exception as e:
        print(f"Warning collecting hook for {package}: {e}")

# Collect pywin32 DLLs explicitly (pythoncom*.dll, pywintypes*.dll)
try:
    pkg_datas, pkg_binaries, pkg_hiddenimports = collect_all('win32com')
    datas.extend(pkg_datas)
    binaries.extend(pkg_binaries)
    hiddenimports.extend(pkg_hiddenimports)
except Exception as e:
    print(f"Warning collecting win32com: {e}")

try:
    pkg_datas, pkg_binaries, pkg_hiddenimports = collect_all('pythoncom')
    datas.extend(pkg_datas)
    binaries.extend(pkg_binaries)
    hiddenimports.extend(pkg_hiddenimports)
except Exception as e:
    print(f"Warning collecting pythoncom: {e}")

try:
    pkg_datas, pkg_binaries, pkg_hiddenimports = collect_all('pywintypes')
    datas.extend(pkg_datas)
    binaries.extend(pkg_binaries)
    hiddenimports.extend(pkg_hiddenimports)
except Exception as e:
    print(f"Warning collecting pywintypes: {e}")

# Collect tkinter data (TCL/TK runtime)
try:
    pkg_datas, pkg_binaries, pkg_hiddenimports = collect_all('tkinter')
    datas.extend(pkg_datas)
    binaries.extend(pkg_binaries)
    hiddenimports.extend(pkg_hiddenimports)
except Exception as e:
    print(f"Warning collecting tkinter: {e}")

# ==================================================================================
# ANALYSIS
# ==================================================================================
a = Analysis(
    ['presentation_app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'PyQt6', 'PyQt5', 'PySide6', 'PySide2', 'scipy', 'pandas'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

# ==================================================================================
# EXE: Single-file output
# ==================================================================================
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='GenSlide',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,        # Clean windowed GUI application (no external terminal window)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
