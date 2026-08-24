# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['presentation_app.py'],
    pathex=[],
    binaries=[],
    datas=[('vosk-model-small-en-us-0.15', 'vosk-model-small-en-us-0.15')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'PyQt6', 'PyQt5', 'PySide6', 'PySide2', 'scipy', 'pandas', 'tkinter.test', 'unittest'],
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
    name='GenSlide',
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
)
