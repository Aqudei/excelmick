# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules


hiddenimports=['win32api', 'win32file', 'win32con', 'win32security', 'win32event']
hiddenimports+= collect_submodules('watchdog')

a = Analysis(
    ['checker.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['win32api', 'win32file', 'win32con', 'win32security', 'win32event'],
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
    name='checker',
    debug=False,x
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
