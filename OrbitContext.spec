from pathlib import Path
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    [str(Path(SPECPATH) / 'app' / 'orbit_context.py')],
    pathex=[str(Path(SPECPATH) / 'app')],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
# Windows 10/11 provide API sets, UCRT and the system ICU shim.
# Exclude unrelated DLLs discovered through developer tool directories in PATH.
a.binaries = [x for x in a.binaries if not (
    x[0].lower().startswith(('api-ms-win-', 'ext-ms-win-')) or
    x[0].lower() in ('ucrtbase.dll', 'icuuc.dll', 'icudt78.dll'))]
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='OrbitContext',
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
    icon=[str(Path(SPECPATH) / 'app' / 'app.ico')],
)
