# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['/home/kdy987/work/pst-utils/src/main.py'],
    pathex=['/home/kdy987/work/pst-utils/src'],
    binaries=[('/home/kdy987/work/pst-utils/.venv/lib/python3.12/site-packages/pypff.cpython-312-x86_64-linux-gnu.so', '.')],
    datas=[],
    hiddenimports=['config'],
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
    name='pst-extract',
    debug=False,
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
