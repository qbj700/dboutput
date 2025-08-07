# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('jre', 'jre'), ('ojdbc8.jar', '.'), ('dboutput.ico', '.')],
    hiddenimports=['pymysql', 'psycopg2', 'jpype1', 'cryptography.hazmat.primitives.kdf.pbkdf2', 'cryptography.hazmat.primitives.hashes', 'cryptography.hazmat.primitives.ciphers', 'cryptography.hazmat.backends.openssl'],
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
    name='DB산출물생성기',
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
    icon='dboutput.ico',
)
