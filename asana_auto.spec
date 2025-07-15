# asana_auto.spec (Final Version 2)

a = Analysis(
    ['asana_auto_main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('config.json', '.'),
        ('template.xltx', '.'),
        ('success.mp3', '.'),
        ('error.mp3', '.')
    ],
    # CORRECTED: Removed 'win32com.gen_py' which can cause build errors.
    hiddenimports=['playsound'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Asana Automation',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Asana Automation'
)