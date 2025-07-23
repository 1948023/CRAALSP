# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['_Main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('Legacy.csv', '.'), 
        ('export_import_functions.py', '.'),
        ('logo.ico', '.'),
        ('attack_graph_threat_relations.csv', '.')
    ],
    hiddenimports=[
        'tkinter',
        'tkinter.ttk',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        'PIL.ImageDraw',
        'subprocess',
        'os',
        'sys'
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
    name='_Main',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Temporarily enable console for debugging
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='logo.ico',
)
