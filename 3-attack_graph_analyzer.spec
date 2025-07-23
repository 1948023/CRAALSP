# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['3-attack_graph_analyzer.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('Control.csv', '.'), 
        ('Legacy.csv', '.'), 
        ('Threat.csv', '.'), 
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
        'networkx',
        'matplotlib',
        'matplotlib.pyplot',
        'matplotlib.backends.backend_tkagg',
        'pandas',
        'numpy',
        'csv',
        'json',
        'os',
        'sys'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tensorflow', 'tensorflow-cpu', 'tensorflow-gpu', 'tf', 'keras', 'torch', 'torchvision', 'sklearn', 'scipy', 'cv2', 'opencv-python'],
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
    name='3-attack_graph_analyzer',
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
    icon='logo.ico',
)