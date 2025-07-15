# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['3-attack_graph_analyzer.py'],
    pathex=[],
    binaries=[],
    datas=[('attack_graph_threat_relations.csv', '.'), ('Asset.json', '.'), ('Control.csv', '.'), ('Legacy.csv', '.'), ('Threat.csv', '.')],
    hiddenimports=[],
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
    name='3-attack_graph_analyzer_fixed',
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
