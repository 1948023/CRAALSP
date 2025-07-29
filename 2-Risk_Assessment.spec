# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['2-Risk_Assessment.py'],
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
        'PIL.Image',
        'PIL.ImageTk',
        'docx',
        'json',
        'csv'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tensorflow', 'tensorflow-cpu', 'tensorflow-gpu', 'tf', 'keras', 'torch', 'torchvision', 
        'sklearn', 'scipy', 'cv2', 'opencv-python', 'matplotlib', 'networkx', 'plotly', 
        'seaborn', 'statsmodels', 'sympy', 'IPython', 'jupyter', 'notebook', 'pytest',
        'setuptools', 'wheel', 'pip', 'conda', 'anaconda', 'spyder', 'idle', 'pydoc',
        'numpy.distutils', 'numpy.f2py', 'numpy.testing', 'pandas.plotting', 
        'pandas.tests', 'PIL.ImageChops', 'PIL.ImageCms', 'PIL.ImageColor', 
        'PIL.ImageEnhance', 'PIL.ImageFile', 'PIL.ImageFilter', 'PIL.ImageFont', 
        'PIL.ImageGrab', 'PIL.ImageMath', 'PIL.ImageMode', 'PIL.ImageOps', 
        'PIL.ImagePalette', 'PIL.ImagePath', 'PIL.ImageQt', 'PIL.ImageSequence', 
        'PIL.ImageShow', 'PIL.ImageStat', 'PIL.ImageTransform', 'PIL.ImageWin',
        'pandas.io', 'pandas.util', 'pandas.compat', 'numpy.core._multiarray_tests',
        'numpy.random._pickle', 'numpy.random._sfc64', 'numpy.random._pcg64',
        'numpy.random._philox', 'numpy.random._mt19937'
    ],
    noarchive=False,
    optimize=2,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='2-Risk_Assessment',
    debug=False,
    bootloader_ignore_signals=False,
    strip=True,
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
