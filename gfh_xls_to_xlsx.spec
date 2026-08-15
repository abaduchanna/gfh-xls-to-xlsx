# -*- mode: python ; coding: utf-8 -*-
import datetime as _dt
_year = _dt.date.today().year

SPEC_DOC = f"""PyInstaller spec
Developed by Abad Umair Channa \u00a9 {_year}
Build command: pyinstaller gfh_xls_to_xlsx.spec
"""


block_cipher = None

a = Analysis(
    ['gfh_xls_to_xlsx.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('assets', 'assets'),
        ('gfh_icon.ico', '.'),
        ('GFH_Telecom_Logo.png', '.'),
        ('gfh_icon.png', '.'),
        ('gfh_wordmark.png', '.'),
        ('theme_manager.py', '.'),
        ('logo_handler.py', '.'),
        ('header_manager.py', '.'),
    ],
    hiddenimports=[
        'tkinter',
        '_tkinter',
        'tkinter._fix',
        'win32com',
        'win32com.client',
        'pythoncom',
        'pywintypes',
        'PIL',
        'theme_manager',
        'logo_handler',
        'header_manager',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludedimports=[
        'doctest',
        'pdb',
        'pyautogui',
        'selenium',
        'pandas',
        'numpy',
        'gspread',
        'oauth2client',
        'requests',
        'pyperclip',
        'torch',
        'torchvision',
        'torchaudio',
        'matplotlib',
        'matplotlib.pyplot',
        'numba',
        'llvmlite',
        'sympy',
        'tensorflow',
        'scipy',
        'sklearn',
        'scikit-learn',
        'speech_recognition',
        'SpeechRecognition',
        'imageio',
        'imageio_ffmpeg',
        'soundfile',
        'gi',
        'pygments',
        'fsspec',
        'tensorboard',
        'IPython',
        'ipython',
        'jupyter',
        'notebook',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='gfh_xls_to_xlsx',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='gfh_icon.ico',
)
