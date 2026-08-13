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
        ('GFH_Telecom_Logo.png', '.'),
        ('gfh_app_icon_new.ico', '.'),
        ('gfh_icon_white.ico', '.'),
        ('theme_manager.py', '.'),
    ],
    hiddenimports=[
        'tkinter',
        'pyautogui',
        'openpyxl',
        'pyperclip',
        'requests',
        'theme_manager',
        'logo_handler',
        'PIL',
        'pandas',
        'gspread',
        'oauth2client',
        'pywin32',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludedimports=[
        'doctest',
        'pdb',
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
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='gfh_app_icon_new.ico',
)
