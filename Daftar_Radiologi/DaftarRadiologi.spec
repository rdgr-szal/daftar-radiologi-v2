# -*- mode: python ; coding: utf-8 -*-
import sys
import os

block_cipher = None

# Determine platform-specific icon
if sys.platform == 'darwin':
    app_icon = 'icon/daftarradiologi.icns' if os.path.exists('icon/daftarradiologi.icns') else 'icon/daftarradiologi.ico'
else:
    app_icon = 'icon/daftarradiologi.ico'

hidden_imports_list = [
    'pywebview',
    'webview',
    'openpyxl',
    'flask',
    'sqlite3',
    'core.config',
    'core.excel_engine',
    'core.phris_engine',
    'core.export_engine',
    'core.backup_engine',
    'core.db_engine',
    'routes.registration',
    'routes.patients',
    'routes.phris',
    'routes.dashboard',
    'routes.export',
    'routes.settings'
]

if sys.platform == 'win32':
    hidden_imports_list.extend([
        'clr_loader',
        'pythonnet',
        'clr',
        'PySide6',
        'PySide6.QtCore',
        'PySide6.QtWidgets',
        'PySide6.QtWebEngineWidgets',
        'PySide6.QtWebEngineCore',
    ])

a = Analysis(
    ['DaftarRadiologi.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('templates', 'templates'),
        ('static', 'static'),
        ('template.xlsx', '.'),
        ('core/smrp_taxonomy.json', 'core'),
        ('icon', 'icon'),
        ('MicrosoftEdgeWebview2Setup.exe', '.') if os.path.exists('MicrosoftEdgeWebview2Setup.exe') else ('icon', 'icon')
    ],
    hiddenimports=hidden_imports_list,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='DaftarRadiologi',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon=app_icon,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='DaftarRadiologi',
)

if sys.platform == 'darwin':
    app = BUNDLE(
        coll,
        name='DaftarRadiologi.app',
        icon='icon/daftarradiologi.icns',
        bundle_identifier='my.gov.moh.daftarradiologi',
        info_plist={
            'CFBundleName': 'Daftar Radiologi',
            'CFBundleDisplayName': 'Daftar Radiologi',
            'CFBundleIdentifier': 'my.gov.moh.daftarradiologi',
            'CFBundleVersion': '2.0.0',
            'CFBundleShortVersionString': '2.0.0',
            'NSHighResolutionCapable': 'True',
            'LSMinimumSystemVersion': '10.13.0',
        }
    )
