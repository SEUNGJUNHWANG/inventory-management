# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 빌드 설정 파일
# 사용법: pyinstaller build.spec --noconfirm

import os
block_cipher = None

# 포함할 데이터 파일/폴더
datas = [
    ('core', 'core'),
    ('ui', 'ui'),
    ('utils', 'utils'),
]

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=datas,
    hiddenimports=[
        'gspread',
        'gspread.auth',
        'google.oauth2.service_account',
        'google.auth.transport.requests',
        'google.auth.crypt._python_rsa',
        'google.auth.crypt.es256',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'requests',
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
        'threading',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'numpy', 'pandas', 'PIL'],
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
    name='재고관리시스템',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # GUI 앱이므로 콘솔 숨김
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,              # 아이콘 경로 (예: 'assets/icon.ico')
)
