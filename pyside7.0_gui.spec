# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['pyside7.0_gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('Pictures', 'Pictures'), 
        ('dark.qss', '.'), 
        ('girly.qss', '.'),
        ('Logo.ico', '.'),
        ('Logo.png', '.')
    ],
    hiddenimports=[
        'PySide6.QtSvg',
        'PySide6.QtGui',
        'PySide6.QtCore',
        'PySide6.QtWidgets',
        'docxtpl',
        'jinja2',
        'pandas',
        'openpyxl',
        'qrcode',
        'PIL',
        'numpy'
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
    name='DocxTpl_Automatisierung_v7.0',
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
    icon='Logo.ico',
)
