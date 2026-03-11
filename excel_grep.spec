# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['excel_grep.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('templates/filelist_template.csv', 'templates'),
        ('requirements.txt', '.'),
    ],
    hiddenimports=[
        'openpyxl',
        'openpyxl.cell._writer',
        'xlrd',
        'pandas',
        'tqdm',
        'colorama',
        'core.logger',
        'core.searcher',
        'core.file_handler',
        'core.exporter',
        'cli.parser',
        'cli.wizard',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='excel_grep',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='icon.ico',  # アイコンファイルがある場合はコメントを外す
)
