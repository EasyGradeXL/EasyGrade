# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

foo_a = Analysis(
    ['GradeBook.pyw'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
bar_a = Analysis(
    ['SetupEasyGrade.pyw'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
MERGE((foo_a, 'foo', 'foo'), (bar_a, 'bar', 'bar'))

foo_pyz = PYZ(foo_a.pure, foo_a.zipped_data, cipher=block_cipher)
foo_exe = EXE(
    foo_pyz,
    foo_a.scripts,
    [],
    exclude_binaries=True,
    name='GradeBook',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
foo_coll = COLLECT(
    foo_exe,
    foo_a.binaries,
    foo_a.zipfiles,
    foo_a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='GradeBook',
)

bar_pyz = PYZ(bar_a.pure, bar_a.zipped_data, cipher=block_cipher)
bar_exe = EXE(
    bar_pyz,
    bar_a.scripts,
    [],
    exclude_binaries=True,
    name='SetupEasyGrade',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
bar_coll = COLLECT(
    bar_exe,
    bar_a.binaries,
    bar_a.zipfiles,
    bar_a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SetupEasyGrade',
)
