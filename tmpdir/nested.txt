# -*- coding: utf-8 -*-
# PyInstaller 打包配置（单文件 exe）
# 用法（在 Windows 的 CMD 中，进入本目录后执行）：pyinstaller build.spec
# 生成的 exe 在 dist\ 目录下，名为「审单核对工具.exe」
# 说明：datas 中的 ('templates','templates') 把模板目录打进 exe；
#       ('products.json','.') 把初始货品数据打进 exe（首次运行会自动复制到用户 AppData）。

block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('templates', 'templates'),
        ('products.json', '.'),
    ],
    hiddenimports=['verifier', 'jinja2', 'markupsafe', 'openpyxl'],
    hookspath=[],
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
    name='审单核对工具',
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
)
