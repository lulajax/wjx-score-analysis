# -*- mode: python ; coding: utf-8 -*-
import os

block_cipher = None
pkg_dir = os.path.join('src', 'wjx_score')

a = Analysis(
    ['launcher.py'],
    pathex=['src'],
    binaries=[],
    datas=[
        (os.path.join(pkg_dir, 'static'), os.path.join('wjx_score', 'static')),
        (os.path.join(pkg_dir, 'template.html'), 'wjx_score'),
    ],
    hiddenimports=['wjx_score'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'unittest', 'test', 'xmlrpc', 'pydoc'],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='zhanpeng-toolbox',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    icon=None,
)
