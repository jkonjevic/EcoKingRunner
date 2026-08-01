# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

# ecoking/, ecoking_daily.py and web/ ship as loose files next to the exe
# (via `datas` below) instead of being baked into the frozen archive, so
# ecoking.selfupdate can refresh them from GitHub between runs. `excludes`
# stops PyInstaller's static analysis from ALSO bundling them as importable
# bytecode -- if it did, the frozen copy would shadow the loose one and
# self-update would silently do nothing.
datas = [
    ('.env', '.'),
    ('stations.json', '.'),
    ('ECO KING BLANKO TABLICA.xlsx', '.'),
    ('ecoking', 'ecoking'),
    ('web', 'web'),
    ('ecoking_daily.py', '.'),
    ('requirements.txt', '.'),
]
binaries = []
# Excluding ecoking/* from analysis (see excludes below) means PyInstaller
# never scans their source, so it can't auto-detect what THEY import either
# -- every stdlib module those files reach for has to be listed here by hand.
hiddenimports = ['http', 'http.client', 'http.cookies', 'http.server', 'json', 'mimetypes', 'secrets', 'webbrowser']
excludes = [
    'ecoking',
    'ecoking.webapp',
    'ecoking.stations',
    'ecoking.logtext',
    'ecoking.selfupdate',
    'ecoking.check',
    'ecoking_daily',
]
for package in ('playwright', 'openpyxl', 'dotenv'):
    tmp_ret = collect_all(package)
    datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['ecoking_web_launcher.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='EcoKingWebRunner',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    # Unlike the Tkinter build, this launcher has no window of its own to
    # show startup problems in -- a visible console is the only diagnostic
    # surface a non-technical user has if something goes wrong.
    console=True,
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
    name='EcoKingWebRunner',
)
