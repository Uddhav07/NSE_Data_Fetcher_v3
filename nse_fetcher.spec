# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec file for NSE Data Fetcher v3.

Build with:
    pyinstaller nse_fetcher.spec
"""

import os
import sys

block_cipher = None

# Paths
ROOT = os.path.abspath(".")
SCRIPTS = os.path.join(ROOT, "scripts")
ASSETS = os.path.join(ROOT, "assets")
CONFIG = os.path.join(ROOT, "config")

a = Analysis(
    ["app.py"],
    pathex=[ROOT],
    binaries=[],
    datas=[
        # Bundle default config alongside the exe
        (os.path.join(CONFIG, "config.json"), "config"),
        # Bundle the icon so the GUI can find it at runtime
        (os.path.join(ASSETS, "icon.ico"), "assets"),
    ],
    hiddenimports=[
        # yfinance pulls in several packages dynamically
        "yfinance",
        "yfinance.utils",
        "yfinance.ticker",
        "pandas",
        "pandas._libs.tslibs.timedeltas",
        "pandas._libs.tslibs.nattype",
        "pandas._libs.tslibs.np_datetime",
        "openpyxl",
        "openpyxl.styles",
        "openpyxl.formatting",
        "openpyxl.formatting.rule",
        "requests",
        "appdirs",
        "lxml",
        "html5lib",
        "frozendict",
        "peewee",
        "charset_normalizer",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Trim bloat — packages we definitely don't need
        "matplotlib",
        "scipy",
        "numpy.testing",
        "pytest",
        "IPython",
        "notebook",
        "jupyter",
        "PIL",  # only used at build time for icon generation
        "Pillow",
        "tkinter.test",
        "unittest",
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
    name="NSE Data Fetcher",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # Windowed app — no terminal popup
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join(ASSETS, "icon.ico"),
)

# One-file mode: no COLLECT needed — everything is packed into the single exe.
