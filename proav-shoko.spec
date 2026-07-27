# -*- mode: python ; coding: utf-8 -*-

import sys
import os
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

# --- Configuration ---
APP_NAME = "ProAV Shoko"
APP_VERSION = "1.0.0"
MAIN_SCRIPT = "src/main.py"

# Data files to include
DATAS = [
    ("src/assets/report.css", "assets"),
]

# Hidden imports
HIDDEN_IMPORTS = [
    "usbmonitor",
    "screeninfo",
    "weasyprint",
    "weasyprint.css",
    "weasyprint.layout",
    "weasyprint.text",
    "cairo",
    "pangocairo",
    "gi",
    "gi.repository.Gtk",
    "gi.repository.Gdk",
    "gi.repository.GdkPixbuf",
    "gi.repository.Pango",
    "gi.repository.PangoCairo",
    "tkinter",
    "tkinter.ttk",
    "tkinter.filedialog",
    "tkinter.messagebox",
]

# Excludes
EXCLUDES = [
    "test",
    "tests",
    "pytest",
    "ruff",
    "mypy",
    "coverage",
]

# Platform-specific options
if sys.platform == "darwin":
    # macOS app bundle
    APP_BUNDLE = True
    ICON_FILE = None  # Add .icns file if available
    PLIST = {
        "CFBundleName": APP_NAME,
        "CFBundleDisplayName": APP_NAME,
        "CFBundleIdentifier": "com.klangche.proav-shoko",
        "CFBundleVersion": APP_VERSION,
        "CFBundleShortVersionString": APP_VERSION,
        "NSHighResolutionCapable": True,
        "LSMinimumSystemVersion": "10.15",
    }
    ICON = ICON_FILE
elif sys.platform == "win32":
    # Windows
    APP_BUNDLE = False
    ICON_FILE = None  # Add .ico file if available
    PLIST = {}
    ICON = ICON_FILE
else:
    # Linux
    APP_BUNDLE = False
    ICON_FILE = None
    PLIST = {}
    ICON = ICON_FILE


# Build Analysis
a = Analysis(
    [MAIN_SCRIPT],
    pathex=[str(Path(__file__).parent)],
    binaries=[],
    datas=DATAS,
    hiddenimports=HIDDEN_IMPORTS,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

# Filter out unnecessary modules
a.datas = [x for x in a.datas if "__pycache__" not in x[0]]

# PYZ
pyz = PYZ(a.pure, a.zipped_data, cipher=None)

# --- CLI Executable ---
cli_exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="proav-shoko",
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
)

# --- GUI Executable ---
gui_exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="ProAV Shoko",
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
    icon=ICON_FILE,
)

# --- macOS App Bundle ---
if sys.platform == "darwin" and APP_BUNDLE:
    app = BUNDLE(
        gui_exe,
        name=f"{APP_NAME}.app",
        icon=ICON_FILE,
        bundle_identifier="com.klangche.proav-shoko",
        version=APP_VERSION,
        info_plist=PLIST,
    )