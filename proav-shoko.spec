# -*- mode: python ; coding: utf-8 -*-

import sys
import os

# SPEC is provided by PyInstaller and contains the spec file path
BASE_DIR = os.path.dirname(os.path.abspath(SPEC))

# Add src to path
sys.path.insert(0, os.path.join(BASE_DIR, "src"))

# --- Configuration ---
APP_NAME = "ProAV Shoko"
APP_VERSION = "1.0.0"
MAIN_SCRIPT = "src/main.py"

# Data files to include (baked into the exe)
DATAS = [
    ("src/assets/report.css", "assets"),
    ("src/assets/usb_data.csv", "assets"),
]

# Hidden imports
HIDDEN_IMPORTS = [
    "customtkinter",
    "usbmonitor",
    "screeninfo",
    "weasyprint",
    "weasyprint.css",
    "weasyprint.layout",
    "weasyprint.text",
    "tkinter",
    "tkinter.ttk",
    "tkinter.filedialog",
    "tkinter.messagebox",
    "wmi",
    "pythoncom",
    "win32com.client",
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

# Icon file for Windows
ICON_FILE = os.path.join(BASE_DIR, "src", "resources", "shoko-icon.ico")


# Build Analysis
a = Analysis(
    [MAIN_SCRIPT],
    pathex=[BASE_DIR],
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

# --- Single Executable (GUI + CLI via --cli flag) ---
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name=f"ProAV Shoko {APP_VERSION}",
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

# macOS App Bundle - paused (WIP)
# if sys.platform == "darwin":
#     app = BUNDLE(
#         gui_exe,
#         name=f"{APP_NAME}.app",
#         icon=ICON_FILE,
#         bundle_identifier="com.klangche.proav-shoko",
#         version=APP_VERSION,
#         info_plist={},
#     )
