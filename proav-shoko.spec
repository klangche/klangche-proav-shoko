# -*- mode: python ; coding: utf-8 -*-

import sys
import os

# SPEC is provided by PyInstaller and contains the spec file path
BASE_DIR = os.path.dirname(os.path.abspath(SPEC))

# Add src to path and load version
sys.path.insert(0, os.path.join(BASE_DIR, "src"))
from version import get_exe_name, get_version
EXE_NAME = get_exe_name()
APP_VERSION = get_version()

# --- Configuration ---
APP_NAME = "ProAV Shoko"
MAIN_SCRIPT = "src/main.py"

# Data files to include (baked into the exe)
DATAS = [
    ("src/assets/report.css", "src/assets"),
    ("src/assets/usb_data.csv", "src/assets"),
    ("src/resources/shoko-icon.ico", "src/resources"),
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


# Collect weasyprint data files and native libs for PDF export
from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs
weasyprint_datas = collect_data_files('weasyprint')
weasyprint_bins = collect_dynamic_libs('weasyprint')

# Build Analysis
a = Analysis(
    [MAIN_SCRIPT],
    pathex=[BASE_DIR],
    binaries=weasyprint_bins,
    datas=DATAS + weasyprint_datas,
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
    name=EXE_NAME,
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
