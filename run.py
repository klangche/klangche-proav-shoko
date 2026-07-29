#!/usr/bin/env python3
"""
ProAV Shoko - Entry point for running from source and for packaged
(PyInstaller) builds.

Run this from the repository root:
    python run.py                                     # CLI mode (default)
    python run.py --gui                               # GUI mode
    python run.py --csv-path path/to/usb_data.csv     # CLI with custom limits
"""

import sys
from pathlib import Path

# Make sure the repo root is on sys.path so "src" resolves as a package,
# both when run directly (python run.py) and when frozen by PyInstaller.
sys.path.insert(0, str(Path(__file__).resolve().parent))

from src.main import main

if __name__ == "__main__":
    main()