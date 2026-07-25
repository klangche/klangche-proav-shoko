#!/usr/bin/env python3
"""
ProAV Shoko - entry point for running from source and for packaged
(PyInstaller) builds.

Run this from the repository root:
    python run.py          # GUI mode
    python run.py --cli    # CLI mode
"""

import sys
from pathlib import Path

# Make sure the repo root is on sys.path so "src" resolves as a package,
# both when run directly (python run.py) and when frozen by PyInstaller.
sys.path.insert(0, str(Path(__file__).resolve().parent))

from src.main import main

if __name__ == "__main__":
    main()
