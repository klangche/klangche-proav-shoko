#!/usr/bin/env python3
"""Generate a macOS .icns icon from src/resources/shoko-icon.png.

Requires macOS (sips + iconutil). Used by CI when building the .app bundle.
"""

import shutil
import subprocess
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
PNG_PATH = REPO_ROOT / "src" / "resources" / "shoko-icon.png"
ICONSET_DIR = REPO_ROOT / "src" / "resources" / "shoko-icon.iconset"
ICNS_PATH = REPO_ROOT / "src" / "resources" / "shoko-icon.icns"

if not PNG_PATH.exists():
    print(f"Source icon not found: {PNG_PATH}")
    raise SystemExit(1)

if shutil.which("sips") is None or shutil.which("iconutil") is None:
    print("sips/iconutil not available, skipping .icns generation")
    raise SystemExit(0)

if ICONSET_DIR.exists():
    shutil.rmtree(ICONSET_DIR)
ICONSET_DIR.mkdir(parents=True)

# Standard iconset sizes: 16, 32, 128, 256, 512 and their @2x variants
sizes = [16, 32, 128, 256, 512]
for size in sizes:
    for scale in (1, 2):
        px = size * scale
        suffix = "" if scale == 1 else "@2x"
        out = ICONSET_DIR / f"icon_{size}x{size}{suffix}.png"
        subprocess.run(
            ["sips", "-z", str(px), str(px), str(PNG_PATH), "--out", str(out)],
            check=True,
        )

subprocess.run(
    ["iconutil", "-c", "icns", str(ICONSET_DIR), "-o", str(ICNS_PATH)],
    check=True,
)
shutil.rmtree(ICONSET_DIR)
print(f"Generated: {ICNS_PATH} ({ICNS_PATH.stat().st_size} bytes)")
