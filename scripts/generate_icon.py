import struct
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
PNG_PATH = REPO_ROOT / "src" / "resources" / "shoko-icon.png"
ICO_PATH = REPO_ROOT / "src" / "resources" / "shoko-icon.ico"

if not PNG_PATH.exists():
    print(f"Source icon not found: {PNG_PATH}")
    exit(1)

png_data = PNG_PATH.read_bytes()

header = struct.pack("<HHH", 0, 1, 1)
entry = struct.pack("<BBBBHHII", 0, 0, 0, 0, 1, 32, len(png_data), 22)

ICO_PATH.write_bytes(header + entry + png_data)
print(f"Generated: {ICO_PATH} ({len(header) + len(entry) + len(png_data)} bytes)")
