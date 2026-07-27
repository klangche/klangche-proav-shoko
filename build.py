#!/usr/bin/env python3
"""
Cross-platform build script for ProAV Shoko
Builds CLI and GUI executables for current platform.
"""

import sys
import subprocess
import shutil
import platform
from pathlib import Path


def run(cmd, cwd=None, check=True):
    """Run command and print output."""
    print(f"\n>>> {cmd}")
    result = subprocess.run(cmd, shell=True, cwd=cwd, capture_output=False)
    if check and result.returncode != 0:
        print(f"❌ Command failed: {cmd}")
        sys.exit(result.returncode)
    return result


def clean():
    """Clean build artifacts."""
    dirs = ["build", "dist", "__pycache__", "src/__pycache__"]
    for d in dirs:
        p = Path(d)
        if p.exists():
            shutil.rmtree(p)
            print(f"Cleaned: {d}")
    # Remove .spec cache
    for spec in Path(".").glob("*.spec"):
        if spec.name != "proav-shoko.spec":
            spec.unlink()


def install_deps():
    """Install build dependencies."""
    run("pip install -U pip pyinstaller")
    run("pip install -e .")


def build():
    """Build using PyInstaller."""
    run("pyinstaller proav-shoko.spec --clean")


def copy_assets():
    """Copy assets to dist for distribution."""
    system = platform.system().lower()
    arch = platform.machine().lower()

    dist = Path("dist")
    if not dist.exists():
        return

    # Determine platform suffix
    if sys.platform == "darwin":
        if "arm" in platform.machine().lower():
            suffix = "macos-arm64"
        else:
            suffix = "macos-x64"
    elif sys.platform == "win32":
        suffix = "windows-x64"
    else:
        suffix = f"linux-{platform.machine().lower()}"

    # Create distribution folder
    dist_name = f"proav-shoko-{APP_VERSION}-{suffix}"
    dist_dir = Path(dist_name)
    if dist_dir.exists():
        shutil.rmtree(dist_dir)
    dist_dir.mkdir()

    # Copy executables
    for item in Path("dist").iterdir():
        if item.is_file() or item.is_dir():
            shutil.copytree(item, dist_dir / item.name) if item.is_dir() else shutil.copy2(item, dist_dir)

    # Create zip
    shutil.make_archive(dist_name, 'zip', dist_name)
    print(f"\n✅ Distribution created: {dist_name}.zip")


# --- Main ---
APP_VERSION = "1.0.0"

if __name__ == "__main__":
    import argparse
    from pathlib import Path

    parser = argparse.ArgumentParser(description="Build ProAV Shoko")
    parser.add_argument("--clean", action="store_true", help="Clean build artifacts")
    parser.add_argument("--install", action="store_true", help="Install dependencies")
    parser.add_argument("--build", action="store_true", help="Build executable")
    parser.add_argument("--all", action="store_true", help="Clean, install, build")
    args = parser.parse_args()

    if args.clean:
        clean()

    if args.install or args.all:
        install_deps()

    if args.build or args.all:
        build()

    if args.all:
        copy_assets()