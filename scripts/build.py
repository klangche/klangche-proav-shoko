#!/usr/bin/env python3
"""
Cross-platform build script for ProAV Shoko.
Auto-increments prerelease number on each build.
"""

import sys
import subprocess
import shutil
import re
import platform
from pathlib import Path


def run(cmd, cwd=None, check=True):
    print(f"\n>>> {cmd}")
    result = subprocess.run(cmd, shell=True, cwd=cwd, capture_output=False)
    if check and result.returncode != 0:
        print(f"Command failed: {cmd}")
        sys.exit(result.returncode)
    return result


def read_version():
    """Read version info from src/version.py."""
    version_py = Path("src/version.py")
    ns = {}
    exec(version_py.read_text(), ns)
    return ns


def write_prerelease(value):
    """Update PRERELEASE in src/version.py."""
    version_py = Path("src/version.py")
    lines = version_py.read_text().splitlines()
    for i, line in enumerate(lines):
        if line.startswith("PRERELEASE = "):
            lines[i] = f"PRERELEASE = {value}"
            break
    version_py.write_text("\n".join(lines) + "\n")


def write_stage(stage):
    """Update STAGE in src/version.py (alpha/beta/rc/stable)."""
    version_py = Path("src/version.py")
    lines = version_py.read_text().splitlines()
    for i, line in enumerate(lines):
        if line.startswith("STAGE = "):
            lines[i] = f'STAGE = "{stage}"'
            break
    version_py.write_text("\n".join(lines) + "\n")


def clean():
    dirs = ["build", "dist", "__pycache__", "src/__pycache__"]
    for d in dirs:
        p = Path(d)
        if p.exists():
            shutil.rmtree(p)
            print(f"Cleaned: {d}")
    for spec in Path(".").glob("*.spec"):
        if spec.name != "proav-shoko.spec":
            spec.unlink()


def install_deps():
    run("pip install -U pip pyinstaller")
    run("pip install -e .")


def build():
    run("python -m PyInstaller proav-shoko.spec --clean")


def copy_assets(version_str):
    system = platform.system().lower()
    suffix = (
        "macos-arm64" if sys.platform == "darwin" and "arm" in platform.machine().lower() else
        "macos-x64" if sys.platform == "darwin" else
        "windows-x64" if sys.platform == "win32" else
        f"linux-{platform.machine().lower()}"
    )

    dist_name = f"proav-shoko-{version_str}-{suffix}"
    dist_dir = Path(dist_name)
    if dist_dir.exists():
        shutil.rmtree(dist_dir)
    dist_dir.mkdir()

    for item in Path("dist").iterdir():
        if item.is_file() or item.is_dir():
            shutil.copytree(item, dist_dir / item.name) if item.is_dir() else shutil.copy2(item, dist_dir)

    shutil.make_archive(dist_name, "zip", dist_name)
    print(f"\nDistribution created: {dist_name}.zip")


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Build ProAV Shoko")
    parser.add_argument("--clean", action="store_true", help="Clean build artifacts")
    parser.add_argument("--install", action="store_true", help="Install dependencies")
    parser.add_argument("--build", action="store_true", help="Build executable")
    parser.add_argument("--all", action="store_true", help="Clean, install, build")
    parser.add_argument("--stage", choices=["alpha", "beta", "rc", "stable"], help="Set release stage before building")
    parser.add_argument("--no-increment", action="store_true", help="Skip auto-increment of prerelease number")
    args = parser.parse_args()

    if args.stage:
        write_stage(args.stage)
        print(f"Stage set to: {args.stage}")

    if not args.no_increment and (args.build or args.all):
        ver = read_version()
        new_pr = ver["PRERELEASE"] + 1
        write_prerelease(new_pr)
        print(f"Prerelease incremented: {ver['PRERELEASE']} -> {new_pr}")

    if args.clean:
        clean()

    if args.install or args.all:
        install_deps()

    if args.build or args.all:
        build()

    if args.all:
        ver = read_version()
        copy_assets(ver["get_version"]())


if __name__ == "__main__":
    main()
