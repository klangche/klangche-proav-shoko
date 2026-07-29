"""
ProAV Shoko - Main entry point
CLI by default when running from source (python run.py).
GUI by default when running as packaged app (exe/.app).
"""

import sys
import argparse
import subprocess
import importlib


def _check_module(module_name, pip_name=None):
    """Check if a Python module is available. Returns True if found."""
    try:
        importlib.import_module(module_name)
        return True
    except ImportError:
        return False
    except Exception:
        return False


def _check_system_deps():
    """Check and offer to install missing system-level dependencies."""
    missing = []
    if sys.platform == "linux":
        import shutil
        if not shutil.which("pango-view"):
            missing.append("libpango-1.0-0")
        if not shutil.which("fc-list"):
            missing.append("fontconfig")
    elif sys.platform == "darwin":
        import shutil
        if not shutil.which("pango-view"):
            missing.append("pango")
    return missing


def _offer_install(pip_packages, system_packages=None):
    """Ask user before installing missing packages."""
    print("\n[!] Missing required dependencies detected:")
    if pip_packages:
        print(f"    Python packages: {', '.join(pip_packages)}")
    if system_packages:
        print(f"    System packages: {', '.join(system_packages)}")
    ans = input("\n    Install now? [Y/n]: ").strip().lower()
    if ans not in ("", "y", "yes"):
        print("    Skipping. Some features may not work.")
        return

    for pkg in pip_packages:
        print(f"    Installing {pkg}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])
            print(f"    {pkg} installed.")
        except subprocess.CalledProcessError:
            print(f"    Failed to install {pkg}. Please install manually: pip install {pkg}")

    if system_packages and sys.platform == "linux":
        print(f"    Installing system packages: {' '.join(system_packages)}...")
        try:
            subprocess.check_call(["sudo", "apt-get", "install", "-y"] + system_packages)
        except (subprocess.CalledProcessError, FileNotFoundError):
            print("    Could not install system packages. Install manually:")
            print(f"    sudo apt-get install {' '.join(system_packages)}")
    elif system_packages and sys.platform == "darwin":
        print(f"    Installing system packages: {' '.join(system_packages)}...")
        try:
            subprocess.check_call(["brew", "install"] + system_packages)
        except (subprocess.CalledProcessError, FileNotFoundError):
            print("    Could not install system packages. Install manually:")
            print(f"    brew install {' '.join(system_packages)}")


def _set_icon_later(root, ico_path):
    """Set window icon after the window is mapped (ensures taskbar icon)."""
    if ico_path and ico_path.exists():
        try:
            root.iconbitmap(str(ico_path))
        except Exception:
            pass


def main():
    """Main function - starts the GUI or the CLI."""
    is_frozen = getattr(sys, 'frozen', False)

    parser = argparse.ArgumentParser(
        description='ProAV Shoko - USB analysis tool for AV environments'
    )
    parser.add_argument(
        '--cli',
        action='store_true',
        default=False,
        help='Run in CLI mode'
    )
    parser.add_argument(
        '--gui',
        action='store_true',
        default=False,
        help='Run in GUI mode'
    )
    parser.add_argument(
        '--csv-path',
        help='Path to custom USB data file (CSV format, e.g. usb_data.csv)'
    )
    args = parser.parse_args()

    if args.gui and args.cli:
        print("[!] Cannot use --gui and --cli together. Defaulting to GUI.")
        args.cli = False
    want_gui = args.gui or (not args.cli and is_frozen)

    if not is_frozen:
        missing_pip = []
        if want_gui and not _check_module("customtkinter"):
            missing_pip.append("customtkinter")
        if want_gui and not _check_module("tkinter"):
            print("[!] tkinter not found. GUI mode requires Python tk support.")
            print("    Linux: sudo apt-get install python3-tk")
            print("    macOS: brew install python-tk")
        if not _check_module("usbmonitor"):
            missing_pip.append("usb-monitor")
        if not _check_module("screeninfo"):
            missing_pip.append("screeninfo")

        system_missing = _check_system_deps()
        if missing_pip or system_missing:
            _offer_install(missing_pip, system_missing)

    if want_gui:
        try:
            import customtkinter as ctk
            ctk.set_appearance_mode("system")
            ctk.set_default_color_theme("blue")
            from src.gui import ProAVShokoGUI
            root = ctk.CTk()
            from pathlib import Path
            ico = Path(__file__).parent / "resources" / "shoko-icon.ico"
            if ico.exists():
                try:
                    root.iconbitmap(str(ico))
                except Exception:
                    pass
            root.after(50, lambda: _set_icon_later(root, ico))
            app = ProAVShokoGUI(root, args.csv_path)
            root.mainloop()
        except ImportError as e:
            print(f"[!] Could not start GUI: {e}")
            sys.exit(1)
    else:
        from src.main_cli import main as cli_main
        cli_main(args.csv_path)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n[!] Interrupted by user.")
        sys.exit(0)
    except Exception as e:
        print(f"\n[!] An error occurred: {e}")
        sys.exit(1)
