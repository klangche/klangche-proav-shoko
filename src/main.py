"""
ProAV Shoko - Main entry point
CLI by default when running from source (python run.py).
GUI by default when running as packaged app (exe/.app).
"""

import sys
import argparse


def main():
    """Main function - starts the GUI or the CLI."""
    is_frozen = getattr(sys, 'frozen', False)

    parser = argparse.ArgumentParser(
        description='ProAV Shoko - USB analysis tool for AV environments'
    )
    if is_frozen:
        parser.add_argument(
            '--cli',
            action='store_true',
            help='Run in CLI mode (default is GUI when packaged)'
        )
    else:
        parser.add_argument(
            '--gui',
            action='store_true',
            help='Run in GUI mode (default is CLI)'
        )
    parser.add_argument(
        '--csv-path',
        help='Path to custom USB data file (CSV format, e.g. usb_data.csv)'
    )
    args = parser.parse_args()

    want_gui = args.gui if not is_frozen else not args.cli

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
                root.iconbitmap(str(ico))
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
