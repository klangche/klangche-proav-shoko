"""
ProAV Shoko - Main entry point
Chooses between GUI and CLI based on arguments
"""

import sys
import argparse


def main():
    """Main function - starts the GUI or the CLI."""
    parser = argparse.ArgumentParser(
        description='ProAV Shoko - USB analysis tool for AV environments'
    )
    parser.add_argument(
        '--cli',
        action='store_true',
        help='Run in CLI mode (without GUI)'
    )
    parser.add_argument(
        '--csv-path',
        help='Path to custom hop_limits CSV file for USB analysis'
    )
    args = parser.parse_args()

    if args.cli:
        # CLI mode
        from src.main_cli import main as cli_main
        cli_main(args.csv_path)
    else:
        # GUI mode (default)
        try:
            from src.gui import ProAVShokoGUI
            import tkinter as tk
            root = tk.Tk()
            app = ProAVShokoGUI(root, args.csv_path)
            root.mainloop()
        except ImportError as e:
            print(f"[!] Could not start GUI: {e}")
            print("   Try --cli to run in the terminal.")
            sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n[!] Interrupted by user.")
        sys.exit(0)
    except Exception as e:
        print(f"\n[!] An error occurred: {e}")
        sys.exit(1)
