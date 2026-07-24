#!/usr/bin/env python3
"""
ProAV Shōko - Huvudprogram
Väljer mellan GUI och CLI baserat på argument
"""

import sys
import argparse


def main():
    """Huvudfunktion - startar GUI eller CLI."""
    parser = argparse.ArgumentParser(
        description='ProAV Shōko - USB-analysverktyg för AV-miljöer'
    )
    parser.add_argument(
        '--cli',
        action='store_true',
        help='Kör i CLI-läge (utan GUI)'
    )
    args = parser.parse_args()

    if args.cli:
        # CLI-läge
        from src.main_cli import main as cli_main
        cli_main()
    else:
        # GUI-läge (standard)
        try:
            from src.gui import main as gui_main
            gui_main()
        except ImportError as e:
            print(f"❌ Kunde inte starta GUI: {e}")
            print("   Försök med --cli för att köra i terminalen.")
            sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️  Avbruten av användaren.")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ Ett fel uppstod: {e}")
        sys.exit(1)
