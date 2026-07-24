#!/usr/bin/env python3
"""
ProAV Shōko - Huvudprogram för USB-analys
"""

import sys
import os
from pathlib import Path

# Lägg till src i sökvägen för att möjliggöra relativa importer
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.usb_analyzer import USBAnalyzer
from src.display_analyzer import DisplayAnalyzer
from src.report_generator import ReportGenerator
from src.platform_utils import PlatformUtils


def main():
    """Huvudfunktion för programmet."""
    print("\n" + "="*60)
    print("  🔍 Kangche ProAV Shōko - USB-analysverktyg")
    print("="*60 + "\n")

    # 1. Plattformsinformation
    platform_info = PlatformUtils.get_platform_info()
    print(f"🖥️  Plattform: {platform_info['os']} {platform_info['version']}")
    print(f"🧠  Arkitektur: {platform_info['architecture']}")
    if platform_info['is_apple_silicon']:
        print("⚠️  Apple Silicon (M-chip) detekterad!")
    print("-"*60)

    # 2. Analysera USB-enheter
    print("\n📡 Analyserar USB-enheter...")
    usb_analyzer = USBAnalyzer()
    usb_tree = usb_analyzer.build_tree()
    hops_data = usb_analyzer.calculate_hops_and_tiers(usb_tree)
    stability = usb_analyzer.assess_stability(hops_data, platform_info['is_apple_silicon'])

    # 3. Analysera skärmar
    print("\n🖥️  Analyserar anslutna skärmar...")
    display_analyzer = DisplayAnalyzer()
    displays = display_analyzer.get_display_info()

    # 4. Generera rapporter
    print("\n📄 Genererar rapporter...")
    report_gen = ReportGenerator()

    # HTML-rapport
    html_path = report_gen.generate_html_report(
        usb_tree,
        hops_data,
        stability,
        displays,
        platform_info
    )
    print(f"✅ HTML-rapport skapad: {html_path}")

    # PDF-rapport
    pdf_path = report_gen.generate_pdf_report(html_path)
    if pdf_path:
        print(f"✅ PDF-rapport skapad: {pdf_path}")

    # 5. Öppna rapporter
    print("\n📂 Öppnar rapporter...")
    report_gen.open_report(html_path)
    if pdf_path:
        report_gen.open_report(pdf_path)

    # 6. Sammanfattning i terminalen
    print("\n" + "="*60)
    print("📊 SAMMANFATTNING")
    print("="*60)
    print(f"📌 USB-enheter: {len(usb_tree)}")
    print(f"📌 Hops (max): {hops_data['max_hops']}")
    print(f"📌 Tiers (max): {hops_data['max_tiers']}")
    print(f"📌 Stabilitetsbedömning: {stability['label']} ({stability['color']})")
    if stability['warning']:
        print(f"⚠️  {stability['warning']}")
    print(f"🖥️  Anslutna skärmar: {len(displays)}")
    print("="*60 + "\n")

    print("✅ Klart! Rapporterna har öppnats i din webbläsare/PDF-visare.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️  Avbruten av användaren.")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ Ett fel uppstod: {e}")
        sys.exit(1)
