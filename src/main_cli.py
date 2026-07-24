#!/usr/bin/env python3
"""
ProAV Shōko - CLI-läge (original PowerShell-port)
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from src.usb_analyzer import USBAnalyzer
from src.display_analyzer import DisplayAnalyzer
from src.report_generator import ReportGenerator
from src.platform_utils import PlatformUtils


def main():
    """Huvudfunktion för CLI-läge."""
    print("\n" + "=" * 60)
    print("  🔍 KANGCHE PROAV SHOKO - USB DETECTIVE")
    print("=" * 60 + "\n")

    # 1. Plattformsinformation
    platform_info = PlatformUtils.get_platform_info()
    print(f"[+] Platform: {platform_info['os']} {platform_info['version']}")
    print(f"[+] Architecture: {platform_info['architecture']}")
    if platform_info['is_apple_silicon']:
        print("[+] Apple Silicon detected!")
    print("-" * 60)

    # 2. USB-analys - ladda hops-gränser från CSV om den finns
    config_path = Path("hop_limits.csv")
    if config_path.exists():
        usb_analyzer = USBAnalyzer(str(config_path))
        print(f"[+] Laddade hops-gränser från: {config_path}")
    else:
        usb_analyzer = USBAnalyzer()
        usb_analyzer.save_hop_limits_csv("hop_limits.csv")
        print(f"[+] Skapade standard hop_limits.csv")

    print("\n[+] Scanning USB devices...")
    usb_tree = usb_analyzer.build_tree()
    hops_data = usb_analyzer.calculate_hops_and_tiers(usb_tree)

    # 3. Stabilitetsbedömning för alla plattformar
    stability = usb_analyzer.assess_stability(hops_data)

    # 4. Skärminformation
    print("\n[+] Scanning displays...")
    display_analyzer = DisplayAnalyzer()
    displays = display_analyzer.get_display_info()

    # 5. Visa USB-träd
    print("\n[+] USB Tree Structure:")
    print("-" * 60)
    if usb_tree:
        _render_tree(usb_tree)
    else:
        print("  Inga USB-enheter hittades.")

    # 6. Hops-analys
    print("\n[+] Hops Analysis:")
    print("-" * 60)
    print(f"Maximum Hops: {hops_data['max_hops']}")
    print(f"Total Tiers: {hops_data['max_tiers']}")
    print(f"Number of Hubs: {len([d for d in usb_tree if d.get('is_hub', False)])}")

    # 7. Stabilitetsbedömning - ALLA plattformar
    print(usb_analyzer.get_stability_summary(stability))

    # 8. Skärminformation
    print("\n[+] Display Information:")
    print("-" * 60)
    if displays:
        for display in displays:
            primary = " (Primary)" if display.get('is_primary', False) else ""
            print(f"  🖥️  {display['resolution']}  {display['name']}{primary}")
    else:
        print("  Inga skärmar hittades.")

    # 9. Generera rapporter
    print("\n[+] Generating reports...")
    report_gen = ReportGenerator()

    html_path = report_gen.generate_html_report(
        usb_tree,
        hops_data,
        stability,
        displays,
        platform_info
    )
    print(f"  📄 HTML: {html_path}")

    pdf_path = report_gen.generate_pdf_report(html_path)
    if pdf_path:
        print(f"  📄 PDF:  {pdf_path}")

    # 10. Öppna rapporter
    print("\n[+] Opening reports...")
    report_gen.open_report(html_path)
    if pdf_path:
        report_gen.open_report(pdf_path)

    print("\n" + "=" * 60)
    print("  ✅ Klart!")
    print("=" * 60)


def _render_tree(tree, level=0):
    """Rekursivt rendera USB-träd."""
    indent = "  " * level
    for node in tree:
        is_hub = node.get('is_hub', False)
        icon = "📌" if is_hub else "🖥️"
        hub_tag = " [HUB]" if is_hub else ""
        hops = node['devpath'].count('/') if node.get('devpath') else 0

        print(f"{indent}{icon} {node.get('model', 'Okänd')}{hub_tag}  hops: {hops}")

        if node.get('children'):
            _render_tree(node['children'], level + 1)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️  Avbruten av användaren.")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ Ett fel uppstod: {e}")
        sys.exit(1)
