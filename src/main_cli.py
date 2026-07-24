#!/usr/bin/env python3
"""
ProAV Shoko - CLI mode (original PowerShell port)
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from src.usb_analyzer import USBAnalyzer
from src.display_analyzer import DisplayAnalyzer
from src.report_generator import ReportGenerator
from src.platform_utils import PlatformUtils


def main():
    """Main function for CLI mode."""
    print("\n" + "=" * 60)
    print("  KLANGCHE PROAV SHOKO - USB DETECTIVE")
    print("=" * 60 + "\n")

    # 1. Platform information
    platform_info = PlatformUtils.get_platform_info()
    print(f"[+] Platform: {platform_info['os']} {platform_info['version']}")
    print(f"[+] Architecture: {platform_info['architecture']}")
    if platform_info['is_apple_silicon']:
        print("[+] Apple Silicon detected!")
    print("-" * 60)

    # 2. USB analysis - load hop limits from CSV if it exists
    config_path = Path("hop_limits.csv")
    if config_path.exists():
        usb_analyzer = USBAnalyzer(str(config_path))
        print(f"[+] Loaded hop limits from: {config_path}")
    else:
        usb_analyzer = USBAnalyzer()
        usb_analyzer.save_hop_limits_csv("hop_limits.csv")
        print(f"[+] Created default hop_limits.csv")

    print("\n[+] Scanning USB devices...")
    usb_tree = usb_analyzer.build_tree()
    hops_data = usb_analyzer.calculate_hops_and_tiers(usb_tree)

    # 3. Stability assessment for all platforms
    stability = usb_analyzer.assess_stability(hops_data)

    # 4. Display information
    print("\n[+] Scanning displays...")
    display_analyzer = DisplayAnalyzer()
    displays = display_analyzer.get_display_info()

    # 5. Show USB tree
    print("\n[+] USB Tree Structure:")
    print("-" * 60)
    if usb_tree:
        _render_tree(usb_tree)
    else:
        print("  No USB devices found.")

    # 6. Hops analysis
    print("\n[+] Hops Analysis:")
    print("-" * 60)
    print(f"Maximum Hops: {hops_data['max_hops']}")
    print(f"Total Tiers: {hops_data['max_tiers']}")
    print(f"Number of Hubs: {len([d for d in usb_tree if d.get('is_hub', False)])}")

    # 7. Stability assessment - ALL platforms
    print(usb_analyzer.get_stability_summary(stability))

    # 8. Display information
    print("\n[+] Display Information:")
    print("-" * 60)
    if displays:
        for display in displays:
            primary = " (Primary)" if display.get('is_primary', False) else ""
            print(f"  {display['resolution']}  {display['name']}{primary}")
    else:
        print("  No displays found.")

    # 9. Generate reports
    print("\n[+] Generating reports...")
    report_gen = ReportGenerator()

    html_path = report_gen.generate_html_report(
        usb_tree,
        hops_data,
        stability,
        displays,
        platform_info,
        platform_notes=usb_analyzer.get_platform_notes()
    )
    print(f"  HTML: {html_path}")

    pdf_path = report_gen.generate_pdf_report(html_path)
    if pdf_path:
        print(f"  PDF:  {pdf_path}")

    # 10. Open reports
    print("\n[+] Opening reports...")
    report_gen.open_report(html_path)
    if pdf_path:
        report_gen.open_report(pdf_path)

    print("\n" + "=" * 60)
    print("  Done!")
    print("=" * 60)


def _render_tree(tree, level=0):
    """Recursively render the USB tree."""
    indent = "  " * level
    for node in tree:
        is_hub = node.get('is_hub', False)
        hub_tag = " [HUB]" if is_hub else ""
        hops = node['devpath'].count('/') if node.get('devpath') else 0

        print(f"{indent}{node.get('model', 'Unknown')}{hub_tag}  hops: {hops}")

        if node.get('children'):
            _render_tree(node['children'], level + 1)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n[!] Interrupted by user.")
        sys.exit(0)
    except Exception as e:
        print(f"\n[!] An error occurred: {e}")
        sys.exit(1)
