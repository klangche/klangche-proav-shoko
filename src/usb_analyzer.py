"""
USB-analysator - hanterar USB-träd, hops, tiers och stabilitetsbedömning
"""

import sys
import csv
import os
from collections import defaultdict
from typing import Dict, List, Any, Optional
from pathlib import Path

try:
    from usbmonitor import USBMonitor
    from usbmonitor.attributes import ID_MODEL, ID_VENDOR, DEVNAME, ID_MODEL_ID, ID_VENDOR_ID
except ImportError:
    print("❌ usbmonitor är inte installerat. Kör: pip install usbmonitor")
    sys.exit(1)


class USBAnalyzer:
    """Analyserar USB-enheter och bygger hierarkiskt träd."""

    # Standardgränser för hops per plattform
    DEFAULT_HOP_LIMITS = {
        'windows_x86': 4,
        'windows_arm': 4,
        'mac_intel': 4,
        'mac_apple_silicon': 3,
        'linux_x86': 4,
        'linux_arm': 4,
        'iphone_lightning': 2,
        'iphone_usbc': 3,
        'samsung_usbc': 4,
        'android_usbc': 4,
        'ipad_lightning': 2,
        'ipad_usbc': 3
    }

    def __init__(self, config_path: Optional[str] = None):
        """Initiera USB-analysatorn."""
        self.monitor = USBMonitor()
        self.devices = {}
        self.hop_limits = self._load_hop_limits(config_path)

    def _load_hop_limits(self, config_path: Optional[str] = None) -> Dict[str, int]:
        """Ladda hops-gränser från CSV eller använd standardvärden."""
        limits = self.DEFAULT_HOP_LIMITS.copy()

        if config_path and os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        platform = row.get('platform', '').strip()
                        max_hops = row.get('max_hops', '').strip()
                        if platform and max_hops.isdigit():
                            limits[platform] = int(max_hops)
                print(f"[+] Laddade hops-gränser från: {config_path}")
            except Exception as e:
                print(f"⚠️  Kunde inte ladda {config_path}: {e}")

        return limits

    def get_hop_limits_csv(self) -> str:
        """Generera CSV-innehåll för hops-gränser."""
        headers = ['platform', 'max_hops', 'description']
        rows = [
            ['windows_x86', str(self.hop_limits.get('windows_x86', 4)), 'Windows (Intel/AMD)'],
            ['windows_arm', str(self.hop_limits.get('windows_arm', 4)), 'Windows (ARM)'],
            ['mac_intel', str(self.hop_limits.get('mac_intel', 4)), 'Mac (Intel)'],
            ['mac_apple_silicon', str(self.hop_limits.get('mac_apple_silicon', 3)), 'Mac (Apple Silicon)'],
            ['linux_x86', str(self.hop_limits.get('linux_x86', 4)), 'Linux (Intel/AMD)'],
            ['linux_arm', str(self.hop_limits.get('linux_arm', 4)), 'Linux (ARM)'],
            ['iphone_lightning', str(self.hop_limits.get('iphone_lightning', 2)), 'iPhone (Lightning)'],
            ['iphone_usbc', str(self.hop_limits.get('iphone_usbc', 3)), 'iPhone (USB-C)'],
            ['samsung_usbc', str(self.hop_limits.get('samsung_usbc', 4)), 'Samsung (USB-C)'],
            ['android_usbc', str(self.hop_limits.get('android_usbc', 4)), 'Android (USB-C)'],
            ['ipad_lightning', str(self.hop_limits.get('ipad_lightning', 2)), 'iPad (Lightning)'],
            ['ipad_usbc', str(self.hop_limits.get('ipad_usbc', 3)), 'iPad (USB-C)']
        ]

        output = []
        output.append(','.join(headers))
        for row in rows:
            output.append(','.join(row))

        return '\n'.join(output)

    def save_hop_limits_csv(self, filepath: str) -> None:
        """Spara hops-gränser till CSV-fil."""
        content = self.get_hop_limits_csv()
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"[+] Sparade hops-gränser till: {filepath}")

    def build_tree(self) -> List[Dict[str, Any]]:
        """
        Bygger ett hierarkiskt träd av USB-enheter.

        Returns:
            Lista med trädstruktur för USB-enheter.
        """
        try:
            dev_dict = self.monitor.get_available_devices()
            self.devices = dev_dict

            tree = []
            for dev_name, attributes in dev_dict.items():
                devpath = attributes.get('devpath', '')

                node = {
                    'name': dev_name,
                    'devpath': devpath,
                    'attributes': attributes,
                    'children': [],
                    'is_hub': self._is_hub(attributes),
                    'model': attributes.get(ID_MODEL, 'Okänd modell'),
                    'vendor': attributes.get(ID_VENDOR, 'Okänd tillverkare'),
                    'product': attributes.get('ID_MODEL', 'Okänd produkt')
                }

                if not devpath or devpath.count('/') <= 1:
                    tree.append(node)
                else:
                    parent_path = '/'.join(devpath.split('/')[:-1])
                    parent = self._find_node_by_path(tree, parent_path)
                    if parent:
                        parent['children'].append(node)
                    else:
                        tree.append(node)

            return tree

        except Exception as e:
            print(f"⚠️  Kunde inte läsa USB-enheter: {e}")
            return []

    def _is_hub(self, attributes: Dict) -> bool:
        """Kontrollera om enhet är en hub."""
        model = attributes.get(ID_MODEL, '').lower()
        product = attributes.get('ID_MODEL', '').lower()
        return 'hub' in model or 'hub' in product

    def _find_node_by_path(self, tree: List[Dict], path: str) -> Optional[Dict]:
        """Hitta nod med specifik devpath i trädet."""
        for node in tree:
            if node['devpath'] == path:
                return node
            if node['children']:
                found = self._find_node_by_path(node['children'], path)
                if found:
                    return found
        return None

    def calculate_hops_and_tiers(self, tree: List[Dict]) -> Dict[str, Any]:
        """
        Beräknar hops och tiers från USB-trädet.

        Args:
            tree: USB-träd som lista med noder.

        Returns:
            Dictionary med hops, tiers och detaljer.
        """
        if not tree:
            return {
                'max_hops': 0,
                'max_tiers': 0,
                'devices_by_hops': {},
                'all_hops': []
            }

        depths = []
        devices_by_hops = defaultdict(list)

        def traverse(node, depth):
            hops = node['devpath'].count('/') if node.get('devpath') else depth
            devices_by_hops[hops].append(node['name'])
            depths.append(hops)

            for child in node['children']:
                traverse(child, depth + 1)

        for root in tree:
            traverse(root, 0)

        max_hops = max(depths) if depths else 0
        max_tiers = len(set(depths)) if depths else 0

        return {
            'max_hops': max_hops,
            'max_tiers': max_tiers,
            'devices_by_hops': dict(devices_by_hops),
            'all_hops': depths
        }

    def assess_stability(self, hops_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Bedömer stabilitet för ALLA plattformar baserat på hops.

        Args:
            hops_data: Data från calculate_hops_and_tiers.

        Returns:
            Dictionary med stabilitetsbedömning för alla plattformar.
        """
        max_hops = hops_data.get('max_hops', 0)

        # Definiera alla plattformar med deras hops-gränser
        platforms = [
            # x86/x64
            {'id': 'windows_x86', 'name': 'Windows', 'arch': 'x86/x64', 'max_hops': self.hop_limits.get('windows_x86', 4)},
            {'id': 'mac_intel', 'name': 'Mac Intel', 'arch': 'x86/x64', 'max_hops': self.hop_limits.get('mac_intel', 4)},
            {'id': 'linux_x86', 'name': 'Linux', 'arch': 'x86/x64', 'max_hops': self.hop_limits.get('linux_x86', 4)},
            # ARM
            {'id': 'windows_arm', 'name': 'Windows', 'arch': 'ARM', 'max_hops': self.hop_limits.get('windows_arm', 4)},
            {'id': 'mac_apple_silicon', 'name': 'Mac Apple Silicon', 'arch': 'ARM', 'max_hops': self.hop_limits.get('mac_apple_silicon', 3)},
            {'id': 'linux_arm', 'name': 'Linux', 'arch': 'ARM', 'max_hops': self.hop_limits.get('linux_arm', 4)},
            # Mobile
            {'id': 'iphone_lightning', 'name': 'iPhone Lightning', 'arch': 'Mobile', 'max_hops': self.hop_limits.get('iphone_lightning', 2)},
            {'id': 'iphone_usbc', 'name': 'iPhone USB-C', 'arch': 'Mobile', 'max_hops': self.hop_limits.get('iphone_usbc', 3)},
            {'id': 'samsung_usbc', 'name': 'Samsung USB-C', 'arch': 'Mobile', 'max_hops': self.hop_limits.get('samsung_usbc', 4)},
            {'id': 'android_usbc', 'name': 'Android USB-C', 'arch': 'Mobile', 'max_hops': self.hop_limits.get('android_usbc', 4)},
            {'id': 'ipad_lightning', 'name': 'iPad Lightning', 'arch': 'Mobile', 'max_hops': self.hop_limits.get('ipad_lightning', 2)},
            {'id': 'ipad_usbc', 'name': 'iPad USB-C', 'arch': 'Mobile', 'max_hops': self.hop_limits.get('ipad_usbc', 3)}
        ]

        # Bedöm varje plattform
        verdicts = []
        current_group = None

        for platform in platforms:
            max_allowed = platform['max_hops']
            is_stable = max_hops <= max_allowed

            if is_stable:
                status = 'STABIL'
                color = 'green'
                emoji = '🟢'
            elif max_hops <= max_allowed + 1:
                status = 'OSÄKER'
                color = 'orange'
                emoji = '🟠'
            else:
                status = 'INSTABIL'
                color = 'red'
                emoji = '🔴'

            verdict = {
                'id': platform['id'],
                'name': platform['name'],
                'arch': platform['arch'],
                'max_hops': max_allowed,
                'current_hops': max_hops,
                'status': status,
                'color': color,
                'emoji': emoji,
                'is_stable': is_stable,
                'warning': f"Max {max_allowed} hops" if not is_stable else None
            }
            verdicts.append(verdict)

        # Gruppera efter arkitektur för visning
        groups = {}
        for v in verdicts:
            if v['arch'] not in groups:
                groups[v['arch']] = []
            groups[v['arch']].append(v)

        return {
            'max_hops': max_hops,
            'verdicts': verdicts,
            'groups': groups,
            'overall_worst': max([v['status'] for v in verdicts], key=lambda x: {'STABIL': 0, 'OSÄKER': 1, 'INSTABIL': 2}[x])
        }

    def get_stability_summary(self, stability_data: Dict[str, Any]) -> str:
        """Skapa en textbaserad sammanfattning av stabilitetsbedömningen."""
        lines = []
        lines.append("[+] Stability Verdict:")
        lines.append("-" * 60)

        groups = stability_data.get('groups', {})
        for arch, verdicts in groups.items():
            lines.append(f"\n{arch}")
            for v in verdicts:
                lines.append(f"  {v['emoji']} {v['name']}")

        # Lägg till varningar
        warnings = [v for v in stability_data.get('verdicts', []) if not v['is_stable']]
        if warnings:
            lines.append("\n⚠️  VARNINGAR:")
            for w in warnings:
                lines.append(f"  • {w['name']}: {w['warning']} (nuvarande hops: {w['current_hops']})")

        return '\n'.join(lines)
