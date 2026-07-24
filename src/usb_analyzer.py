"""
USB analyzer - handles the USB tree, hops, tiers and stability assessment
"""

import sys
import csv
from collections import defaultdict
from typing import Dict, List, Any, Optional
from pathlib import Path

try:
    from usbmonitor import USBMonitor
    from usbmonitor.attributes import ID_MODEL, ID_VENDOR, DEVNAME, ID_MODEL_ID, ID_VENDOR_ID
except ImportError:
    print("[!] usbmonitor is not installed. Run: pip install usbmonitor")
    sys.exit(1)


class USBAnalyzer:
    """Analyzes USB devices and builds a hierarchical tree."""

    def __init__(self, config_path: Optional[str] = None):
        """
        Initialize the USB analyzer.

        Args:
            config_path: Optional path to a hop_limits.csv file to load
                instead of the one bundled in src/assets.
        """
        self.monitor = USBMonitor()
        self.devices = {}

        # Load hop limits either from the given config path or the
        # bundled CSV file in assets.
        if config_path:
            self.config_path = Path(config_path)
        else:
            self.config_path = Path(__file__).parent / "assets" / "hop_limits.csv"

        self.hop_limits = self._load_hop_limits()
        self.platform_notes = self._load_platform_notes()

    def _load_hop_limits(self) -> Dict[str, int]:
        """Load hop limits from the CSV file."""
        limits = {}

        if self.config_path and self.config_path.exists():
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        platform = row.get('platform', '').strip()
                        max_hops = row.get('max_hops', '').strip()
                        if platform and max_hops.isdigit():
                            limits[platform] = int(max_hops)
                print(f"[+] Loaded hop limits from: {self.config_path}")
            except Exception as e:
                print(f"[!] Could not load {self.config_path}: {e}")
                limits = self._get_fallback_limits()
        else:
            print(f"[!] {self.config_path} not found, using fallback values")
            limits = self._get_fallback_limits()

        return limits

    def _get_fallback_limits(self) -> Dict[str, int]:
        """Fallback values if the CSV is missing."""
        return {
            'windows_x86': 4,
            'windows_arm': 4,
            'mac_intel': 7,
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

    def _load_platform_notes(self) -> List[Dict[str, str]]:
        """Load notes from the CSV file."""
        notes = []
        if self.config_path and self.config_path.exists():
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        platform = row.get('platform', '').strip()
                        description = row.get('description', '').strip()
                        note = row.get('notes', '').strip()
                        if platform and note:
                            notes.append({
                                'platform': platform,
                                'description': description,
                                'note': note
                            })
            except Exception as e:
                print(f"[!] Could not read notes: {e}")
        return notes

    def get_platform_notes(self) -> List[Dict[str, str]]:
        """Get notes for use in reports."""
        return self.platform_notes

    def save_hop_limits_csv(self, path: str) -> None:
        """
        Save the current hop limits (and any known notes) to a CSV file.

        Args:
            path: Destination path for the CSV file.
        """
        notes_by_platform = {n['platform']: n for n in self.platform_notes}

        with open(path, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['platform', 'max_hops', 'description', 'notes'])
            for platform, max_hops in self.hop_limits.items():
                note = notes_by_platform.get(platform, {})
                writer.writerow([
                    platform,
                    max_hops,
                    note.get('description', ''),
                    note.get('note', '')
                ])

    def build_tree(self) -> List[Dict[str, Any]]:
        """Builds a hierarchical tree of USB devices."""
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
                    'model': attributes.get(ID_MODEL, 'Unknown model'),
                    'vendor': attributes.get(ID_VENDOR, 'Unknown vendor'),
                    'product': attributes.get('ID_MODEL', 'Unknown product')
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
            print(f"[!] Could not read USB devices: {e}")
            return []

    def _is_hub(self, attributes: Dict) -> bool:
        """Check whether a device is a hub."""
        model = attributes.get(ID_MODEL, '').lower()
        product = attributes.get('ID_MODEL', '').lower()
        return 'hub' in model or 'hub' in product

    def _find_node_by_path(self, tree: List[Dict], path: str) -> Optional[Dict]:
        """Find the node with a specific devpath in the tree."""
        for node in tree:
            if node['devpath'] == path:
                return node
            if node['children']:
                found = self._find_node_by_path(node['children'], path)
                if found:
                    return found
        return None

    def calculate_hops_and_tiers(self, tree: List[Dict]) -> Dict[str, Any]:
        """Calculates hops and tiers from the USB tree."""
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
        Assesses stability for ALL platforms based on hops.

        Color logic:
        - Green: hops < max_hops (STABLE)
        - Orange: hops == max_hops (AT LIMIT / UNCERTAIN)
        - Red: hops > max_hops (UNSTABLE)
        """
        max_hops = hops_data.get('max_hops', 0)

        # Define all platforms with their hop limits
        platforms = [
            # x86/x64
            {'id': 'windows_x86', 'name': 'Windows', 'arch': 'x86/x64', 'max_hops': self.hop_limits.get('windows_x86', 4)},
            {'id': 'mac_intel', 'name': 'Mac Intel', 'arch': 'x86/x64', 'max_hops': self.hop_limits.get('mac_intel', 7)},
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

        # Assess each platform with the color logic above
        verdicts = []
        for platform in platforms:
            max_allowed = platform['max_hops']

            # Color logic:
            # - Green: hops < max_hops
            # - Orange: hops == max_hops
            # - Red: hops > max_hops
            if max_hops < max_allowed:
                status = 'STABLE'
                color = 'green'
                emoji = '🟢'
                warning = None
            elif max_hops == max_allowed:
                status = 'AT LIMIT'
                color = 'orange'
                emoji = '🟠'
                warning = f"Max {max_allowed} hops, at the limit!"
            else:  # max_hops > max_allowed
                status = 'UNSTABLE'
                color = 'red'
                emoji = '🔴'
                warning = f"Exceeds max {max_allowed} hops!"

            verdict = {
                'id': platform['id'],
                'name': platform['name'],
                'arch': platform['arch'],
                'max_hops': max_allowed,
                'current_hops': max_hops,
                'status': status,
                'color': color,
                'emoji': emoji,
                'is_stable': max_hops <= max_allowed,
                'warning': warning
            }
            verdicts.append(verdict)

        # Group by architecture for display
        groups = {}
        for v in verdicts:
            if v['arch'] not in groups:
                groups[v['arch']] = []
            groups[v['arch']].append(v)

        return {
            'max_hops': max_hops,
            'verdicts': verdicts,
            'groups': groups,
            'overall_worst': max([v['status'] for v in verdicts],
                                key=lambda x: {'STABLE': 0, 'AT LIMIT': 1, 'UNSTABLE': 2}[x])
        }

    def get_stability_summary(self, stability_data: Dict[str, Any]) -> str:
        """Builds a text summary of the stability assessment."""
        lines = []
        lines.append("[+] Stability Verdict:")
        lines.append("-" * 60)

        groups = stability_data.get('groups', {})
        for arch, verdicts in groups.items():
            lines.append(f"\n{arch}")
            for v in verdicts:
                lines.append(f"  {v['emoji']} {v['name']}")

        # Add warnings
        warnings = [v for v in stability_data.get('verdicts', []) if v['warning']]
        if warnings:
            lines.append("\n[!] WARNINGS:")
            for w in warnings:
                lines.append(f"  - {w['name']}: {w['warning']} (current hops: {w['current_hops']})")

        return '\n'.join(lines)
