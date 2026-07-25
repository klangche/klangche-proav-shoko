"""
USB analyzer - handles the USB tree, hops, tiers and stability assessment
"""

import sys
import csv
from collections import defaultdict
from typing import Dict, List, Any, Optional, Callable
from pathlib import Path

try:
    from usbmonitor import USBMonitor
    from usbmonitor.attributes import ID_MODEL, ID_VENDOR, DEVNAME, ID_MODEL_ID, ID_VENDOR_ID
except ImportError:
    print("[!] usbmonitor is not installed. Run: pip install usb-monitor")
    sys.exit(1)


# Status ordering used to combine multiple verdicts (e.g. hops + tiers)
# into a single overall status, worst-wins.
_STATUS_RANK = {'STABLE': 0, 'AT LIMIT': 1, 'UNSTABLE': 2}


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

        # Load limits either from the given config path or the bundled
        # CSV file in assets.
        if config_path:
            self.config_path = Path(config_path)
        else:
            self.config_path = Path(__file__).parent / "assets" / "hop_limits.csv"

        self.hop_limits = self._load_limits('max_hops')
        self.tier_limits = self._load_limits('max_tiers')
        self.platform_notes = self._load_platform_notes()

    def _load_limits(self, column: str) -> Dict[str, int]:
        """Load a numeric limit column (max_hops or max_tiers) from the CSV file."""
        limits = {}

        if self.config_path and self.config_path.exists():
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        platform = row.get('platform', '').strip()
                        value = row.get(column, '').strip()
                        if platform and value.isdigit():
                            limits[platform] = int(value)
                if limits:
                    print(f"[+] Loaded {column} from: {self.config_path}")
            except Exception as e:
                print(f"[!] Could not load {column} from {self.config_path}: {e}")
                limits = {}

        if not limits:
            print(f"[!] No {column} values found in {self.config_path}, using fallback values")
            limits = self._get_fallback_limits()

        return limits

    def _get_fallback_limits(self) -> Dict[str, int]:
        """Fallback values if the CSV is missing or a column can't be read.
        Used for both max_hops and max_tiers when the CSV doesn't provide them."""
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
        Save the current hop and tier limits (and any known notes) to a CSV file.

        Args:
            path: Destination path for the CSV file.
        """
        notes_by_platform = {n['platform']: n for n in self.platform_notes}
        platforms = sorted(set(self.hop_limits) | set(self.tier_limits))

        with open(path, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['platform', 'max_hops', 'max_tiers', 'description', 'notes'])
            for platform in platforms:
                note = notes_by_platform.get(platform, {})
                writer.writerow([
                    platform,
                    self.hop_limits.get(platform, ''),
                    self.tier_limits.get(platform, ''),
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
        """Calculates hops and tiers from the USB tree.

        - hops: depth of the deepest device in the chain (root = 0)
        - tiers: number of distinct depth levels populated by a device
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

    def _evaluate(self, current: int, limit: int) -> Dict[str, Any]:
        """Compares a current value (hops or tiers) against its limit and
        returns the STABLE / AT LIMIT / UNSTABLE verdict for that dimension."""
        if current < limit:
            return {'status': 'STABLE', 'warning': None}
        elif current == limit:
            return {'status': 'AT LIMIT', 'warning': f"at the limit ({current}/{limit})"}
        else:
            return {'status': 'UNSTABLE', 'warning': f"exceeds limit ({current}/{limit})"}

    def assess_stability(self, hops_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Assesses stability for ALL platforms based on BOTH hops and tiers.

        Each platform has its own max_hops and max_tiers limit. The combined
        verdict for a platform is the worse of the two:
        - Green (STABLE): both hops and tiers are under their limit
        - Orange (AT LIMIT): the worse of the two is exactly at its limit
        - Red (UNSTABLE): the worse of the two exceeds its limit
        """
        current_hops = hops_data.get('max_hops', 0)
        current_tiers = hops_data.get('max_tiers', 0)

        # Define all platforms with their hop and tier limits
        platforms = [
            # x86/x64
            {'id': 'windows_x86', 'name': 'Windows', 'arch': 'x86/x64'},
            {'id': 'mac_intel', 'name': 'Mac Intel', 'arch': 'x86/x64'},
            {'id': 'linux_x86', 'name': 'Linux', 'arch': 'x86/x64'},
            # ARM
            {'id': 'windows_arm', 'name': 'Windows', 'arch': 'ARM'},
            {'id': 'mac_apple_silicon', 'name': 'Mac Apple Silicon', 'arch': 'ARM'},
            {'id': 'linux_arm', 'name': 'Linux', 'arch': 'ARM'},
            # Mobile
            {'id': 'iphone_lightning', 'name': 'iPhone Lightning', 'arch': 'Mobile'},
            {'id': 'iphone_usbc', 'name': 'iPhone USB-C', 'arch': 'Mobile'},
            {'id': 'samsung_usbc', 'name': 'Samsung USB-C', 'arch': 'Mobile'},
            {'id': 'android_usbc', 'name': 'Android USB-C', 'arch': 'Mobile'},
            {'id': 'ipad_lightning', 'name': 'iPad Lightning', 'arch': 'Mobile'},
            {'id': 'ipad_usbc', 'name': 'iPad USB-C', 'arch': 'Mobile'}
        ]

        verdicts = []
        for platform in platforms:
            max_hops_allowed = self.hop_limits.get(platform['id'], 4)
            max_tiers_allowed = self.tier_limits.get(platform['id'], max_hops_allowed)

            hops_eval = self._evaluate(current_hops, max_hops_allowed)
            tiers_eval = self._evaluate(current_tiers, max_tiers_allowed)

            # Combined status is the worse of the two dimensions
            if _STATUS_RANK[hops_eval['status']] >= _STATUS_RANK[tiers_eval['status']]:
                status = hops_eval['status']
            else:
                status = tiers_eval['status']

            warnings = []
            if hops_eval['warning']:
                warnings.append(f"Hops {hops_eval['warning']}")
            if tiers_eval['warning']:
                warnings.append(f"Tiers {tiers_eval['warning']}")
            warning = "; ".join(warnings) if warnings else None

            color_map = {'STABLE': 'green', 'AT LIMIT': 'orange', 'UNSTABLE': 'red'}
            emoji_map = {'STABLE': '🟢', 'AT LIMIT': '🟠', 'UNSTABLE': '🔴'}

            verdict = {
                'id': platform['id'],
                'name': platform['name'],
                'arch': platform['arch'],
                'max_hops': max_hops_allowed,
                'current_hops': current_hops,
                'max_tiers': max_tiers_allowed,
                'current_tiers': current_tiers,
                'status': status,
                'color': color_map[status],
                'emoji': emoji_map[status],
                'is_stable': status != 'UNSTABLE',
                'warning': warning
            }
            verdicts.append(verdict)

        groups = {}
        for v in verdicts:
            if v['arch'] not in groups:
                groups[v['arch']] = []
            groups[v['arch']].append(v)

        return {
            'max_hops': current_hops,
            'max_tiers': current_tiers,
            'verdicts': verdicts,
            'groups': groups,
            'overall_worst': max([v['status'] for v in verdicts], key=lambda x: _STATUS_RANK[x])
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

        warnings = [v for v in stability_data.get('verdicts', []) if v['warning']]
        if warnings:
            lines.append("\n[!] WARNINGS:")
            for w in warnings:
                lines.append(f"  - {w['name']}: {w['warning']}")

        return '\n'.join(lines)

    # --- Live monitoring -------------------------------------------------
    #
    # Wraps usbmonitor's built-in polling thread. This does NOT capture
    # protocol-level USB handshakes (link training, renegotiation, CRC
    # errors) - that would require OS-specific low-level tracing (ETW on
    # Windows, usbmon/ftrace on Linux, IOKit on macOS) and is out of scope.
    # What this does capture: every time a device connects, disconnects,
    # or re-enumerates, which in practice is a solid proxy for instability -
    # a device that's struggling to hold a link over too many hops will
    # show up as repeated connect/disconnect cycles here.

    def start_live_monitoring(
        self,
        on_connect: Callable[[str, Dict], None],
        on_disconnect: Callable[[str, Dict], None],
        check_every_seconds: float = 1.0
    ) -> None:
        """Starts background monitoring for USB connect/disconnect events."""
        self.monitor.start_monitoring(
            on_connect=on_connect,
            on_disconnect=on_disconnect,
            check_every_seconds=check_every_seconds
        )

    def stop_live_monitoring(self) -> None:
        """Stops background USB monitoring."""
        try:
            self.monitor.stop_monitoring()
        except Exception as e:
            print(f"[!] Could not stop USB monitoring cleanly: {e}")
