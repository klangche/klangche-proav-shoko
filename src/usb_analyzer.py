"""
USB analyzer - handles the USB tree, hops, tiers and stability assessment
"""

import sys
import csv
import platform
from collections import defaultdict
from typing import Dict, List, Any, Optional, Callable
from pathlib import Path

try:
    from usbmonitor import USBMonitor
    from usbmonitor.attributes import ID_MODEL, ID_VENDOR, DEVNAME, ID_MODEL_ID, ID_VENDOR_ID, ID_MODEL_FROM_DATABASE, ID_VENDOR_FROM_DATABASE
except ImportError:
    USBMonitor = None

try:
    from .usb_topology import get_parent_map, group_companion_hubs
except ImportError:
    from usb_topology import get_parent_map, group_companion_hubs


# Status ordering used to combine multiple verdicts (e.g. hops + tiers)
# into a single overall status, worst-wins.
_STATUS_RANK = {'STABLE': 0, 'AT LIMIT': 1, 'UNSTABLE': 2}


class USBAnalyzer:
    """Analyzes USB devices and builds a hierarchical tree."""

    def __init__(self, config_path: Optional[str] = None):
        """
        Initialize the USB analyzer.

        Args:
            config_path: Optional path to a usb_data.txt file to load
                instead of the one bundled in src/assets.
        """
        self.monitor = USBMonitor()
        self.devices = {}

        # Load limits either from the given config path or the bundled
        # data file in assets.
        if config_path:
            self.config_path = Path(config_path)
        else:
            self.config_path = Path(__file__).parent / "assets" / "usb_data.csv"

        self.hop_limits = self._load_limits('max_hops')
        self.tier_limits = self._load_limits('max_tiers')
        self.hub_limits = self._load_limits('max_hubs')
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
                    pass
            except Exception as e:
                print(f"Could not load {column} from {self.config_path}: {e}")
                limits = {}

        if not limits:
            print(f"No {column} values found in {self.config_path}, using fallback values")
            limits = self._get_fallback_limits(column)

        return limits

    def _get_fallback_limits(self, column: str = 'max_hops') -> Dict[str, int]:
        """Fallback values if the CSV is missing or a column can't be read."""
        base = {
            'windows_x86': 4, 'windows_arm': 4, 'mac_intel': 7,
            'mac_apple_silicon': 3, 'linux_x86': 4, 'linux_arm': 4,
            'iphone_lightning': 2, 'iphone_usbc': 3, 'samsung_usbc': 4,
            'android_usbc': 4, 'ipad_lightning': 2, 'ipad_usbc': 3
        }
        if column == 'max_hubs':
            return {
                'windows_x86': 4, 'windows_arm': 4, 'mac_intel': 5,
                'mac_apple_silicon': 2, 'linux_x86': 4, 'linux_arm': 4,
                'iphone_lightning': 1, 'iphone_usbc': 2, 'samsung_usbc': 3,
                'android_usbc': 3, 'ipad_lightning': 1, 'ipad_usbc': 2
            }
        return base

    def _load_platform_notes(self) -> List[Dict[str, str]]:
        """Load notes from the CSV file."""
        notes = []
        if self.config_path and self.config_path.exists():
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        platform = row.get('platform', '').strip()
                        system = row.get('system', '').strip()
                        description = row.get('description', '').strip()
                        note = row.get('notes', '').strip()
                        if platform:
                            notes.append({
                                'platform': platform,
                                'system': system,
                                'description': description,
                                'note': note
                            })
            except Exception as e:
                print(f"Could not read notes: {e}")
        return notes

    def get_platform_notes(self) -> List[Dict[str, str]]:
        """Get notes for use in reports."""
        return self.platform_notes

    def save_usb_data(self, path: str) -> None:
        """
        Save the current USB limits data to a file.

        Args:
            path: Destination path for the data file.
        """
        notes_by_platform = {n['platform']: n for n in self.platform_notes}
        platforms = sorted(set(self.hop_limits) | set(self.tier_limits) | set(self.hub_limits))

        with open(path, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['platform', 'max_hops', 'max_tiers', 'max_hubs', 'description', 'notes'])
            for platform in platforms:
                note = notes_by_platform.get(platform, {})
                writer.writerow([
                    platform,
                    self.hop_limits.get(platform, ''),
                    self.tier_limits.get(platform, ''),
                    self.hub_limits.get(platform, ''),
                    note.get('description', ''),
                    note.get('note', '')
                ])

    def _parse_devpath(self, devname: str) -> Dict[str, Any]:
        """Parse a Windows device instance path into components."""
        result = {'hub_id': '', 'port': 0, 'is_composite_interface': False, 'depth': 0, 'interface_number': None}

        if not devname:
            return result

        parts = devname.split('\\')
        if len(parts) < 3:
            result['depth'] = 0
            return result

        hwid = parts[1]  # e.g. VID_0B05&PID_1ACE or VID_0B05&PID_1ACE&MI_00
        instance = parts[2]  # e.g. T6MPKRD00HWM5 or 6&1A89C1F2&0&0000

        result['is_composite_interface'] = '&MI_' in hwid
        if result['is_composite_interface']:
            mi_match = __import__('re').search(r'&MI_(\d+)', hwid)
            if mi_match:
                result['interface_number'] = int(mi_match.group(1))

        # Parse Windows hub instance path: HubHi&HubLo&Flags&Port
        instance_parts = instance.split('&')
        if len(instance_parts) >= 4 and all(p.isdigit() or all(c in '0123456789ABCDEFabcdef' for c in p) for p in instance_parts[:2]):
            result['hub_id'] = f"{instance_parts[0]}&{instance_parts[1]}"
            try:
                result['port'] = int(instance_parts[3])
            except ValueError:
                result['port'] = 0
            result['depth'] = 2
        elif not instance_parts[0].startswith('6&') and not instance_parts[0].startswith('5&'):
            # Serial number or simple instance - root level
            result['depth'] = 0
        else:
            result['depth'] = 1

        return result

    @staticmethod
    def _describe_interface(attrs: dict, path_info: dict) -> str:
        """Return a human-readable interface description from USB class info."""
        import re
        # Parse protocol from ID_USB_INTERFACES
        ifaces = attrs.get('ID_USB_INTERFACES', ())
        cls_code = 0
        prot = 0
        subcls = 0
        if ifaces:
            best = ifaces[0]
            cm = re.search(r'Class_(\d+)', best)
            if cm:
                cls_code = int(cm.group(1))
            pm = re.search(r'Prot_(\d+)', best)
            if pm:
                prot = int(pm.group(1))
            sm = re.search(r'SubClass_(\d+)', best)
            if sm:
                subcls = int(sm.group(1))

        iface_class = attrs.get('ID_USB_CLASS_FROM_DATABASE', '')

        # For HID devices, use protocol to distinguish Keyboard/Mouse
        if cls_code == 3 or 'HID' in iface_class:
            if prot == 1:
                return 'HID Keyboard'
            elif prot == 2:
                return 'HID Mouse'
            elif prot == 0 and subcls == 1:
                return 'HID Keyboard'
            if iface_class and 'HID' in iface_class:
                return 'HID'
            return 'HID'

        # For known class codes, return descriptive name
        class_map = {
            1: 'Audio', 2: 'Communications', 5: 'Physical',
            6: 'Image', 7: 'Printer', 8: 'Mass Storage',
            9: 'Hub', 10: 'CDC Data', 14: 'Camera',
            224: 'Wireless', 239: 'Miscellaneous',
        }
        if cls_code in class_map:
            return class_map[cls_code]

        # Fall back to ID_USB_CLASS_FROM_DATABASE for Windows-provided names
        if iface_class and 'Class' not in iface_class and 'class' not in iface_class:
            return iface_class

        if cls_code:
            return f'Class_{cls_code:02X}'
        return ''

    def _is_internal(self, node: Dict) -> bool:
        """Heuristic to determine if a device is internal to the computer."""
        model = node.get('model', '').lower()
        name = node.get('name', '').lower()
        vendor = node.get('vendor', '').lower()
        devpath = node.get('devpath', '').lower()
        combined = f"{model} {name} {vendor}"

        # Only truly integrated/internal components
        internal_kw = ['integrated', 'bluetooth', 'camera', 'intel wireless', 'razer blade']
        if any(kw in combined for kw in internal_kw):
            return True

        # Internal hubs (part of system)
        if 'hub' in model and 'usb' in model:
            return True

        # SD card readers and mass storage are EXTERNAL unless explicitly integrated
        # Don't flag 'semc', 'sd card', 'card reader', 'mass storage' as internal

        return False

    def build_tree(self) -> List[Dict[str, Any]]:
        """Builds a hierarchical tree of USB devices."""
        try:
            dev_dict = self.monitor.get_available_devices()
            self.devices = dev_dict

            # Phase 1: create nodes with parsed Windows path info
            nodes = {}
            for dev_name, attrs in dev_dict.items():
                devname = attrs.get('DEVNAME', dev_name)
                path_info = self._parse_devpath(devname)

                model = attrs.get(ID_MODEL_FROM_DATABASE, attrs.get(ID_MODEL, 'Unknown'))
                vendor = attrs.get(ID_VENDOR_FROM_DATABASE, attrs.get(ID_VENDOR, 'Unknown'))
                vid = attrs.get(ID_VENDOR_ID, '')
                pid = attrs.get(ID_MODEL_ID, '')

                # Build a descriptive label showing VID:PID when model is generic
                if vid and pid:
                    device_info = "VID_{0}&PID_{1}".format(vid, pid)
                else:
                    device_info = ""

                # Parse USB interface class info for human-readable type
                iface_class = attrs.get('ID_USB_CLASS_FROM_DATABASE', '')
                iface_desc = self._describe_interface(attrs, path_info)

                # Detect hubs: check English "hub", Swedish "nav", and USB interface class
                model_lower = attrs.get(ID_MODEL, '').lower()
                ifaces_raw = attrs.get('ID_USB_INTERFACES', ())
                is_hub = 'hub' in model_lower or 'nav' in model_lower
                if not is_hub:
                    for ifc in ifaces_raw:
                        if 'HUB' in ifc.upper():
                            is_hub = True
                            break

                nodes[dev_name] = {
                    'name': dev_name,
                    'devpath': devname,
                    'attributes': attrs,
                    'children': [],
                    'is_hub': is_hub,
                    'model': model,
                    'vendor': vendor,
                    'vid': vid,
                    'pid': pid,
                    'device_info': device_info,
                    'is_composite_interface': path_info['is_composite_interface'],
                    'interface_number': path_info['interface_number'],
                    'interface_class': iface_class,
                    'interface_desc': iface_desc,
                    'hub_id': path_info['hub_id'],
                    'port': path_info['port'],
                    'depth': path_info['depth']
                }
                nodes[dev_name]['is_internal'] = self._is_internal(nodes[dev_name])

            # Phase 2: build tree structure using platform topology + companion hub grouping
            parent_map = get_parent_map(dev_dict, nodes)
            root_nodes = []
            assigned = set()

            # Helper: find node by devpath (case-insensitive match)
            def find_node(devpath):
                lower = devpath.lower()
                for n_name, n_node in nodes.items():
                    if n_node.get('devpath', '').lower() == lower:
                        return n_node
                return None

            # Attach each node to its parent, or mark as root
            for name, node in nodes.items():
                devpath = node.get('devpath', name)
                parent_devpath = parent_map.get(devpath)

                if parent_devpath:
                    parent = find_node(parent_devpath)
                    if parent:
                        parent.setdefault('children', []).append(node)
                        assigned.add(name)
                        continue
                # No parent found in our nodes → root node
                root_nodes.append(node)
                assigned.add(name)

            # Group companion hubs (e.g. Apple MSFT20/MSFT30) under virtual parent
            root_nodes = group_companion_hubs(nodes, root_nodes)

            # Propagate is_internal from children to parents
            def propagate_internal(nodes_list):
                for n in nodes_list:
                    if n.get('children'):
                        propagate_internal(n['children'])
                        if any(c.get('is_internal') for c in n['children']):
                            n['is_internal'] = True

            propagate_internal(root_nodes)

            # Calculate hops (depth) for each node
            def assign_hops(nodes_list, depth=0):
                for n in nodes_list:
                    n['hops'] = depth
                    if n.get('children'):
                        assign_hops(n['children'], depth + 1)

            assign_hops(root_nodes, 0)

            # Wrap everything under a root node representing this computer
            tree = [{
                'name': 'This Computer',
                'devpath': '',
                'attributes': {},
                'children': root_nodes,
                'is_hub': False,
                'model': platform.node(),
                'vendor': platform.system(),
                'is_composite_interface': False,
                'hub_id': '',
                'port': 0,
                'depth': 0,
                'hops': 0
            }]

            return tree

        except Exception as e:
            print(f"Could not read USB devices: {e}")
            return []

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
            devices_by_hops[depth].append(node['name'])
            depths.append(depth)

            for child in node.get('children', []):
                next_depth = depth if child.get('is_virtual') else depth + 1
                traverse(child, next_depth)

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
        """Compares a current value (hops or tiers) against its limit."""
        if current < limit:
            return {'status': 'STABLE', 'warning': None}
        elif current == limit:
            return {'status': 'AT LIMIT', 'warning': f"at the limit ({current}/{limit})"}
        else:
            return {'status': 'UNSTABLE', 'warning': f"exceeds limit ({current}/{limit})"}

    _PLATFORMS = [
        {'id': 'windows_x86', 'name': 'Windows', 'arch': 'x86/x64'},
        {'id': 'mac_intel', 'name': 'Mac Intel', 'arch': 'x86/x64'},
        {'id': 'linux_x86', 'name': 'Linux', 'arch': 'x86/x64'},
        {'id': 'windows_arm', 'name': 'Windows', 'arch': 'ARM'},
        {'id': 'mac_apple_silicon', 'name': 'Mac Apple Silicon', 'arch': 'ARM'},
        {'id': 'linux_arm', 'name': 'Linux', 'arch': 'ARM'},
        {'id': 'iphone_lightning', 'name': 'iPhone Lightning', 'arch': 'Mobile'},
        {'id': 'iphone_usbc', 'name': 'iPhone USB-C', 'arch': 'Mobile'},
        {'id': 'samsung_usbc', 'name': 'Samsung USB-C', 'arch': 'Mobile'},
        {'id': 'android_usbc', 'name': 'Android USB-C', 'arch': 'Mobile'},
        {'id': 'ipad_lightning', 'name': 'iPad Lightning', 'arch': 'Mobile'},
        {'id': 'ipad_usbc', 'name': 'iPad USB-C', 'arch': 'Mobile'}
    ]

    # Room name feature (disabled by default). Set to True to activate.
    ROOM_NAME_ENABLED = False

    def _platform_verdicts(self, current_hops: int, current_tiers: int, current_hubs: int = 0) -> list:
        """Generate per-platform verdicts for given hops/tiers/hubs."""
        notes_by_platform = {n['platform']: n for n in self.platform_notes}
        verdicts = []
        for platform in self._PLATFORMS:
            max_hops_allowed = self.hop_limits.get(platform['id'], 4)
            max_tiers_allowed = self.tier_limits.get(platform['id'], max_hops_allowed)
            max_hubs_allowed = self.hub_limits.get(platform['id'], max_hops_allowed)

            hops_eval = self._evaluate(current_hops, max_hops_allowed)
            tiers_eval = self._evaluate(current_tiers, max_tiers_allowed)
            hubs_eval = self._evaluate(current_hubs, max_hubs_allowed)

            combined = [hops_eval, tiers_eval, hubs_eval]
            worst = max(combined, key=lambda x: _STATUS_RANK.get(x['status'], 0))
            status = worst['status']

            warnings = []
            if hops_eval['warning']:
                warnings.append(f"Hops {hops_eval['warning']}")
            if tiers_eval['warning']:
                warnings.append(f"Tiers {tiers_eval['warning']}")
            if hubs_eval['warning']:
                warnings.append(f"Hubs {hubs_eval['warning']}")

            color_map = {'STABLE': 'green', 'AT LIMIT': 'orange', 'UNSTABLE': 'red'}
            emoji_map = {'STABLE': '🟢', 'AT LIMIT': '🟠', 'UNSTABLE': '🔴'}

            note = notes_by_platform.get(platform['id'], {})

            verdicts.append({
                'id': platform['id'],
                'name': platform['name'],
                'arch': platform['arch'],
                'description': note.get('system', platform['name']),
                'max_hops': max_hops_allowed,
                'current_hops': current_hops,
                'max_tiers': max_tiers_allowed,
                'current_tiers': current_tiers,
                'max_hubs': max_hubs_allowed,
                'current_hubs': current_hubs,
                'status': status,
                'color': color_map[status],
                'emoji': emoji_map[status],
                'is_stable': status != 'UNSTABLE',
                'warning': "; ".join(warnings) if warnings else None
            })
        return verdicts

    def _count_hubs_in_path(self, tree: list) -> int:
        """Count the maximum number of hubs along any path in the tree."""
        def traverse(node, hub_count):
            max_hubs = hub_count
            if node.get('is_hub'):
                hub_count += 1
                max_hubs = hub_count
            for child in node.get('children', []):
                child_hubs = traverse(child, hub_count)
                max_hubs = max(max_hubs, child_hubs)
            return max_hubs

        max_hubs = 0
        for root in tree:
            max_hubs = max(max_hubs, traverse(root, 0))
        return max_hubs

    def assess_stability(self, hops_data: Dict[str, Any], tree: Optional[list] = None) -> Dict[str, Any]:
        """
        Assesses stability per USB port and overall.

        Each root-level device or hub is treated as a USB port, with its own
        hops/tiers and per-platform verdicts.
        """
        current_hops = hops_data.get('max_hops', 0)
        current_tiers = hops_data.get('max_tiers', 0)
        current_hubs = self._count_hubs_in_path(tree) if tree else 0

        overall_verdicts = self._platform_verdicts(current_hops, current_tiers, current_hubs)

        ports = []
        if tree:
            root = tree[0] if tree else {}
            for i, child in enumerate(root.get('children', [])):
                # Get all devices under this port
                devices_under = self._walk_devices(child)
                
                if not devices_under:
                    continue

                # Calculate hops and tiers for this port
                # The port itself (child) is at depth relative to root
                # We need to count hubs and tiers from this port downward
                
                port_hops = 0
                port_tiers = 0
                max_hops_from_port = 0
                hubs_in_chain = set()
                depth_levels = set()
                
                # Build a mapping of device -> its depth from the port
                def collect_depths(node, depth_from_port):
                    nonlocal max_hops_from_port, port_tiers
                    max_hops_from_port = max(max_hops_from_port, depth_from_port)
                    depth_levels.add(depth_from_port)
                    
                    # Count hubs
                    if node.get('is_hub'):
                        hub_id = node.get('hub_id')
                        if not hub_id:
                            vid = node.get('vid', '')
                            pid = node.get('pid', '')
                            hub_id = f"{vid}:{pid}" if vid and pid else node.get('name', '')
                        hubs_in_chain.add(hub_id)
                    
                    for child in node.get('children', []):
                        next_depth = depth_from_port if child.get('is_virtual') else depth_from_port + 1
                        collect_depths(child, next_depth)
                
                # Start from the port itself (depth 0)
                collect_depths(child, 0)
                
                # Hops = max depth from port (how many levels down)
                port_hops = max_hops_from_port
                
                # Tiers = number of depth levels with devices (including port itself)
                # Port itself is tier 1, each level down adds a tier
                port_tiers = len(depth_levels) if depth_levels else 1
                
                # Count external hubs (excluding the root port if it's a hub)
                # Hubs in the chain from this port
                external_hubs = len(hubs_in_chain)
                
                # All nodes under this port (including port node)
                all_under = self._walk_devices(child)
                # Endpoint devices: if port has children, count children only;
                # if port is a leaf (no children), count the port itself as 1 endpoint.
                children_only = [d for d in all_under if d is not child]
                if children_only:
                    endpoint_devices = children_only
                else:
                    endpoint_devices = [child]  # leaf device = endpoint
                endpoint_names = [
                    d.get('model', d.get('name', 'Unknown'))
                    for d in endpoint_devices
                ]

                model = child.get('model', child.get('name', ''))
                device_info = child.get('device_info', '')
                iface_desc = child.get('interface_desc', '')
                iface_num = child.get('interface_number')
                if child.get('is_composite_interface'):
                    mi = f"MI_{iface_num:02d}" if iface_num is not None else ""
                    if model and 'USB-enhet' not in model and 'sammansatt' not in model and 'Composite' not in model:
                        label = model
                        if mi:
                            label += f" {mi}"
                    elif iface_desc:
                        iface_tag = "HID Keyboard" if "Keyboard" in iface_desc else \
                                    "HID Mouse" if "Mouse" in iface_desc else \
                                    iface_desc
                        label = f"{iface_tag} {mi}".strip()
                    else:
                        label = model
                    if device_info:
                        label += f" ({device_info})"
                else:
                    if device_info:
                        model = f"{model} ({device_info})"
                    label = model if not child.get('is_hub') else f"{model}"

                ports.append({
                    'id': i + 1,
                    'label': label,
                    'max_hops': port_hops,
                    'max_tiers': port_tiers,
                    'external_hubs': external_hubs,
                    'devices': endpoint_names,
                    'verdicts': self._platform_verdicts(port_hops, port_tiers, external_hubs)
                })

        # Group overall verdicts by arch for convenience
        groups = {}
        for v in overall_verdicts:
            groups.setdefault(v['arch'], []).append(v)

        total_endpoints = sum(len(p.get('devices', [])) for p in ports)

        return {
            'max_hops': current_hops,
            'max_tiers': current_tiers,
            'max_hubs': current_hubs,
            'total_endpoints': total_endpoints,
            'verdicts': overall_verdicts,
            'groups': groups,
            'ports': ports,
            'overall_worst': max(
                (v['status'] for v in overall_verdicts),
                key=lambda x: _STATUS_RANK.get(x, 0)
            )
        }

    def _walk_devices(self, node: Dict) -> list:
        """Walk all non-root devices under a node."""
        result = []
        stack = [node]
        while stack:
            n = stack.pop()
            result.append(n)
            stack.extend(n.get('children', []))
        return result

    def get_stability_summary(self, stability_data: Dict[str, Any]) -> str:
        """Builds a text summary of the stability assessment, per USB port."""
        lines = []
        overall = stability_data.get('overall_worst', 'STABLE')
        total = stability_data.get('total_endpoints', 0)
        ep_label = "endpoint" if total == 1 else "endpoints"
        lines.append(f"Overall stability: {overall} ({total} {ep_label}, hops={stability_data.get('max_hops', 0)}, tiers={stability_data.get('max_tiers', 0)}, hubs={stability_data.get('max_hubs', 0)})")
        lines.append("")

        ports = stability_data.get('ports', [])
        if ports:
            for port in ports:
                label = port.get('label', f"Port {port.get('id', '')}")
                ph = port.get('max_hops', 0)
                pt = port.get('max_tiers', 0)
                p_hub = port.get('external_hubs', 0)
                dc_count = len(port['devices'])
                ep_label = "endpoint" if dc_count == 1 else "endpoints"
                lines.append(f"  {label} ({dc_count} {ep_label}, hops={ph}, tiers={pt}, hubs={p_hub})")
                for v in port['verdicts']:
                    status_char = '+' if v['color'] == 'green' else ('~' if v['color'] == 'orange' else '!')
                    hubs_str = "hubs {ch}/{mh}  ".format(ch=v.get('current_hubs', 0), mh=v.get('max_hubs', 0)) if 'current_hubs' in v else ""
                    lines.append(
                        f"    {status_char} {v['name']:<20s} "
                        f"{v['status']:<9s} "
                        f"hops {v['current_hops']}/{v['max_hops']}  "
                        f"tiers {v['current_tiers']}/{v['max_tiers']}  "
                        f"{hubs_str}"
                    )
                lines.append("")

        lines.append("  Overall (all ports combined)")
        for v in stability_data.get('verdicts', []):
            status_char = '+' if v['color'] == 'green' else ('~' if v['color'] == 'orange' else '!')
            hubs_str = "hubs {ch}/{mh}  ".format(ch=v.get('current_hubs', 0), mh=v.get('max_hubs', 0)) if 'current_hubs' in v else ""
            lines.append(
                f"    {status_char} {v['name']:<20s} "
                f"{v['status']:<9s} "
                f"hops {v['current_hops']}/{v['max_hops']}  "
                f"tiers {v['current_tiers']}/{v['max_tiers']}  "
                f"{hubs_str}"
            )

        warnings = [v for v in stability_data.get('verdicts', []) if v.get('warning')]
        if warnings:
            lines.append("")
            lines.append("  WARNINGS:")
            for w in warnings:
                lines.append(f"    - {w['name']}: {w['warning']}")

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
            print(f"Could not stop USB monitoring cleanly: {e}")
