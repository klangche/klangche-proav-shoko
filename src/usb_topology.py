"""
USB topology — cross-platform parent-child relationships and hub grouping.

Windows:  inline PowerShell (Get-PnpDeviceProperty, no admin or .ps1 file needed)
Linux:    parsed from lsusb -t tree, matched by VID:PID
macOS:    parsed from system_profiler SPUSBDataType -json tree
"""

import sys, re, json, subprocess
from typing import Dict, List, Optional, Any

INSTANCE_PREFIX_PATTERN = re.compile(r'^MSFT(\d+)')


def get_parent_map(dev_dict: dict, nodes: dict) -> Dict[str, str]:
    """Return a dict mapping child-devpath → parent-devpath for every device.

    Falls back to platform-specific strategies; returns {} on failure.
    """
    platform = sys.platform
    if platform == 'win32':
        return _win_parent_map()
    elif platform == 'darwin':
        return _mac_parent_map(nodes)
    elif platform.startswith('linux'):
        return _linux_parent_map(nodes)
    return {}


def group_companion_hubs(nodes: dict, root_nodes: list) -> list:
    """Merge companion hubs (e.g. Apple MSFT20/MSFT30) under a virtual parent.

    Mutates *root_nodes* in place by replacing sibling hubs that share
    a common serial suffix with a single grouping node.
    """
    # Collect hubs whose instance IDs share a base serial after MSFT-prefix
    groups: Dict[str, list] = {}
    hub_devpaths: Dict[str, str] = {}  # devpath → base serial

    for name, node in nodes.items():
        if not node.get('is_hub'):
            continue
        devpath = node.get('devpath', name)
        inst = devpath.split('\\')[-1]  # last segment = instance ID
        m = INSTANCE_PREFIX_PATTERN.match(inst)
        if m:
            base = inst[m.end():]       # serial after MSFTxx
            groups.setdefault(base, []).append(node)
            hub_devpaths[devpath] = base

    if not groups:
        return root_nodes

    # Find hubs in root_nodes that belong to a group
    base_to_root_idx: Dict[str, list] = {}
    for i, rn in enumerate(root_nodes):
        dp = rn.get('devpath', '')
        base = hub_devpaths.get(dp)
        if base and base in groups and len(groups[base]) >= 2:
            base_to_root_idx.setdefault(base, []).append(i)

    if not base_to_root_idx:
        return root_nodes

    # Build new root_nodes list
    new_roots = []
    replaced = set()
    for base, indices in base_to_root_idx.items():
        group_nodes = [root_nodes[idx] for idx in indices]
        # Derive a label from the Billboard child or vendor info
        label = _group_label(group_nodes, nodes)
        virtual = {
            'name': label,
            'devpath': f'\\virtual_{base}',
            'attributes': {},
            'children': group_nodes,
            'is_hub': False,
            'is_virtual': True,
            'model': label,
            'vendor': 'Generic',
            'is_composite_interface': False,
            'hub_id': '',
            'port': 0,
            'depth': 1,
            'is_internal': False,
        }
        new_roots.append(virtual)
        for idx in indices:
            replaced.add(idx)

    for i, rn in enumerate(root_nodes):
        if i not in replaced:
            new_roots.append(rn)

    return new_roots


def _group_label(group_nodes: list, all_nodes: dict) -> str:
    """Build a label for the companion-hub group from Billboard or VID."""
    # Prefer the Billboard child's model if available
    for gn in group_nodes:
        for child in gn.get('children', []):
            cm = child.get('model', '')
            if cm and 'multiport' in cm.lower() or 'adapter' in cm.lower():
                return cm
    # Fall back to vendor name + VID
    for gn in group_nodes:
        vid = gn.get('vid', '')
        if vid:
            from usbmonitor.attributes import ID_VENDOR_FROM_DATABASE
            # Try to get a nice vendor name
            try:
                import usbmonitor
                db = usbmonitor.USBMonitor._load_usb_ids()
            except Exception:
                db = None
            vendor_name = f'VID_{vid}'
            # Simple fallback
            return f'USB-C Adapter ({vendor_name})'
    return 'USB Hub Group'


def _win_parent_map() -> Dict[str, str]:
    """Query Windows PnP parent relationships via inline PowerShell."""
    ps_cmd = (
        'Get-PnpDevice -PresentOnly | '
        'Where-Object { $_.InstanceId -like \"USB*\" -or $_.InstanceId -like \"USBSTOR*\" } | '
        'ForEach-Object { '
        '$p = $null; '
        'try { $pp = $_ | Get-PnpDeviceProperty -KeyName \"DEVPKEY_Device_Parent\" -ErrorAction Stop; '
        '$p = $pp.Data } catch {}; '
        '[PSCustomObject]@{ InstanceId = $_.InstanceId; Parent = $p } '
        '} | ConvertTo-Json -Compress'
    )
    try:
        out = subprocess.run(
            ['powershell', '-NoProfile', '-Command', ps_cmd],
            capture_output=True, text=True, timeout=30
        )
        if not out.stdout:
            return {}
        pairs = json.loads(out.stdout)
        if isinstance(pairs, dict):
            pairs = [pairs]
        return {p['InstanceId']: p['Parent'] for p in pairs if p.get('Parent')}
    except Exception:
        return {}


def _mac_parent_map(nodes: dict) -> Dict[str, str]:
    """Build USB parent map from macOS system_profiler output."""
    try:
        out = subprocess.run(
            ['system_profiler', 'SPUSBDataType', '-json'],
            capture_output=True, text=True, timeout=30
        )
        data = json.loads(out.stdout)
        items = data.get('SPUSBDataType', [])
    except Exception:
        return {}

    # Build a VID:PID → locationID map from the system_profiler tree
    vidpid_to_loc = {}
    parent_map = {}

    def walk(entries, parent_loc=None):
        for item in entries:
            loc = item.get('location_id', '')
            vendor_id = item.get('vendor_id', '')
            product_id = item.get('product_id', '')
            if vendor_id and product_id:
                vid = vendor_id.replace('0x', '').zfill(4).upper()
                pid = product_id.replace('0x', '').zfill(4).upper()
                key = f'{vid}:{pid}'
                vidpid_to_loc[key] = (loc, parent_loc)
            children = item.get('_items', [])
            if children:
                walk(children, loc)

    walk(items)

    # Match each node from USBMonitor by VID:PID and assign parent
    result = {}
    for name, node in nodes.items():
        vid = node.get('vid', '').upper()
        pid = node.get('pid', '').upper()
        key = f'{vid}:{pid}'
        if key in vidpid_to_loc:
            loc, parent_loc = vidpid_to_loc[key]
            devpath = node.get('devpath', name)
            if parent_loc:
                # Find the parent's devpath
                for pname, pnode in nodes.items():
                    pvid = pnode.get('vid', '').upper()
                    ppid = pnode.get('pid', '').upper()
                    pkey = f'{pvid}:{ppid}'
                    ploc, _ = vidpid_to_loc.get(pkey, (None, None))
                    if ploc == parent_loc:
                        result[devpath] = pnode.get('devpath', pname)
                        break
    return result


def _linux_parent_map(nodes: dict) -> Dict[str, str]:
    """Build USB parent map from Linux lsusb -t and lsusb output."""
    try:
        out_tree = subprocess.run(
            ['lsusb', '-t'], capture_output=True, text=True, timeout=10
        )
        out_list = subprocess.run(
            ['lsusb'], capture_output=True, text=True, timeout=10
        )
    except Exception:
        return {}
    if not out_tree.stdout or not out_list.stdout:
        return {}

    # Parse lsusb list: Bus XXX Device YYY → VID:PID
    usb_dev = {}
    for line in out_list.stdout.splitlines():
        m = re.search(r'Bus\s+(\d+)\s+Device\s+(\d+):\s+ID\s+([0-9a-fA-F]+):([0-9a-fA-F]+)', line)
        if m:
            bus, dev = int(m.group(1)), int(m.group(2))
            vid, pid = m.group(3).upper(), m.group(4).upper()
            usb_dev[(bus, dev)] = {'vid': vid, 'pid': pid}

    # Parse lsusb -t tree to build parent bus:dev → child bus:dev
    parent_of = {}  # (bus, dev) → parent (bus, dev)
    stack = []      # (bus, dev, indent)
    for line in out_tree.stdout.splitlines():
        stripped = line.lstrip()
        indent = len(line) - len(stripped)
        m = re.search(r'Dev\s+(\d+)', stripped)
        if not m:
            continue
        devnum = int(m.group(1))
        # Extract bus number from the first match
        bus_m = re.search(r'Bus\s+(\d+)', stripped)
        bus = int(bus_m.group(1)) if bus_m else 1
        # Pop stack to correct indent level
        while stack and stack[-1][2] >= indent:
            stack.pop()
        # Parent is top of stack
        if stack:
            parent_of[(bus, devnum)] = (stack[-1][0], stack[-1][1])
        stack.append((bus, devnum, indent))

    # Match nodes by VID:PID
    result = {}
    for name, node in nodes.items():
        vid = node.get('vid', '').upper()
        pid = node.get('pid', '').upper()
        if not vid or not pid:
            continue
        # Find matching bus:dev
        for (bus, dev), info in usb_dev.items():
            if info['vid'] == vid and info['pid'] == pid:
                devpath = node.get('devpath', name)
                parent_bus_dev = parent_of.get((bus, dev))
                if parent_bus_dev:
                    # Find the parent node by VID:PID
                    pb, pd = parent_bus_dev
                    pinfo = usb_dev.get((pb, pd))
                    if pinfo:
                        for pname, pnode in nodes.items():
                            pvid = pnode.get('vid', '').upper()
                            ppid = pnode.get('pid', '').upper()
                            if pvid == pinfo['vid'] and ppid == pinfo['pid']:
                                result[devpath] = pnode.get('devpath', pname)
                                break
    return result
