#!/usr/bin/env python3
"""
ProAV Shoko - CLI mode with logging, port selection, and report format options
"""

import sys
import threading
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

# Ensure UTF-8 output for Unicode box-drawing characters
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

from src.usb_analyzer import USBAnalyzer
from src.display_analyzer import DisplayAnalyzer
from src.report_generator import ReportGenerator
from src.platform_utils import PlatformUtils


class LogManager:
    """Manages logging with ability to stop and generate report."""
    
    def __init__(self):
        self.logs = []
        self.lock = threading.Lock()
        self.stop_event = threading.Event()
        self.monitor_thread = None
        self.print_callback = None  # Callback to print logs in real-time
        # Track device reconnections for stability assessment
        self.device_events = {}  # device_name -> {'connects': 0, 'disconnects': 0, 'last_action': None}
        self.unstable_devices = set()  # Devices that reconnected
    
    def set_print_callback(self, callback):
        """Set callback for real-time log printing."""
        self.print_callback = callback
    
    def log(self, message: str):
        with self.lock:
            log_entry = f"[{time.strftime('%H:%M:%S')}] {message}"
            self.logs.append(log_entry)
            # Print in real-time if callback is set
            if self.print_callback:
                self.print_callback(log_entry)
            # Track device reconnections
            self._track_device_event(message)
    
    def _track_device_event(self, message: str):
        """Track device connect/disconnect events to detect reconnections."""
        import re
        # Match CONNECTED: or DISCONNECTED: messages
        match = re.match(r'\[.*?\] (CONNECTED|DISCONNECTED): (.+)', message)
        if match:
            action, device = match.groups()
            if device not in self.device_events:
                self.device_events[device] = {'connects': 0, 'disconnects': 0, 'last_action': None}
            if action == 'CONNECTED':
                self.device_events[device]['connects'] += 1
                # If this device was disconnected before, it's a reconnection
                if self.device_events[device]['last_action'] == 'DISCONNECTED':
                    self.unstable_devices.add(device)
            elif action == 'DISCONNECTED':
                self.device_events[device]['disconnects'] += 1
            self.device_events[device]['last_action'] = action
    
    def get_unstable_devices(self) -> set:
        """Get set of devices that reconnected during monitoring."""
        with self.lock:
            return self.unstable_devices.copy()
    
    def get_logs(self) -> list:
        with self.lock:
            return self.logs.copy()
    
    def start_monitoring(self, usb_analyzer):
        """Start background USB monitoring."""
        def monitor():
            try:
                usb_analyzer.start_live_monitoring(
                    on_connect=lambda dev_id, info: self.log(f"CONNECTED: {info.get('ID_MODEL', dev_id)}"),
                    on_disconnect=lambda dev_id, info: self.log(f"DISCONNECTED: {info.get('ID_MODEL', dev_id)}"),
                    check_every_seconds=0.1
                )
                while not self.stop_event.is_set():
                    time.sleep(0.1)
            except Exception as e:
                self.log(f"Monitor error: {e}")
        
        self.monitor_thread = threading.Thread(target=monitor, daemon=True)
        self.monitor_thread.start()
        self.log("Monitoring started...")
    
    def stop_monitoring(self, usb_analyzer):
        """Stop background USB monitoring."""
        self.stop_event.set()
        try:
            usb_analyzer.stop_live_monitoring()
        except:
            pass
        self.log("Monitoring stopped.")


def _print_header():
    print("\n" + "=" * 70)
    print("KLANGCHE PROAV SHOKO - USB DETECTIVE")
    print("=" * 70 + "\n")


def _print_separator():
    print("-" * 70)


def _print_section_header(title):
    print(f"\n{title}")
    print("-" * 70)


def _print_tag(tag: str):
    """Print a machine-parseable tag for GUI parsing. Only tag name, no [TAG:] wrapper."""
    print(tag)


def _node_label(node):
    """Build a display label with interface type, model name and VID:PID."""
    model = node.get('model', node.get('name', 'Unknown'))
    device_info = node.get('device_info', '')
    iface_desc = node.get('interface_desc', '')
    iface_num = node.get('interface_number')
    if node.get('is_composite_interface'):
        mi = f"MI_{iface_num:02d}" if iface_num is not None else ""
        suffix = f" ({device_info})" if device_info else ""
        # Use model (ID_MODEL_FROM_DATABASE) when it's a specific name, not generic
        if model and 'USB-enhet' not in model and 'sammansatt' not in model and 'Composite' not in model:
            label = model
            if mi:
                label += f" {mi}"
            return f"{label}{suffix}"
        # Fall back to interface type description
        if iface_desc:
            iface_tag = f"HID Keyboard" if "Keyboard" in iface_desc else \
                        f"HID Mouse" if "Mouse" in iface_desc else \
                        iface_desc
            return f"{iface_tag} {mi}{suffix}".strip()
    return f"{model} ({device_info})" if device_info else model

def _print_tree(nodes, prefix="", _show_internal=False, _parent_is_internal=False):
    for i, node in enumerate(nodes):
        is_last = i == len(nodes) - 1
        connector = "└── " if is_last else "├── "

        badges = []
        if node.get('is_hub'):
            badges.append('HUB')
        if node.get('is_display'):
            badges.append('DISPLAY')
        if node.get('is_internal') and _show_internal and not _parent_is_internal:
            badges.insert(0, 'INTERNAL')

        badge_str = ""
        if badges:
            badge_str = "[" + "][".join(badges) + "] "

        port = node.get('port', 0)
        show_port = port and not node.get('is_composite_interface')
        port_info_str = f" [port {port}]" if show_port else ""

        label = _node_label(node)
        print(f"{prefix}{connector}{badge_str}{label}{port_info_str}")

        if node.get('children'):
            child_prefix = prefix + ("    " if is_last else "│   ")
            new_parent_int = _parent_is_internal or node.get('is_internal', False)
            _print_tree(node['children'], child_prefix, _show_internal, new_parent_int)


def _print_port_tree(port_node):
    """Print tree for a single port (children of the port node)."""
    children = port_node.get('children', [])
    if children:
        _print_tree(children, "    ")


def _print_verdict(v):
    """Print a single verdict line with hops/tiers/hubs."""
    status_char = '+' if v['color'] == 'green' else ('~' if v['color'] == 'orange' else '!')
    hubs_str = f"hubs {v['current_hubs']}/{v['max_hubs']}  " if 'current_hubs' in v else ""
    desc = v.get('description', v.get('name', ''))
    print(
        f"    {status_char} {desc:<22s} "
        f"{v['status']:<9s} "
        f"hops {v['current_hops']}/{v['max_hops']}  "
        f"tiers {v['current_tiers']}/{v['max_tiers']}  "
        f"{hubs_str}"
    )

def _print_stability_port(port_info, stability_data):
    """Print stability verdicts for a single port."""
    for v in stability_data:
        _print_verdict(v)


def _prompt_report_format() -> str:
    """Prompt user for report format: [Enter]HTML / [P]DF / [N]o report."""
    while True:
        choice = input("\nReport: [Enter]HTML / [P]DF / [N]o report: ").strip().upper()
        if choice in ('', 'H', 'HTML'):
            return 'html'
        elif choice in ('P', 'PDF'):
            return 'pdf'
        elif choice in ('N', 'NO', 'NONE'):
            return 'none'
        print("  Please press Enter for HTML, P for PDF, or N for no report")


def _prompt_monitor() -> bool:
    """Ask if user wants to start live monitoring."""
    while True:
        choice = input("\nStart live USB monitoring? [Y/n]: ").strip().lower()
        if choice in ('', 'y', 'yes'):
            return True
        elif choice in ('n', 'no'):
            return False
        print("  Please enter Y or N")


def main(csv_path=None):
    _print_header()

    # 1. Platform information
    platform_info = PlatformUtils.get_platform_info()
    print(f"Platform: {platform_info['os']} {platform_info['version']}")
    print(f"Architecture: {platform_info['architecture']}")
    if platform_info['is_apple_silicon']:
        print("Apple Silicon: Yes")
    _print_separator()

    # 2. USB analysis
    usb_analyzer = USBAnalyzer(csv_path) if csv_path else USBAnalyzer()
    _print_separator()

    print("Scanning USB devices...")
    usb_tree = usb_analyzer.build_tree()
    hops_data = usb_analyzer.calculate_hops_and_tiers(usb_tree)

    # 3. Stability
    stability = usb_analyzer.assess_stability(hops_data, usb_tree)

    # 4. Display information
    print("\nScanning displays...")
    display_analyzer = DisplayAnalyzer()
    displays = display_analyzer.get_display_info()

    # 5. Overall stability + Full tree + Per-port sections
    overall = stability.get('overall_worst', 'STABLE')
    mh = stability.get('max_hops', 0)
    mt = stability.get('max_tiers', 0)
    mhub = stability.get('max_hubs', 0)
    total = stability.get('total_endpoints', 0)
    ep_label = "endpoint" if total == 1 else "endpoints"
    print(f"Overall: {overall} ({total} {ep_label}, hops={mh}, tiers={mt}, hubs={mhub})")
    print()

    # Save original root children for per-port matching
    root_orig = usb_tree[0] if usb_tree else {}
    orig_children = list(root_orig.get('children', []))

    print("  Full USB & Display Tree")
    _print_tag("overall.tree")
    if usb_tree:
        root_node = usb_tree[0]

        # Add displays directly into tree (not under a "Displays" parent)
        if displays:
            for d in displays:
                prim = " (Primary)" if d.get('is_primary', False) else ""
                int_disp = d.get('is_internal', False)
                root_node.setdefault('children', []).append({
                    'model': f"{d['resolution']}  {d['name']}{prim}",
                    'name': d['name'], 'children': [], 'hops': 1,
                    'is_hub': False, 'is_internal': int_disp, 'is_display': True, 'port': 0
                })

        _print_tree(usb_tree, "  ", _show_internal=True)
    else:
        print("  No USB devices found.")
    print()

    print("  Overall rating")
    _print_tag("overall.verdict")
    for v in stability.get('verdicts', []):
        _print_verdict(v)
    
    # Print overall warnings if any
    overall_warnings = [v for v in stability.get('verdicts', []) if v.get('warning')]
    if overall_warnings:
        _print_tag("overall.warnings")
        for w in overall_warnings:
            print(f"    ! {w['name']}: {w['warning']}")
    print()
    print("PER PORT" + "=" * 31)
    print()

    ports_data = stability.get('ports', [])

    def _print_port_child(child, idx):
        port_info = next((p for p in ports_data if p.get('id') == idx + 1), None)
        label = port_info['label'] if port_info else child.get('model', 'Port')
        dc = len(port_info['devices']) if port_info else 0
        ph = port_info['max_hops'] if port_info else 0
        pt = port_info['max_tiers'] if port_info else 0
        p_hub = port_info.get('external_hubs', 0) if port_info else 0
        is_int = child.get('is_internal', False)
        int_pre = "[INTERNAL] " if is_int else ""
        ep_label = "endpoint" if dc == 1 else "endpoints"
        print(f"  {int_pre}{label} ({dc} {ep_label}, hops={ph}, tiers={pt}, hubs={p_hub})")
        _print_port_tree(child)
        if is_int:
            pass

    sep = "  " + "- " * 35

    def print_section(header, is_internal_filter, tag_prefix):
        print(header + "-" * 31)
        _print_tag(f"{tag_prefix}.section")
        first = True
        for idx, child in enumerate(orig_children):
            if child.get('is_display'):
                continue
            if is_internal_filter(child) != True:
                continue
            if not first:
                print()
            
            port_tag = f"{tag_prefix}.port{idx + 1}"
            _print_tag(f"{port_tag}.tree")
            _print_port_child(child, idx)
            print()
            
            port_info = next((p for p in ports_data if p.get('id') == idx + 1), None)
            if port_info and tag_prefix != "internal":
                _print_tag(f"{port_tag}.verdict")
                _print_stability_port(port_info, port_info['verdicts'])
                
                port_warnings = [v for v in port_info['verdicts'] if v.get('warning')]
                if port_warnings:
                    _print_tag(f"{port_tag}.warnings")
                    for w in port_warnings:
                        print(f"    ! {w['name']}: {w['warning']}")
            
            print()
            print(sep)
            first = False

    print_section("EXTERNAL", lambda c: not c.get('is_internal', False), "external")
    print_section("INTERNAL", lambda c: c.get('is_internal', False), "internal")

    # Live monitoring
    log_manager = LogManager()
    monitoring_logs = []
    unstable_devices = set()  # Initialize to empty set
    
    if _prompt_monitor():
        # Set up real-time log printing
        def print_log_entry(entry):
            print(entry)
        log_manager.set_print_callback(print_log_entry)
        
        log_manager.start_monitoring(usb_analyzer)
        
        print("\n" + "=" * 70)
        print("LIVE MONITORING - Press Enter to stop and generate report")
        print("=" * 70)
        print("Logs appear in real-time below:")
        print()
        
        # Wait for Enter key
        input()
        
        log_manager.stop_monitoring(usb_analyzer)
        
        # Clear the callback to prevent double printing
        log_manager.set_print_callback(None)
        
        monitoring_logs = log_manager.get_logs()
        unstable_devices = log_manager.get_unstable_devices()
        
        print("\n" + "=" * 70)
        print("MONITORING STOPPED - Logs captured")
        if unstable_devices:
            print("UNSTABLE DEVICES DETECTED (reconnected during monitoring):")
            for dev in sorted(unstable_devices):
                print(f"  ! {dev}")
        print("=" * 70)
        print()

    # 6. Room name
    room_name = input("Room name (optional, press Enter to skip): ").strip()

    # 7. Report generation - ask for format (always print all ports)
    format_type = _prompt_report_format()
    
    if format_type == 'none':
        _print_separator()
        print("Done!")
        print("=" * 70)
        return
    
    selected_ports = None  # None means all ports
    
    print("\n  Generating report...")
    report_gen = ReportGenerator(room_name=room_name)

    if format_type == 'html':
        html_path = report_gen.generate_html_report(
            usb_tree,
            hops_data,
            stability,
            displays,
            platform_info,
            platform_notes=usb_analyzer.get_platform_notes(),
            selected_ports=selected_ports,
            monitoring_logs=monitoring_logs,
            unstable_devices=unstable_devices
        )
        print(f"  HTML Report: {html_path}")
        report_gen.open_report(html_path)
    else:
        pdf_path = report_gen.generate_pdf_report(
            usb_tree=usb_tree, hops_data=hops_data,
            stability=stability, displays=displays,
            platform_info=platform_info,
            platform_notes=usb_analyzer.get_platform_notes(),
            monitoring_logs=monitoring_logs,
            unstable_devices=unstable_devices
        )
        if pdf_path:
            print(f"  PDF Report:  {pdf_path}")
            report_gen.open_report(pdf_path)

    _print_separator()
    print("Done!")
    print("=" * 70)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nInterrupted by user.")
        sys.exit(0)
    except Exception as e:
        print(f"\nAn error occurred: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)