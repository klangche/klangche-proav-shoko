"""
Report generator - creates HTML/PDF reports matching CLI output style
"""

import os
import webbrowser
import subprocess
import sys
import tempfile
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Any, Optional


# Room name feature (disabled by default). Set to True to activate.
_ROOM_NAME_ENABLED = False


class ReportGenerator:
    def __init__(self, output_dir: str = None, room_name: str = ""):
        if output_dir is None:
            self.output_dir = Path(tempfile.gettempdir())
        else:
            self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        self.room_name = room_name
        if _ROOM_NAME_ENABLED and self.room_name:
            self.timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.file_tag = f"{self.room_name}_{self.timestamp}"
        else:
            self.timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.file_tag = self.timestamp
        self.css_path = Path(__file__).parent / "assets" / "report.css"

    def _load_css(self) -> str:
        if self.css_path.exists():
            with open(self.css_path, 'r', encoding='utf-8') as f:
                return f.read()
        return ""

    def _escape(self, text: str) -> str:
        return (text
                .replace('&', '&')
                .replace('<', '<')
                .replace('>', '>')
                .replace('"', '"')
                .replace("'", "'"))

    def _format_verdict(self, v):
        status_char = '+' if v['color'] == 'green' else ('~' if v['color'] == 'orange' else '!')
        hubs_str = "hubs {0}/{1}  ".format(v.get('current_hubs', 0), v.get('max_hubs', 0)) if 'current_hubs' in v else ""
        desc = v.get('description', v.get('name', ''))
        return ("    {sc} {desc:<22} "
                "{status:<9} "
                "hops {ch}/{mh}  "
                "tiers {ct}/{mt}  "
                "{hubs}").format(
                    sc=status_char,
                    desc=desc,
                    status=v['status'],
                    ch=v['current_hops'], mh=v['max_hops'],
                    ct=v['current_tiers'], mt=v['max_tiers'],
                    hubs=hubs_str)

    @staticmethod
    def _node_label(node):
        """Build a display label with interface type, model name and VID:PID."""
        model = node.get('model', node.get('name', 'Unknown'))
        device_info = node.get('device_info', '')
        iface_desc = node.get('interface_desc', '')
        iface_num = node.get('interface_number')
        if node.get('is_composite_interface'):
            mi = "MI_{0:02d}".format(iface_num) if iface_num is not None else ""
            suffix = " ({0})".format(device_info) if device_info else ""
            if model and 'USB-enhet' not in model and 'sammansatt' not in model and 'Composite' not in model:
                label = model
                if mi:
                    label += " " + mi
                return label + suffix
            if iface_desc:
                iface_tag = "HID Keyboard" if "Keyboard" in iface_desc else \
                            "HID Mouse" if "Mouse" in iface_desc else \
                            iface_desc
                return "{0} {1}{2}".format(iface_tag, mi, suffix).strip()
        return "{0} ({1})".format(model, device_info) if device_info else model

    def _build_tree_html(self, nodes, prefix="", is_last_list=None):
        if is_last_list is None:
            is_last_list = [True] * len(nodes)
        
        lines = []
        for i, node in enumerate(nodes):
            is_last = is_last_list[i]
            connector = "└── " if is_last else "├── "
            
            model = self._node_label(node)
            
            badges = []
            if node.get('is_hub'):
                badges.append('HUB')
            if node.get('is_display'):
                badges.append('DISPLAY')
            if node.get('is_internal', False):
                badges.insert(0, 'INTERNAL')
            
            badge_str = ""
            if badges:
                badge_str = "[" + "][".join(badges) + "] "
            
            port = node.get('port', 0)
            show_port = port and not node.get('is_composite_interface')
            port_str = " [port {0}]".format(port) if show_port else ""
            
            line = "{0}{1}{2}{3}{4}".format(prefix, connector, badge_str, model, port_str)
            lines.append(self._escape(line))
            
            children = node.get('children', [])
            if children:
                child_prefix = prefix + ("    " if is_last else "│   ")
                child_is_last = [j == len(children) - 1 for j in range(len(children))]
                lines.extend(self._build_tree_html(children, child_prefix, child_is_last))
        return lines

    def _build_port_tree_html(self, port_node):
        children = port_node.get('children', [])
        if not children:
            return []
        
        lines = []
        child_is_last = [j == len(port_node.get('children', [])) - 1 for j in range(len(port_node.get('children', [])))]
        lines.extend(self._build_tree_html(port_node.get('children', []), "    ", child_is_last))
        return lines

    def generate_html_report(self, usb_tree, hops_data, stability, displays, platform_info,
                             platform_notes=None, custom_path=None, selected_ports=None,
                             monitoring_logs=None, unstable_devices=None):
        html = self._build_html(usb_tree, hops_data, stability, displays, platform_info,
                                platform_notes, monitoring_logs, unstable_devices)
        fn = Path(custom_path) if custom_path else self.output_dir / "proav-shoko_report_{0}.html".format(self.file_tag)
        fn.parent.mkdir(parents=True, exist_ok=True)
        with open(fn, 'w', encoding='utf-8') as f:
            f.write(html)
        return str(fn)

    def generate_pdf_report(self, html_path: str, custom_path: str = None) -> Optional[str]:
        """Convert an HTML report to PDF using weasyprint."""
        try:
            from weasyprint import HTML
            pdf_path = Path(custom_path) if custom_path else html_path.replace('.html', '.pdf')
            HTML(filename=html_path).write_pdf(pdf_path)
            return str(pdf_path)
        except ImportError:
            print("  weasyprint is not installed. Install it with: pip install weasyprint")
            return None
        except Exception as e:
            print(f"  PDF generation failed: {e}")
            return None

    def open_report(self, path: str) -> None:
        """Open a report file in the default application."""
        try:
            if sys.platform == 'win32':
                os.startfile(path)
            elif sys.platform == 'darwin':
                subprocess.run(['open', path], check=False)
            else:
                subprocess.run(['xdg-open', path], check=False)
        except Exception as e:
            print(f"  Could not open report: {e}")

    def _build_html(self, usb_tree, hops_data, stability, displays, platform_info,
                    platform_notes, monitoring_logs, unstable_devices):
        
        # Build tree
        tree_html = "<br>".join(self._escape(line) for line in self._build_tree_html(usb_tree)) if usb_tree else "No USB devices found."
        
        # Overall rating
        overall = stability.get('overall_worst', 'STABLE')
        mh = stability.get('max_hops', 0)
        mt = stability.get('max_tiers', 0)
        mhub = stability.get('max_hubs', 0)
        total = stability.get('total_endpoints', sum(len(p.get('devices', [])) for p in stability.get('ports', [])))
        ep_label_out = "endpoint" if total == 1 else "endpoints"
        overall_html = "Overall: {0} ({1} {ep}, hops={2}, tiers={3}, hubs={4})".format(overall, total, mh, mt, mhub, ep=ep_label_out)
        overall_lines = []
        for v in stability.get('verdicts', []):
            overall_lines.append(self._format_verdict(v))
        overall_html_joined = "<br>".join([overall_html] + overall_lines)
        
        # Ports data
        ports_data = stability.get('ports', [])
        orig_children = usb_tree[0].get('children', []) if usb_tree else []
        
        def _build_port_lines(header, internal_filter):
            lines = ["-------------------------------{0}-------------------------------".format(header)]
            for idx, child in enumerate(usb_tree[0].get('children', []) if usb_tree else []):
                if child.get('is_display') or internal_filter(child):
                    continue
                port_info = next((p for p in stability.get('ports', []) if p.get('id') == idx + 1), None)
                label = port_info['label'] if port_info else child.get('model', 'Port')
                dc = len(port_info['devices']) if port_info else 0
                ph = port_info['max_hops'] if port_info else 0
                pt = port_info['max_tiers'] if port_info else 0
                p_hub = port_info.get('external_hubs', 0) if port_info else 0
                int_pfx = "[INTERNAL] " if header == "INTERNAL" else ""
                
                ep_label_out = "endpoint" if dc == 1 else "endpoints"
                lines.append("  {int_pfx}{label} ({dc} {ep}, hops={ph}, tiers={pt}, hubs={p_hub})".format(
                    int_pfx=int_pfx, label=label, dc=dc, ep=ep_label_out, ph=ph, pt=pt, p_hub=p_hub))
                
                for line in self._build_port_tree_html(usb_tree[0].get('children', [])[idx]):
                    lines.append("    " + line)
                
                if port_info and header != "INTERNAL":
                    for v in port_info['verdicts']:
                        lines.append(self._format_verdict(v))
                if header == "INTERNAL":
                    lines.append("    (internal)")
                lines.append("  " + "- " * 35)
            return lines
        
        ext_lines = _build_port_lines("EXTERNAL", lambda c: c.get('is_internal', False))
        int_lines = _build_port_lines("INTERNAL", lambda c: not c.get('is_internal', False))
        
        # Displays
        disp_lines = []
        if displays:
            for d in displays:
                primary = " (Primary)" if d.get('is_primary', False) else ""
                int_disp = "[INTERNAL] " if d.get('is_internal', False) else ""
                disp_lines.append("  {0}[DISPLAY] {1}  {2}{3}".format(int_disp, d['resolution'], d['name'], primary))
        else:
            disp_lines.append("  No displays found.")
        
        # Platform notes
        notes_html = ""
        if platform_notes:
            note_items = []
            for n in platform_notes:
                p = n.get('platform', '').replace('_', ' ').title()
                note_items.append('<div class="note"><span class="np">{0}</span> {1} &mdash; {2}</div>'.format(p, n.get("description", ""), n.get("note", "")))
            notes_html = '<div class="section"><div class="section-title">PLATFORM NOTES</div>{0}</div>'.format("".join(note_items))
        
        # Monitoring logs
        monitoring_html = ""
        if monitoring_logs:
            log_entries = "<br>".join(self._escape(log) for log in monitoring_logs)
            monitoring_html = '''
<div class="section">
    <div class="section-title">MONITORING LOG</div>
    <div class="log-content">{0}</div>
</div>'''.format(log_entries)
        
        # Unstable devices
        unstable_html = ""
        if unstable_devices:
            unstable_entries = "<br>".join('[!] {0} - Reconnected during monitoring (UNSTABLE)'.format(self._escape(device)) for device in unstable_devices)
            unstable_html = '''
<div class="section unstable">
    <div class="section-title">[!] UNSTABLE DEVICES DETECTED</div>
    <div class="log-content">{0}</div>
</div>'''.format(unstable_entries)
        
        css = self._load_css()
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ap_html = '<span class="tag apple">Apple Silicon</span>' if platform_info.get('is_apple_silicon') else ''
        room_header = ""
        if _ROOM_NAME_ENABLED and self.room_name:
            room_header = '<div class="room-name">Room: {0}</div>'.format(self._escape(self.room_name))
        
        ext_lines_html = "<br>".join(self._escape(line) for line in ext_lines)
        int_lines_html = "<br>".join(self._escape(line) for line in int_lines)
        disp_lines_html = "<br>".join(self._escape(line) for line in disp_lines)
        
        return """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>ProAV Shoko - USB Analysis</title>
<style>
{css}
body {{ font-family: 'Cascadia Code', 'Consolas', 'Fira Code', 'Courier New', monospace; background:#0c0c10; color:#d4d4d4; padding:24px; font-size:13px; line-height:1.5 }}
.wrapper {{ max-width: 960px; margin:0 auto; }}
.header {{ color:#569CD6; font-size:1.3rem; font-weight:600; border-bottom:1px solid #1a1a2e; padding-bottom:10px; margin-bottom:4px }}
.subtitle {{ color:#808080; font-size:0.75rem; margin-bottom:20px }}
.tags {{ display:flex; flex-wrap:wrap; gap:4px; margin-bottom:16px }}
.tag {{ background:#0a0a12; border:1px solid #1a1a2e; padding:2px 10px; border-radius:3px; font-size:0.7rem; color:#808080 }}
.tag.apple {{ background:#1a0a00; border-color:#D7BA7D; color:#D7BA7D }}
.stats {{ display:flex; gap:8px; margin:16px 0 }}
.stat {{ flex:1; min-width:80px; background:#0a0a12; border:1px solid #1a1a2e; padding:12px; text-align:center; border-radius:3px }}
.stat-value {{ font-size:1.3rem; font-weight:700; color:#569CD6 }}
.stat-label {{ color:#808080; font-size:0.6rem; text-transform:uppercase; letter-spacing:0.3px }}
.section {{ background:#0a0a12; border:1px solid #1a1a2e; border-radius:3px; padding:16px; margin:12px 0 }}
.section-title {{ color:#569CD6; font-size:1rem; font-weight:600; margin:24px 0 12px 0; padding-bottom:6px; border-bottom:1px solid #1a1a2e }}
.section.unstable {{ border-color:#D16969; }}
.log-content {{ white-space:pre-wrap; font-family:inherit; font-size:0.75rem }}
.note {{ color:#808080; font-size:0.65rem; margin:2px 0 }}
.np {{ color:#d4d4d4; font-weight:500 }}
.footer {{ margin-top:24px; padding-top:12px; border-top:1px solid #1a1a2e; color:#2a2a3e; font-size:0.6rem; text-align:center }}
.pre {{ white-space:pre; font-family:inherit; }}
</style>
</head>
<body>
<div class="wrapper">
<div class="header">ProAV Shoko</div>
<div class="subtitle">{room_header}USB Analysis &middot; {ts}</div>
<div class="tags">
<span class="tag">{platform_info[os]} {platform_info[version]}</span>
<span class="tag">{platform_info[architecture]}</span>
{ap_html}
</div>
<div class="stats">
<div class="stat"><div class="stat-value">{mh}</div><div class="stat-label">Max Hops</div></div>
<div class="stat"><div class="stat-value">{mt}</div><div class="stat-label">Tiers</div></div>
<div class="stat"><div class="stat-value">{mhub}</div><div class="stat-label">Hubs</div></div>
<div class="stat"><div class="stat-value">{len_displays}</div><div class="stat-label">Displays</div></div>
</div>

<div class="section">
<div class="section-title">FULL USB & DISPLAY TREE</div>
<pre class="pre">{tree_html}</pre>
</div>

<div class="section">
<div class="section-title">OVERALL RATING</div>
<pre class="pre">{overall_html}</pre>
</div>

<div class="section">
<div class="section-title">PER PORT</div>
<pre class="pre">{ext_lines_html}</pre>
</div>

<div class="section">
<div class="section-title">PER PORT (INTERNAL)</div>
<pre class="pre">{int_lines_html}</pre>
</div>

<div class="section">
<div class="section-title">CONNECTED DISPLAYS</div>
<pre class="pre">{disp_lines_html}</pre>
</div>

{notes_html}
{monitoring_html}
{unstable_html}

<div class="footer">ProAV Shoko v1.0.0 &middot; {ts}</div>
</div>
</body>
</html>""".format(
            css=self._load_css(),
            ts=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            platform_info=platform_info,
            ap_html='<span class="tag apple">Apple Silicon</span>' if platform_info.get('is_apple_silicon') else '',
            room_header=room_header,
            mh=stability.get('max_hops', 0),
            mt=stability.get('max_tiers', 0),
            mhub=stability.get('max_hubs', 0),
            len_displays=len(displays),
            tree_html=tree_html,
            overall_html=overall_html_joined,
            ext_lines_html=ext_lines_html,
            int_lines_html=int_lines_html,
            disp_lines_html=disp_lines_html,
            notes_html=notes_html,
            monitoring_html=monitoring_html,
            unstable_html=unstable_html)