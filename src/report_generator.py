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


_ROOM_NAME_ENABLED = True


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

    def generate_pdf_report(self, html_path: str = None, custom_path: str = None, usb_tree=None,
                            hops_data=None, stability=None, displays=None,
                            platform_info=None, platform_notes=None,
                            selected_ports=None, monitoring_logs=None,
                            unstable_devices=None) -> Optional[str]:
        """Generate a PDF report directly from data using built-in PDF generator."""
        return self._generate_text_pdf(custom_path, usb_tree, hops_data,
                                       stability, displays, platform_info,
                                       platform_notes, selected_ports,
                                       monitoring_logs, unstable_devices)

    def _generate_text_pdf(self, output_path, usb_tree, hops_data, stability, displays,
                           platform_info, platform_notes, selected_ports,
                           monitoring_logs, unstable_devices) -> Optional[str]:
        """Fallback: generate a minimal text-based PDF using only Python stdlib.
        Creates a single long page with Courier font (no external deps needed)."""
        if not output_path:
            output_path = self.output_dir / "proav-shoko_report_{0}.pdf".format(self.file_tag)
        output_path = Path(output_path)

        lines = []
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        lines.append("ProAV Shoko - USB Analysis Report")
        if _ROOM_NAME_ENABLED and self.room_name:
            lines.append("Room: {0}".format(self.room_name))
        lines.append("=" * 60)
        lines.append("Generated: {0}".format(ts))
        lines.append("")

        if platform_info:
            os_str = platform_info.get('os', '')
            ver = platform_info.get('version', '')
            arch = platform_info.get('architecture', '')
            ap = " Apple Silicon" if platform_info.get('is_apple_silicon') else ""
            lines.append("System: {0} {1} {2}{3}".format(os_str, ver, arch, ap))
            lines.append("")

        lines.append("FULL USB & DISPLAY TREE")
        lines.append("-" * 60)
        if usb_tree:
            for node in usb_tree:
                self._collect_text_tree(node, lines, "")
        else:
            lines.append("  No USB devices found.")
        lines.append("")

        lines.append("OVERALL RATING")
        lines.append("-" * 60)
        if stability:
            overall = stability.get('overall_worst', 'STABLE')
            mh = stability.get('max_hops', 0)
            mt = stability.get('max_tiers', 0)
            mhub = stability.get('max_hubs', 0)
            total = stability.get('total_endpoints',
                                  sum(len(p.get('devices', [])) for p in stability.get('ports', [])))
            ep_label = "endpoint" if total == 1 else "endpoints"
            lines.append("Overall: {0} ({1} {2}, hops={3}, tiers={4}, hubs={5})".format(
                overall, total, ep_label, mh, mt, mhub))
            for v in stability.get('verdicts', []):
                lines.append(self._format_verdict(v))
        lines.append("")

        ports_data = stability.get('ports', []) if stability else []
        for header, internal_check in [("EXTERNAL", False), ("INTERNAL", True)]:
            lines.append("PER PORT ({0})".format(header))
            lines.append("-" * 60)
            root_orig = usb_tree[0] if usb_tree else {}
            orig_children = [c for c in root_orig.get('children', []) if not c.get('is_display')]
            for idx, child in enumerate(orig_children):
                if internal_check != child.get('is_internal', False):
                    continue
                port_info = next((p for p in ports_data if p.get('id') == idx + 1), None)
                label = port_info['label'] if port_info else child.get('model', 'Port')
                dc = len(port_info['devices']) if port_info else 0
                ph = port_info['max_hops'] if port_info else 0
                pt = port_info['max_tiers'] if port_info else 0
                p_hub = port_info.get('external_hubs', 0) if port_info else 0
                int_pfx = "[INTERNAL] " if header == "INTERNAL" else ""
                ep_label = "endpoint" if dc == 1 else "endpoints"
                lines.append("  {0}{1} ({2} {3}, hops={4}, tiers={5}, hubs={6})".format(
                    int_pfx, label, dc, ep_label, ph, pt, p_hub))
                for line in self._build_port_tree_html(child):
                    lines.append("    " + line)
                if port_info and header == "EXTERNAL":
                    for v in port_info['verdicts']:
                        lines.append(self._format_verdict(v))
                lines.append("")
            lines.append("")

        lines.append("CONNECTED DISPLAYS")
        lines.append("-" * 60)
        if displays:
            for d in displays:
                primary = " (Primary)" if d.get('is_primary', False) else ""
                intern = "[INTERNAL] " if d.get('is_internal', False) else ""
                lines.append("  {0}[DISPLAY] {1}  {2}{3}".format(intern, d['resolution'], d['name'], primary))
        else:
            lines.append("  No displays found.")
        lines.append("")

        if platform_notes:
            lines.append("PLATFORM NOTES")
            lines.append("-" * 60)
            for n in platform_notes:
                p = n.get('platform', '').replace('_', ' ').title()
                lines.append("  {0}: {1} -- {2}".format(p, n.get("description", ""), n.get("note", "")))
            lines.append("")

        if monitoring_logs:
            lines.append("MONITORING LOG")
            lines.append("-" * 60)
            for log in monitoring_logs:
                lines.append("  {0}".format(log))
            lines.append("")

        if unstable_devices:
            lines.append("[!] UNSTABLE DEVICES DETECTED")
            lines.append("-" * 60)
            for dev in unstable_devices:
                lines.append("  [!] {0} - Reconnected during monitoring (UNSTABLE)".format(dev))
            lines.append("")

        lines.append("=" * 60)
        lines.append("ProAV Shoko - Generated {0}".format(ts))

        text = "\n".join(lines)
        return self._write_minimal_pdf(text, output_path)

    def _collect_text_tree(self, nodes, lines, prefix, is_last_list=None):
        if isinstance(nodes, dict):
            nodes = [nodes]
        if is_last_list is None:
            is_last_list = [True] * len(nodes)
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
            lines.append("{0}{1}{2}{3}{4}".format(prefix, connector, badge_str, model, port_str))
            children = node.get('children', [])
            if children:
                child_prefix = prefix + ("    " if is_last else "│   ")
                child_is_last = [j == len(children) - 1 for j in range(len(children))]
                self._collect_text_tree(children, lines, child_prefix, child_is_last)

    @staticmethod
    def _write_minimal_pdf(text: str, output_path: Path) -> str:
        """Write a minimal valid PDF with text content, no external deps.
        Uses Courier (standard PDF Type1 font) on a single long page."""
        text = text.encode('ascii', errors='replace').decode('ascii')
        lines = text.split('\n')
        font_size = 10
        leading = 14
        margin_left = 50
        margin_top = 50
        page_width = 612
        line_height = leading
        content_height = len(lines) * line_height + margin_top + 50
        page_height = max(792, content_height)

        content_lines = []
        y = page_height - margin_top
        for line in lines:
            esc = line.replace('\\', '\\\\').replace('(', '\\(').replace(')', '\\)')
            content_lines.append("BT /F1 {0} Tf 1 0 0 1 {1} {2} Tm ({3}) Tj ET".format(
                font_size, margin_left, y, esc))
            y -= line_height

        stream_data = "\n".join(content_lines)
        stream_length = len(stream_data.encode('latin-1'))

        objects = []
        objects.append("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj")
        objects.append("2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj")
        objects.append((
            "3 0 obj\n"
            "<< /Type /Page /Parent 2 0 R\n"
            "   /MediaBox [0 0 {0} {1}]\n".format(page_width, page_height) +
            "   /Contents 4 0 R\n"
            "   /Resources << /Font << /F1 5 0 R >> >>\n"
            ">>\nendobj"
        ))
        objects.append((
            "4 0 obj\n"
            "<< /Length {0} >>\n".format(stream_length) +
            "stream\n{0}\nendstream\nendobj".format(stream_data)
        ))
        objects.append("5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj")

        body = "\n".join(objects)
        byte_offsets = []
        output = b"%PDF-1.4\n"
        byte_offsets.append(len(output))

        preamble = b"%\xFF\xFF\xFF\xFF\n"
        output += preamble
        byte_offsets.append(len(output))

        obj_num = 1
        for obj_text in objects:
            obj_bytes = obj_text.encode('latin-1') + b"\n"
            output += obj_bytes
            byte_offsets.append(len(output))
            obj_num += 1

        xref_offset = len(output)
        xref = "xref\n0 {0}\n".format(len(byte_offsets)).encode('ascii')
        xref += b"0000000000 65535 f \n"
        for offset in byte_offsets[:-1]:
            xref += "{0:010d} 00000 n \n".format(offset).encode('ascii')

        output += xref
        trailer = "trailer\n<< /Size {0} /Root 1 0 R >>\n".format(len(byte_offsets))
        output += trailer.encode('ascii')
        output += "startxref\n{0}\n%%EOF".format(xref_offset).encode('ascii')

        output_path.parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, 'wb') as f:
            f.write(output)
        return str(output_path)

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
@page {{ size: auto; margin: 0mm; }}
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