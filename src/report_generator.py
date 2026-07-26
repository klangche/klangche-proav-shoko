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


class ReportGenerator:
    def __init__(self, output_dir: str = None):
        if output_dir is None:
            self.output_dir = Path(tempfile.gettempdir())
        else:
            self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        self.timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
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
                .replace("'", '''))

    def _format_verdict(self, v):
        status_char = '+' if v['color'] == 'green' else ('~' if v['color'] == 'orange' else '!')
        return ("    {0} {1:<20} "
                "{2:<9} "
                "hops {3}/{4}  "
                "tiers {5}/{6}").format(
                    v['name'], v['status'],
                    v['current_hops'], v['max_hops'],
                    v['current_tiers'], v['max_tiers'])

    def _build_tree_html(self, nodes, prefix="", is_last_list=None):
        if is_last_list is None:
            is_last_list = [True] * len(nodes)
        
        lines = []
        for i, node in enumerate(nodes):
            is_last = is_last_list[i]
            connector = "└── " if is_last else "├── "
            
            model = node.get('model', node.get('name', 'Unknown'))
            
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
            port_str = " [port {0}]".format(port) if port else ""
            
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

    def _format_verdict(self, v):
        status_char = '+' if v['color'] == 'green' else ('~' if v['color'] == 'orange' else '!')
        return ("    {0} {1:<20} "
                "{1:<9} "
                "hops {2}/{3}  "
                "tiers {4}/{5}").format(
                    status_char,
                    v['name'], v['status'],
                    v['current_hops'], v['max_hops'],
                    v['current_tiers'], v['max_tiers'])

    def generate_html_report(self, usb_tree, hops_data, stability, displays, platform_info,
                             platform_notes=None, custom_path=None, selected_ports=None,
                             monitoring_logs=None, unstable_devices=None):
        html = self._build_html(usb_tree, hops_data, stability, displays, platform_info,
                                platform_notes, monitoring_logs, unstable_devices)
        fn = Path(custom_path) if custom_path else self.output_dir / "proav-shoko_report_{0}.html".format(self.timestamp)
        fn.parent.mkdir(parents=True, exist_ok=True)
        with open(fn, 'w', encoding='utf-8') as f:
            f.write(html)
        return str(fn)

    def _build_html(self, usb_tree, hops_data, stability, displays, platform_info,
                    platform_notes, monitoring_logs, unstable_devices):
        
        # Build tree
        tree_lines = self._build_tree_html(usb_tree) if usb_tree else ["No USB devices found."]
        tree_html = "<br>".join(tree_lines)
        
        # Overall rating
        overall = stability.get('overall_worst', 'STABLE')
        mh = stability.get('max_hops', 0)
        mt = stability.get('max_tiers', 0)
        overall_html = "Overall: {0} ({1} hops, {2} tiers)".format(overall, mh, mt)
        overall_lines = []
        for v in stability.get('verdicts', []):
            overall_lines.append(self._format_verdict(v))
        overall_html = "<br>".join([overall_html] + overall_lines)
        
        # Ports data
        ports_data = stability.get('ports', [])
        orig_children = usb_tree[0].get('children', []) if usb_tree else []
        
        # EXTERNAL section
        ext_lines = ["-------------------------------EXTERNAL-------------------------------"]
        for idx, child in enumerate(usb_tree[0].get('children', []) if usb_tree else []):
            if child.get('is_display') or child.get('is_internal', False):
                continue
            port_info = next((p for p in stability.get('ports', []) if p.get('id') == idx + 1), None)
            label = port_info['label'] if port_info else child.get('model', 'Port')
            dc = len(port_info['devices']) if port_info else 0
            ph = port_info['max_hops'] if port_info else 0
            pt = port_info['max_tiers'] if port_info else 0
            
            ext_lines.append("  {0} ({1} devices, {2} hops, {4} tiers)".format(label, dc, ph, pt))
            
            for line in self._build_port_tree_html(usb_tree[0].get('children', [])[idx]):
                ext_lines.append("    " + line)
            
            if port_info:
                for v in port_info['verdicts']:
                    ext_lines.append(self._format_verdict(v))
            ext_lines.append("  " + "- " * 35)
        
        # INTERNAL section
        int_lines = ["-------------------------------INTERNAL-------------------------------"]
        for idx, child in enumerate(usb_tree[0].get('children', []) if usb_tree else []):
            if child.get('is_display') or not child.get('is_internal', False):
                continue
            port_info = next((p for p in stability.get('ports', []) if p.get('id') == idx + 1), None)
            label = port_info['label'] if port_info else child.get('model', 'Port')
            dc = len(port_info['devices']) if port_info else 0
            ph = port_info['max_hops'] if port_info else 0
            pt = port_info['max_tiers'] if port_info else 0
            
            int_lines.append("  [INTERNAL] {0} ({1} devices, {2} hops, {4} tiers)".format(label, dc, ph, pt))
            
            for line in self._build_port_tree_html(usb_tree[0].get('children', [])[idx]):
                int_lines.append("    " + line)
            
            int_lines.append("    (internal)")
            int_lines.append("  " + "- " * 35)
        
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
        css = self._load_css()
        
        # Prepare pre-formatted HTML strings
        ext_lines_html = "<br>".join(self._escape(line) for line in ext_lines)
        int_lines_html = "<br>".join(self._escape(line) for line in int_lines)
        disp_lines_html = "<br>".join(self._escape(line) for line in disp_lines)
        tree_html = "<br>".join(self._escape(line) for line in self._build_tree_html(usb_tree)) if usb_tree else "No USB devices found."
        
        overall = stability.get('overall_worst', 'STABLE')
        mh = stability.get('max_hops', 0)
        mt = stability.get('max_tiers', 0)
        overall_html = "Overall: {0} ({1} hops, {2} tiers)".format(overall, mh, mt)
        overall_lines = []
        for v in stability.get('verdicts', []):
            overall_lines.append(self._format_verdict(v))
        overall_html = "<br>".join([overall_html] + overall_lines)
        
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ap_html = '<span class="tag apple">Apple Silicon</span>' if platform_info.get('is_apple_silicon') else ''
        css = self._load_css()
        
        ext_lines_html = "<br>".join(self._escape(line) for line in ext_lines)
        int_lines_html = "<br>".join(self._escape(line) for line in int_lines)
        disp_lines_html = "<br>".join(self._escape(line) for line in disp_lines)
        tree_html = "<br>".join(self._escape(line) for line in self._build_tree_html(usb_tree)) if usb_tree else "No USB devices found."
        
        overall = stability.get('overall_worst', 'STABLE')
        mh = stability.get('max_hops', 0)
        mt = stability.get('max_tiers', 0)
        overall_html = "Overall: {0} ({1} hops, {2} tiers)".format(overall, mh, mt)
        overall_lines = []
        for v in stability.get('verdicts', []):
            overall_lines.append(self._format_verdict(v))
        overall_html = "<br>".join([overall_html] + overall_lines)
        
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ap_html = '<span class="tag apple">Apple Silicon</span>' if platform_info.get('is_apple_silicon') else ''
        css = self._load_css()
        
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
<div class="subtitle">USB Analysis &middot; {ts}</div>
<div class="tags">
<span class="tag">{platform_info[os]} {platform_info[version]}</span>
<span class="tag">{platform_info[architecture]}</span>
{ap_html}
</div>
<div class="stats">
<div class="stat"><div class="stat-value">{mh}</div><div class="stat-label">Max Hops</div></div>
<div class="stat"><div class="stat-value">{mt}</div><div class="stat-label">Tiers</div></div>
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
            mh=stability.get('max_hops', 0),
            mt=stability.get('max_tiers', 0),
            len_displays=len(displays),
            tree_html="<br>".join(self._escape(line) for line in self._build_tree_html(usb_tree)) if usb_tree else "No USB devices found.",
            overall_html="<br>".join([overall_html] + [self._format_verdict(v) for v in stability.get('verdicts', [])]),
            ext_lines_html="<br>".join(self._escape(line) for line in ext_lines),
            int_lines_html="<br>".join(self._escape(line) for line in int_lines),
            disp_lines_html="<br>".join(self._escape(line) for line in disp_lines),
            notes_html=notes_html,
            monitoring_html=monitoring_html,
            unstable_html=unstable_html,
            ts=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))