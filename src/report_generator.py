"""
Rapportgenerator - skapar HTML- och PDF-rapporter
"""

import os
import webbrowser
import subprocess
import sys
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Any, Optional

try:
    from weasyprint import HTML
except ImportError:
    print("⚠️  weasyprint inte installerat. PDF-generering är inaktiverad.")
    HTML = None


class ReportGenerator:
    """Genererar HTML- och PDF-rapporter."""

    def __init__(self, output_dir: str = "reports"):
        """Initiera rapportgeneratorn."""
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        self.timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    def generate_html_report(
        self,
        usb_tree: List[Dict],
        hops_data: Dict[str, Any],
        stability: Dict[str, Any],
        displays: List[Dict[str, Any]],
        platform_info: Dict[str, Any],
        custom_path: Optional[str] = None
    ) -> str:
        """
        Genererar HTML-rapport med mörk bakgrund.

        Returns:
            Sökväg till HTML-filen.
        """
        html_content = self._build_html_content(
            usb_tree, hops_data, stability, displays, platform_info
        )

        if custom_path:
            filename = Path(custom_path)
        else:
            filename = self.output_dir / f"proav-shoko_report_{self.timestamp}.html"

        with open(filename, 'w', encoding='utf-8') as f:
            f.write(html_content)

        return str(filename)

    def _build_html_content(self, usb_tree, hops_data, stability, displays, platform_info) -> str:
        """Bygger HTML-innehållet."""
        usb_html = self._render_tree(usb_tree)
        display_html = self._render_displays(displays)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        return f"""
<!DOCTYPE html>
<html lang="sv">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ProAV Shōko - USB-analys</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: 'Menlo', 'Monaco', 'Courier New', monospace;
            background-color: #1a1a2e;
            color: #e0e0e0;
            padding: 20px;
            line-height: 1.6;
        }}
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background: #16213e;
            padding: 30px;
            border-radius: 12px;
            box-shadow: 0 8px 32px rgba(0,0,0,0.5);
        }}
        h1 {{
            color: #00d4ff;
            font-size: 2.2em;
            border-bottom: 2px solid #00d4ff;
            padding-bottom: 10px;
            margin-bottom: 20px;
        }}
        .subtitle {{ color: #aaa; font-size: 0.9em; margin-bottom: 30px; }}
        .card {{
            background: #1a1a2e;
            border-left: 4px solid #00d4ff;
            padding: 15px 20px;
            margin: 20px 0;
            border-radius: 6px;
        }}
        .card-green {{ border-left-color: #00cc66; }}
        .card-yellow {{ border-left-color: #ffcc00; }}
        .card-orange {{ border-left-color: #ff8800; }}
        .card-red {{ border-left-color: #ff3333; }}
        .tree {{
            font-family: 'Menlo', 'Monaco', 'Courier New', monospace;
            font-size: 0.9em;
            padding: 10px 0;
        }}
        .tree ul {{ list-style: none; padding-left: 20px; }}
        .tree li {{ padding: 3px 0; border-left: 2px dotted #444; padding-left: 15px; margin-left: 10px; }}
        .tree .hub {{ color: #ffaa00; font-weight: bold; }}
        .tree .device {{ color: #88ccff; }}
        .hops-badge {{
            background: #333;
            color: #fff;
            border-radius: 12px;
            padding: 0 10px;
            font-size: 0.8em;
            margin-left: 10px;
        }}
        .stats {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 15px;
            margin: 20px 0;
        }}
        .stat-item {{
            background: #1a1a2e;
            padding: 15px;
            border-radius: 8px;
            text-align: center;
            border: 1px solid #333;
        }}
        .stat-value {{ font-size: 2em; font-weight: bold; color: #00d4ff; }}
        .stat-label {{ color: #aaa; font-size: 0.8em; }}
        .display-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin: 10px 0;
        }}
        .display-item {{
            background: #1a1a2e;
            padding: 15px;
            border-radius: 8px;
            border: 1px solid #333;
            text-align: center;
        }}
        .display-item .resolution {{ font-size: 1.4em; color: #00d4ff; font-weight: bold; }}
        .stability-badge {{
            display: inline-block;
            padding: 10px 25px;
            border-radius: 30px;
            font-weight: bold;
            font-size: 1.2em;
        }}
        .badge-green {{ background: #00cc66; color: #000; }}
        .badge-yellow {{ background: #ffcc00; color: #000; }}
        .badge-orange {{ background: #ff8800; color: #000; }}
        .badge-red {{ background: #ff3333; color: #fff; }}
        .warning {{
            background: #ff3333;
            color: #fff;
            padding: 12px 20px;
            border-radius: 8px;
            margin: 15px 0;
            font-weight: bold;
        }}
        .warning-icon {{ font-size: 1.5em; margin-right: 10px; }}
        .footer {{
            margin-top: 40px;
            padding-top: 20px;
            border-top: 1px solid #333;
            color: #666;
            font-size: 0.8em;
            text-align: center;
        }}
        .platform-info {{
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            background: #1a1a2e;
            padding: 15px;
            border-radius: 8px;
            margin: 10px 0;
        }}
        .platform-tag {{
            background: #333;
            padding: 5px 15px;
            border-radius: 20px;
            font-size: 0.8em;
        }}
        .apple-silicon {{ background: #ff8800; color: #000; font-weight: bold; }}
    </style>
</head>
<body>
    <div class="container">
        <h1>🔍 ProAV Shōko</h1>
        <div class="subtitle">USB-analysrapport • {timestamp}</div>

        <div class="platform-info">
            <span class="platform-tag">🖥️ {platform_info['os']} {platform_info['version']}</span>
            <span class="platform-tag">🧠 {platform_info['architecture']}</span>
            {'''<span class="platform-tag apple-silicon">🍎 Apple Silicon</span>''' if platform_info['is_apple_silicon'] else ''}
        </div>

        <div class="card card-{stability['color']}">
            <h3>📊 Stabilitetsbedömning</h3>
            <div style="display: flex; align-items: center; gap: 20px; flex-wrap: wrap;">
                <span class="stability-badge badge-{stability['color']}">{stability['label']}</span>
                <span>Max hops: <strong>{hops_data['max_hops']}</strong> • Tiers: <strong>{hops_data['max_tiers']}</strong></span>
            </div>
            {f'''<div class="warning"><span class="warning-icon">⚠️</span> {stability['warning']}</div>''' if stability['warning'] else ''}
        </div>

        <div class="stats">
            <div class="stat-item">
                <div class="stat-value">{len(usb_tree)}</div>
                <div class="stat-label">USB-enheter (rot)</div>
            </div>
            <div class="stat-item">
                <div class="stat-value">{hops_data['max_hops']}</div>
                <div class="stat-label">Max hops</div>
            </div>
            <div class="stat-item">
                <div class="stat-value">{hops_data['max_tiers']}</div>
                <div class="stat-label">Tiers</div>
            </div>
            <div class="stat-item">
                <div class="stat-value">{len(displays)}</div>
                <div class="stat-label">Anslutna skärmar</div>
            </div>
        </div>

        <h2>🌳 USB-träd</h2>
        <div class="tree">{usb_html}</div>

        <h2 style="margin-top: 40px;">🖥️ Anslutna skärmar</h2>
        <div class="display-grid">{display_html}</div>

        <div class="footer">Genererad av ProAV Shōko v1.0.0 • {timestamp}</div>
    </div>
</body>
</html>
        """

    def _render_tree(self, tree: List[Dict], level: int = 0) -> str:
        """Rekursivt renderar USB-trädet som HTML."""
        if not tree:
            return "<p>Inga USB-enheter hittades.</p>"

        html = "<ul>"
        for node in tree:
            is_hub = node.get('is_hub', False)
            icon = "🔌" if is_hub else "🖥️"
            cls = "hub" if is_hub else "device"
            devpath = node.get('devpath', '')
            hops = devpath.count('/') if devpath else 0

            html += f"""
            <li>
                <span class="node-label">
                    <span class="{cls}">{icon} {node.get('model', node.get('name', 'Okänd'))}</span>
                    <span class="hops-badge">hops: {hops}</span>
                    {'''<span style="color:#ffaa00;font-size:0.8em;"> [HUB]</span>''' if is_hub else ''}
                </span>
            """

            if node.get('children'):
                html += self._render_tree(node['children'], level + 1)

            html += "</li>"

        html += "</ul>"
        return html

    def _render_displays(self, displays: List[Dict[str, Any]]) -> str:
        """Renderar skärminformation som HTML."""
        if not displays:
            return "<p>Inga skärmar hittades.</p>"

        html = ""
        for display in displays:
            primary_mark = " ⭐" if display.get('is_primary', False) else ""
            html += f"""
            <div class="display-item">
                <div class="resolution">{display['resolution']}</div>
                <div>{display['name']}{primary_mark}</div>
                <div style="color:#888;font-size:0.8em;">{display.get('width', 0)} × {display.get('height', 0)} px</div>
            </div>
            """
        return html

    def generate_pdf_report(self, html_path: str, custom_path: Optional[str] = None) -> Optional[str]:
        """
        Genererar PDF-rapport från HTML.

        Returns:
            Sökväg till PDF-filen eller None vid fel.
        """
        if HTML is None:
            print("⚠️  weasyprint inte installerat, hoppar över PDF-generering.")
            return None

        try:
            if custom_path:
                pdf_filename = Path(custom_path)
            else:
                pdf_filename = Path(html_path).with_suffix('.pdf')

            HTML(filename=html_path).write_pdf(str(pdf_filename))
            return str(pdf_filename)

        except Exception as e:
            print(f"⚠️  Kunde inte generera PDF: {e}")
            return None

    def open_report(self, file_path: str) -> None:
        """Öppnar rapport i standardprogram."""
        try:
            abs_path = os.path.abspath(file_path)

            if sys.platform == 'darwin':
                subprocess.run(['open', abs_path], check=False)
            elif sys.platform == 'win32':
                os.startfile(abs_path)
            else:
                webbrowser.open(f'file://{abs_path}')

            print(f"📂 Öppnade: {abs_path}")

        except Exception as e:
            print(f"⚠️  Kunde inte öppna filen: {e}")
