"""
Report generator - creates professional HTML and PDF reports
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
except (ImportError, OSError):
    HTML = None


class ReportGenerator:
    """Generates professional HTML and PDF reports."""

    def __init__(self, output_dir: str = "reports"):
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
        platform_notes: Optional[List[Dict[str, str]]] = None,
        custom_path: Optional[str] = None
    ) -> str:
        html_content = self._build_html_content(
            usb_tree, hops_data, stability, displays, platform_info, platform_notes
        )
        if custom_path:
            filename = Path(custom_path)
        else:
            filename = self.output_dir / f"proav-shoko_report_{self.timestamp}.html"
        filename.parent.mkdir(parents=True, exist_ok=True)
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(html_content)
        return str(filename)

    def _build_mermaid_tree(self, usb_tree: List[Dict]) -> str:
        lines = ["graph TD"]

        def add_node(node: Dict, parent_id: Optional[str] = None, depth: int = 0) -> str:
            node_id = f"n{node.get('devpath', '').replace('/', '_').replace('-', '_').replace('.', '_')}"
            if not node_id or node_id == "n":
                node_id = f"n{id(node)}"
            model = node.get('model', node.get('name', 'Unknown Device'))
            is_hub = node.get('is_hub', False)
            devpath = node.get('devpath', '')
            hops = devpath.count('/') if devpath else 0
            label = f"{model}{' [HUB]' if is_hub else ''} ({hops})"
            shape = ":::hub" if is_hub else ":::device"
            lines.append(f'    {node_id}["{label}"]{shape}')
            if parent_id:
                lines.append(f"    {parent_id} --> {node_id}")
            if node.get('children'):
                for child in node['children']:
                    add_node(child, node_id, depth + 1)
            return node_id

        for root in usb_tree:
            add_node(root)

        lines.append("    classDef hub fill:#1a1a2e,stroke:#ffaa00,stroke-width:2px,color:#ffaa00")
        lines.append("    classDef device fill:#1a1a2e,stroke:#00d4ff,stroke-width:1px,color:#e0e0e0")
        return "\n".join(lines)

    def _build_stability_html(self, stability_data: Dict[str, Any]) -> str:
        parts = []
        groups = stability_data.get('groups', {})

        for arch, verdicts in groups.items():
            rows = []
            for v in verdicts:
                color = v.get('color', 'green')
                status_class = "s-stable" if color == "green" else ("s-warn" if color == "orange" else "s-fail")
                rows.append(
                    f'<tr class="{status_class}">'
                    f'<td class="s-name">{v["emoji"]} {v["name"]}</td>'
                    f'<td class="s-status">{v["status"]}</td>'
                    f'<td class="s-metric">{v["current_hops"]}<span class="s-sep">/</span>{v["max_hops"]}</td>'
                    f'<td class="s-metric">{v["current_tiers"]}<span class="s-sep">/</span>{v["max_tiers"]}</td>'
                    f'</tr>'
                )
            parts.append(
                f'<div class="s-arch">'
                f'<div class="s-arch-name">{arch}</div>'
                f'<table class="s-table"><tbody>{"".join(rows)}</tbody></table>'
                f'</div>'
            )

        warnings = [v for v in stability_data.get('verdicts', []) if v.get('warning')]
        if warnings:
            warn_rows = []
            for w in warnings:
                warn_rows.append(f'<div class="w-item">[{w["name"]}] {w["warning"]} (hops: {w["current_hops"]})</div>')
            parts.append(f'<div class="s-warnings">{"".join(warn_rows)}</div>')

        return "".join(parts)

    def _build_html_content(
        self,
        usb_tree: List[Dict],
        hops_data: Dict[str, Any],
        stability: Dict[str, Any],
        displays: List[Dict[str, Any]],
        platform_info: Dict[str, Any],
        platform_notes: Optional[List[Dict[str, str]]] = None
    ) -> str:
        mermaid_tree = self._build_mermaid_tree(usb_tree)
        stability_html = self._build_stability_html(stability)
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        hub_count = sum(1 for d in usb_tree if d.get('is_hub', False))

        apple_tag = ""
        if platform_info.get('is_apple_silicon'):
            apple_tag = '<span class="tag tag-apple">Apple Silicon</span>'

        displays_html = ""
        if displays:
            items = []
            for d in displays:
                prim = " &middot; Primary" if d.get('is_primary', False) else ""
                items.append(
                    f'<div class="disp-item">'
                    f'<div class="disp-res">{d["resolution"]}</div>'
                    f'<div class="disp-name">{d["name"]}{prim}</div>'
                    f'</div>'
                )
            displays_html = f'<div class="disp-grid">{"".join(items)}</div>'
        else:
            displays_html = '<p class="muted">No displays detected.</p>'

        notes_html = ""
        if platform_notes:
            ns = []
            for n in platform_notes:
                p = n.get('platform', '').replace('_', ' ').title()
                ns.append(
                    f'<div class="note"><span class="note-p">{p}</span> '
                    f'<span class="note-d">{n.get("description", "")}</span> &mdash; {n.get("note", "")}</div>'
                )
            notes_html = f'<div class="notes">{"".join(ns)}</div>'

        return f"""<!DOCTYPE html><html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>ProAV Shoko &mdash; USB Analysis</title>
<script src="https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js"></script>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:'Segoe UI',-apple-system,BlinkMacSystemFont,sans-serif;background:#0d1117;color:#e6edf3;padding:24px;line-height:1.5;font-size:14px}}
.wrap{{max-width:960px;margin:0 auto}}
h1{{font-size:1.5rem;font-weight:600;color:#f0f6fc;border-bottom:1px solid #21262d;padding-bottom:10px;margin-bottom:4px}}
.sub{{color:#8b949e;font-size:0.8rem;margin-bottom:20px}}
h2{{color:#f0f6fc;font-size:1.1rem;font-weight:600;margin:24px 0 12px 0;padding-bottom:6px;border-bottom:1px solid #21262d}}

.tags{{display:flex;flex-wrap:wrap;gap:4px;margin-bottom:16px}}
.tag{{background:#161b22;border:1px solid #30363d;padding:2px 10px;border-radius:4px;font-size:0.7rem;color:#8b949e}}
.tag-apple{{background:#3d2e1f;border-color:#ff8800;color:#ffaa00;font-weight:500}}

.summary{{display:flex;gap:8px;margin:16px 0}}
.sum-item{{flex:1;min-width:80px;background:#161b22;border:1px solid #21262d;padding:12px;text-align:center;border-radius:6px}}
.sum-val{{font-size:1.4rem;font-weight:700;color:#58a6ff}}
.sum-lbl{{color:#8b949e;font-size:0.65rem;text-transform:uppercase;letter-spacing:0.4px}}

.card{{background:#161b22;border:1px solid #21262d;border-radius:6px;padding:16px;margin:12px 0}}

/* Stability table */
.s-arch{{margin-bottom:10px}}
.s-arch-name{{color:#8b949e;font-size:0.7rem;text-transform:uppercase;letter-spacing:0.4px;margin-bottom:4px;font-weight:600}}
.s-table{{width:100%;border-collapse:collapse;font-size:0.8rem}}
.s-table td{{padding:4px 8px;border-bottom:1px solid #21262d}}
.s-name{{width:40%}}
.s-status{{width:18%;font-size:0.65rem;font-weight:600}}
.s-metric{{width:16%;text-align:right;color:#8b949e;font-variant-numeric:tabular-nums}}
.s-sep{{color:#30363d;margin:0 2px}}
.s-stable td.s-status{{color:#00cc66}}
.s-warn td.s-status{{color:#ffaa00}}
.s-fail td.s-status{{color:#ff3333}}
.s-table tr:last-child td{{border-bottom:none}}
.s-warnings{{margin-top:8px;padding:8px 12px;background:#1a0a0a;border:1px solid #331111;border-radius:4px}}
.w-item{{color:#ff7777;font-size:0.75rem;margin:2px 0}}

/* Mermaid */
.mermaid-box{{background:#0d1117;border:1px solid #21262d;border-radius:6px;padding:12px;overflow-x:auto}}

/* Displays */
.disp-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:8px}}
.disp-item{{background:#0d1117;border:1px solid #21262d;padding:12px;text-align:center;border-radius:6px}}
.disp-res{{font-size:1.2rem;font-weight:600;color:#58a6ff}}
.disp-name{{color:#e6edf3;font-size:0.8rem}}

.notes{{margin-top:16px;padding:12px;background:#0d1117;border:1px solid #21262d;border-radius:6px}}
.note{{color:#8b949e;font-size:0.7rem;margin:3px 0}}
.note-p{{color:#e6edf3;font-weight:500}}

.footer{{margin-top:24px;padding-top:12px;border-top:1px solid #21262d;color:#484f58;font-size:0.65rem;text-align:center}}
.muted{{color:#8b949e;font-size:0.8rem}}

@media (max-width:600px){{.summary{{flex-wrap:wrap}},.s-table,.s-table tbody,.s-table tr,.s-table td{{display:block;width:100%}}.s-table td{{border:none;padding:2px 8px}}.s-table tr{{border-bottom:1px solid #21262d;padding:4px 0}}}}
</style>
</head>
<body>
<div class="wrap">
<h1>ProAV Shoko</h1>
<div class="sub">USB Analysis &middot; {ts}</div>

<div class="tags">
<span class="tag">{platform_info['os']} {platform_info['version']}</span>
<span class="tag">{platform_info['architecture']}</span>
{apple_tag}
</div>

<div class="summary">
<div class="sum-item"><div class="sum-val">{len(usb_tree)}</div><div class="sum-lbl">Devices</div></div>
<div class="sum-item"><div class="sum-val">{hops_data['max_hops']}</div><div class="sum-lbl">Max Hops</div></div>
<div class="sum-item"><div class="sum-val">{hops_data['max_tiers']}</div><div class="sum-lbl">Tiers</div></div>
<div class="sum-item"><div class="sum-val">{hub_count}</div><div class="sum-lbl">Hubs</div></div>
<div class="sum-item"><div class="sum-val">{len(displays)}</div><div class="sum-lbl">Displays</div></div>
</div>

<h2>Stability Assessment</h2>
<div class="card">{stability_html}</div>

<h2>USB Tree</h2>
<div class="card"><div class="mermaid-box"><pre class="mermaid">{mermaid_tree}</pre></div></div>

<h2>Connected Displays</h2>
<div class="card">{displays_html}</div>

{notes_html}

<div class="footer">ProAV Shoko v1.0.0 &middot; {ts} &middot; hop_limits.csv from src/assets/</div>
</div>
<script>
mermaid.initialize({{
    startOnLoad:true,
    theme:'dark',
    themeVariables:{{
        primaryColor:'#161b22', primaryTextColor:'#e6edf3',
        primaryBorderColor:'#30363d', lineColor:'#484f58',
        secondaryColor:'#161b22', tertiaryColor:'#0d1117',
        background:'#0d1117', mainBkg:'#0d1117',
        secondBkg:'#161b22', tertiaryBkg:'#21262d'
    }},
    flowchart:{{useMaxWidth:true,htmlLabels:true,curve:'basis'}}
}});
</script>
</body>
</html>"""

    def generate_pdf_report(self, html_path: str, custom_path: Optional[str] = None) -> Optional[str]:
        if HTML is None:
            return None
        try:
            pdf_path = Path(custom_path) if custom_path else Path(html_path).with_suffix('.pdf')
            pdf_path.parent.mkdir(parents=True, exist_ok=True)
            HTML(filename=html_path).write_pdf(str(pdf_path))
            return str(pdf_path)
        except Exception:
            return None

    def open_report(self, file_path: str) -> None:
        try:
            abs_path = os.path.abspath(file_path)
            if sys.platform == 'darwin':
                subprocess.run(['open', abs_path], check=False)
            elif sys.platform == 'win32':
                os.startfile(abs_path)
            else:
                webbrowser.open(f'file://{abs_path}')
        except Exception:
            pass