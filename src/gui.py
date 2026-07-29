#!/usr/bin/env python3
"""
ProAV Shoko - USB Detective GUI

Clean, simple GUI that shows USB tree and live logs.
Matches CLI format exactly with modern look.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import queue
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

from src.usb_analyzer import USBAnalyzer
from src.display_analyzer import DisplayAnalyzer
from src.report_generator import ReportGenerator
from src.platform_utils import PlatformUtils

# Below this window width (px), the live log panel is hidden and the
# tree takes the full width. Above it, tree:log keeps roughly a 70/30 split.
NARROW_WINDOW_THRESHOLD = 900


class LogRedirector:
    """Redirects stdout/stderr to the GUI log."""

    def __init__(self, text_widget, queue):
        self.text_widget = text_widget
        self.queue = queue
        self.original_stdout = sys.stdout

    def write(self, text):
        """Write to both the original stdout and the GUI."""
        self.original_stdout.write(text)
        self.queue.put(text)

    def flush(self):
        self.original_stdout.flush()


class ProAVShokoGUI:
    """Main GUI for ProAV Shoko."""

    def __init__(self, root, csv_path=None):
        self.root = root
        self.root.title("ProAV Shoko - USB Detective")
        
        # Set window to center screen
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = 1400
        window_height = 800
        x_position = (screen_width - window_width) // 2
        y_position = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
        self.root.minsize(700, 500)

        # Variables
        self.is_running = False
        self.is_monitoring = False
        self.log_queue = queue.Queue()
        self.usb_analyzer = None
        self.display_analyzer = None
        self.report_generator = None
        self.current_data = None
        self.platform_info = None
        self.log_panel_visible = True
        self.csv_path = csv_path

        # Simple color scheme
        self.colors = {
            'bg': '#1e293b',
            'bg_light': '#334155',
            'bg_card': '#1e2a4a',
            'accent': '#3b82f6',
            'accent_hover': '#60a5fa',
            'success': '#10b981',
            'warning': '#f59e0b',
            'error': '#ef4444',
            'text': '#f1f5f9',
            'text_secondary': '#94a3b8',
            'green': '#22c55e',
            'orange': '#fb923c',
            'blue': '#3b82f6'
        }

        self.root.configure(bg=self.colors['bg'])

        # Build GUI
        self._build_gui()

        # React to window resizing (hide/show the log panel)
        self.root.bind('<Configure>', self._on_window_resize)

        # Clean shutdown: stop the USB monitor thread when the window closes
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # Start log updates
        self._update_log()

    def _build_gui(self):
        """Build the GUI interface."""
        # Main container
        main_frame = tk.Frame(self.root, bg=self.colors['bg'])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # === TOP: Start / Reset buttons ===
        top_frame = tk.Frame(main_frame, bg=self.colors['bg'])
        top_frame.pack(fill=tk.X, pady=(0, 15))

        # Platform info
        self.platform_label = tk.Label(
            top_frame,
            text="Click Start to begin",
            font=('Segoe UI', 12),
            bg=self.colors['bg'],
            fg=self.colors['accent']
        )
        self.platform_label.pack(side=tk.LEFT)

        # Buttons on the right
        btn_right_frame = tk.Frame(top_frame, bg=self.colors['bg'])
        btn_right_frame.pack(side=tk.RIGHT)

        # Reset button (hidden until analysis starts)
        self.reset_btn = tk.Button(
            btn_right_frame,
            text="Reset",
            font=('Segoe UI', 11, 'bold'),
            bg=self.colors['accent'],
            fg='white',
            padx=15,
            pady=5,
            command=self._reset_analysis,
            cursor='hand2'
        )

        # Stop button (hidden until monitoring starts)
        self.stop_btn = tk.Button(
            btn_right_frame,
            text="Stop",
            font=('Segoe UI', 11, 'bold'),
            bg=self.colors['error'],
            fg='white',
            padx=15,
            pady=5,
            command=self._on_stop_clicked,
            cursor='hand2'
        )

        # Start Analysis button
        self.start_btn = tk.Button(
            btn_right_frame,
            text="Start Analysis",
            font=('Segoe UI', 11, 'bold'),
            bg=self.colors['success'],
            fg='white',
            padx=15,
            pady=5,
            command=self._on_start_clicked,
            cursor='hand2'
        )
        self.start_btn.pack(side=tk.RIGHT, padx=(10, 0))

        # === MIDDLE: Tree (left, ~70%) + Log (right, ~30%), resizable ===
        middle_container = tk.Frame(main_frame, bg=self.colors['bg'])
        middle_container.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        # Left: USB tree with stability - the dominant panel
        self.left_frame = tk.Frame(middle_container, bg=self.colors['bg_card'], relief=tk.RAISED, borderwidth=1)
        self.left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        # Tree header
        tree_header = tk.Frame(self.left_frame, bg=self.colors['bg_light'], height=40)
        tree_header.pack(fill=tk.X)
        tree_header_label = tk.Label(
            tree_header,
            text="USB Tree & Stability",
            font=('Segoe UI', 13, 'bold'),
            bg=self.colors['bg_light'],
            fg=self.colors['text']
        )
        tree_header_label.pack(side=tk.LEFT, padx=15, pady=8)

        # Tree text area
        tree_frame = tk.Frame(self.left_frame, bg=self.colors['bg_card'])
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        self.tree_text = tk.Text(
            tree_frame,
            bg=self.colors['bg_card'],
            fg=self.colors['text'],
            font=('Consolas', 11),
            wrap=tk.WORD,
            relief=tk.FLAT,
            borderwidth=0
        )
        self.tree_text.pack(fill=tk.BOTH, expand=True)

        # Right: Live log - connect/disconnect/re-enumeration events as they happen
        self.right_frame = tk.Frame(middle_container, bg=self.colors['bg_card'], relief=tk.RAISED, borderwidth=1)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(0, 0))

        # Log header
        log_header = tk.Frame(self.right_frame, bg=self.colors['bg_light'], height=40)
        log_header.pack(fill=tk.X)
        log_header_label = tk.Label(
            log_header,
            text="Live Log",
            font=('Segoe UI', 13, 'bold'),
            bg=self.colors['bg_light'],
            fg=self.colors['text']
        )
        log_header_label.pack(side=tk.LEFT, padx=15, pady=8)

        # Log text area - using standard Text widget for better tag support
        log_frame = tk.Frame(self.right_frame, bg=self.colors['bg_card'])
        log_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        self.log_text = tk.Text(
            log_frame,
            bg=self.colors['bg'],
            fg=self.colors['text'],
            font=('Consolas', 10),
            wrap=tk.WORD,
            relief=tk.FLAT,
            borderwidth=0
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.log_text.tag_configure('event_connect', foreground=self.colors['green'])
        self.log_text.tag_configure('event_disconnect', foreground=self.colors['orange'])
        self.log_text.tag_configure('timestamp', foreground=self.colors['text_secondary'])

        # === BOTTOM: Report and Action buttons ===
        bottom_frame = tk.Frame(main_frame, bg=self.colors['bg'])
        bottom_frame.pack(fill=tk.X)

        # Status
        self.status_label = tk.Label(
            bottom_frame,
            text="Ready",
            font=('Segoe UI', 11),
            bg=self.colors['bg'],
            fg=self.colors['success']
        )
        self.status_label.pack(side=tk.LEFT)

        # Buttons
        btn_frame = tk.Frame(bottom_frame, bg=self.colors['bg'])
        btn_frame.pack(side=tk.RIGHT)

        # Report button
        report_btn = tk.Button(
            btn_frame,
            text="Generate Report...",
            font=('Segoe UI', 11, 'bold'),
            bg=self.colors['accent'],
            fg='white',
            padx=15,
            pady=5,
            command=self._generate_report_dialog,
            cursor='hand2'
        )
        report_btn.pack(side=tk.LEFT)

    def _on_window_resize(self, event):
        """Handle window resize to hide/show the log panel."""
        if event.widget != self.root:
            return

        width = self.root.winfo_width()
        should_show = width >= NARROW_WINDOW_THRESHOLD

        if should_show and not self.log_panel_visible:
            self.right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(0, 0))
            self.left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
            self.log_panel_visible = True
        elif not should_show and self.log_panel_visible:
            self.right_frame.pack_forget()
            self.log_panel_visible = False

    def _on_start_clicked(self):
        """Handle Start button click."""
        self.start_btn.pack_forget()
        self.stop_btn.pack(side=tk.RIGHT, padx=(10, 0))
        self.reset_btn.pack(side=tk.RIGHT)
        self._start_analysis()

    def _on_stop_clicked(self):
        """Handle Stop button click — stops live monitoring, keeps results."""
        self._stop_monitoring()
        self.stop_btn.pack_forget()
        self.status_label.configure(text="Stopped", text_color=self.colors['error'])

    def _start_analysis(self):
        """Start the analysis in a background thread."""
        if self.is_running:
            return

        self.is_running = True
        self.status_label.configure(text="Analyzing...", text_color=self.colors['warning'])
        self.log_text.delete('0.0', tk.END)
        self.tree_text.delete('0.0', tk.END)

        # Redirect stdout to the log
        sys.stdout = LogRedirector(self.log_text, self.log_queue)

        # Start thread
        thread = threading.Thread(target=self._run_analysis, daemon=True)
        thread.start()

    def _run_analysis(self):
        """Run the initial analysis in the background, then start live monitoring."""
        try:
            # 1. Platform
            self.platform_info = PlatformUtils.get_platform_info()
            print("=" * 60)
            print("  KLANGCHE PROAV SHOKO - USB DETECTIVE")
            print("=" * 60)
            print(f"\n[+] Platform: {self.platform_info['os']} {self.platform_info['version']}")
            print(f"[+] Architecture: {self.platform_info['architecture']}")
            if self.platform_info['is_apple_silicon']:
                print("[+] Apple Silicon detected!")

            self.root.after(0, self._update_platform_label)

            # 2. USB analysis - load limits from custom data file if specified, otherwise default
            self.usb_analyzer = USBAnalyzer(self.csv_path) if self.csv_path else USBAnalyzer()
            if self.csv_path:
                print(f"[+] Loaded limits from: {self.csv_path}")
            elif Path("usb_data.csv").exists():
                print(f"[+] Loaded limits from: usb_data.csv")

            print("\n[+] Scanning USB devices...")
            self._refresh_tree_and_stability()

            # 3. Display information
            print("\n[+] Scanning displays...")
            self.display_analyzer = DisplayAnalyzer()

            print("\n[+] Initial scan complete. Starting live monitoring...")
            self.root.after(0, lambda: self.status_label.configure(
                text="Monitoring (live)",
                text_color=self.colors['success']
            ))

            # 4. Start live connect/disconnect monitoring.
            self.usb_analyzer.start_live_monitoring(
                on_connect=self._on_device_connect,
                on_disconnect=self._on_device_disconnect,
                check_every_seconds=0.1
            )
            self.is_monitoring = True

        except Exception as e:
            err_msg = str(e)
            print(f"\n[!] Error: {err_msg}")
            self.root.after(0, lambda m=err_msg: self.status_label.configure(
                text=f"Error: {m}",
                text_color=self.colors['error']
            ))
        finally:
            self.is_running = False
            if hasattr(sys.stdout, 'original_stdout'):
                sys.stdout = sys.stdout.original_stdout

    def _refresh_tree_and_stability(self):
        """Re-scan the USB tree, recompute hops/tiers/stability and update the GUI.
        Called on startup and again after every connect/disconnect event."""
        usb_tree = self.usb_analyzer.build_tree()
        hops_data = self.usb_analyzer.calculate_hops_and_tiers(usb_tree)
        stability = self.usb_analyzer.assess_stability(hops_data)

        displays = self.display_analyzer.get_display_info() if self.display_analyzer else []

        self.current_data = {
            'usb_tree': usb_tree,
            'hops_data': hops_data,
            'stability': stability,
            'displays': displays,
            'platform_info': self.platform_info
        }

        self.root.after(0, self._update_tree_display, usb_tree, hops_data, stability, displays)
        print(self.usb_analyzer.get_stability_summary(stability))

    def _on_device_connect(self, device_id, device_info):
        """Callback from usbmonitor's background thread when a device connects."""
        model = device_info.get('ID_MODEL', device_id)
        self.root.after(0, self._log_event, f"CONNECTED: {model}", 'event_connect')
        self._refresh_tree_and_stability()

    def _on_device_disconnect(self, device_id, device_info):
        """Callback from usbmonitor's background thread when a device disconnects."""
        model = device_info.get('ID_MODEL', device_id)
        self.root.after(0, self._log_event, f"DISCONNECTED: {model}", 'event_disconnect')
        self._refresh_tree_and_stability()

    def _log_event(self, message, tag):
        """Append a timestamped live event line to the log panel."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] ", ('timestamp',))
        self.log_text.insert(tk.END, f"{message}\n", (tag,))
        self.log_text.see(tk.END)

    def _update_platform_label(self):
        """Update the platform label."""
        if self.platform_info:
            text = f"{self.platform_info['os']} {self.platform_info['architecture']}"
            if self.platform_info['is_apple_silicon']:
                text += " (Apple Silicon)"
            self.platform_label.config(text=text)

    def _update_tree_display(self, usb_tree, hops_data, stability, displays):
        """Update the tree display to match CLI exactly."""
        self.tree_text.delete('0.0', tk.END)

        # VERDICTS
        self.tree_text.insert(tk.END, "=" * 31 + "VERDICT" + "=" * 31)
        self.tree_text.insert(tk.END, "\n")
        for v in stability.get('verdicts', []):
            status_char = '+' if v['color'] == 'green' else ('~' if v['color'] == 'orange' else '!')
            hubs_str = f"hubs {v['current_hubs']}/{v['max_hubs']}  " if 'current_hubs' in v else ""
            desc = v.get('description', v.get('name', ''))
            self.tree_text.insert(tk.END, f"  {status_char} {desc:<22s} ")
            self.tree_text.insert(tk.END, f"{v['status']:<9s} ")
            self.tree_text.insert(tk.END, f"hops {v['current_hops']}/{v['max_hops']}  ")
            self.tree_text.insert(tk.END, f"tiers {v['current_tiers']}/{v['max_tiers']}  ")
            self.tree_text.insert(tk.END, f"{hubs_str}\n")
        self.tree_text.insert(tk.END, "\n")

        # FULL USB & Display Tree
        self.tree_text.insert(tk.END, "  Full USB & Display Tree")
        if usb_tree:
            root_node = usb_tree[0]

            # Add displays directly into tree (not under a "Displays" parent)
            if displays:
                for i, d in enumerate(displays):
                    prim = " (Primary)" if d.get('is_primary', False) else ""
                    int_disp = d.get('is_internal', False)
                    root_node.setdefault('children', []).append({
                        'model': f"{d['resolution']}  {d['name']}{prim}",
                        'name': d['name'], 'children': [], 'hops': 1,
                        'is_hub': False, 'is_internal': int_disp, 'is_display': True, 'port': 0
                    })

            self._print_tree(usb_tree, "  ", _show_internal=True)
        else:
            self.tree_text.insert(tk.END, "  No USB devices found.\n")
        self.tree_text.insert(tk.END, "\n")

        # PER PORT - EXTERNAL
        root_orig = usb_tree[0] if usb_tree else {}
        orig_children = list(root_orig.get('children', []))
        self.tree_text.insert(tk.END, "=" * 31 + "EXTERNAL" + "=" * 31)
        self.tree_text.insert(tk.END, "\n")
        self._print_port_section(orig_children, ports_data=stability.get('ports', []), is_internal_filter=lambda c: not c.get('is_internal', False))
        self.tree_text.insert(tk.END, "\n" + "  " + "- " * 35 + "\n")

        # PER PORT - INTERNAL
        self.tree_text.insert(tk.END, "=" * 31 + "INTERNAL" + "=" * 31)
        self.tree_text.insert(tk.END, "\n")
        self._print_port_section(orig_children, ports_data=stability.get('ports', []), is_internal_filter=lambda c: c.get('is_internal', False))
        self.tree_text.insert(tk.END, "\n")

    def _print_tree(self, nodes, prefix="", _show_internal=False, _parent_is_internal=False):
        """Print tree with box-drawing characters, matching CLI exactly."""
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

            label = self._node_label(node)
            self.tree_text.insert(tk.END, f"{prefix}{connector}{badge_str}{label}{port_info_str}\n")

            if node.get('children'):
                child_prefix = prefix + ("    " if is_last else "│   ")
                new_parent_int = _parent_is_internal or node.get('is_internal', False)
                self._print_tree(node['children'], child_prefix, _show_internal, new_parent_int)

    def _print_port_section(self, orig_children, ports_data, is_internal_filter):
        """Print a section (EXTERNAL or INTERNAL) with per-port info."""
        sep = "  " + "- " * 35
        first = True
        for idx, child in enumerate(orig_children):
            if child.get('is_display'):
                continue
            if not is_internal_filter(child):
                continue
            if not first:
                self.tree_text.insert(tk.END, "\n")
            self._print_port_child(child, idx, ports_data)
            self.tree_text.insert(tk.END, "\n")
            self.tree_text.insert(tk.END, sep)
            self.tree_text.insert(tk.END, "\n")
            first = False

    def _print_port_child(self, child, idx, ports_data):
        """Print a single port's tree and stability info."""
        port_info = next((p for p in ports_data if p.get('id') == idx + 1), None)
        label = port_info['label'] if port_info else child.get('model', 'Port')
        dc = len(port_info['devices']) if port_info else 0
        ph = port_info['max_hops'] if port_info else 0
        pt = port_info['max_tiers'] if port_info else 0
        p_hub = port_info.get('external_hubs', 0) if port_info else 0
        is_int = child.get('is_internal', False)
        int_pre = "[INTERNAL] " if is_int else ""
        ep_label = "endpoint" if dc == 1 else "endpoints"
        self.tree_text.insert(tk.END, f"  {int_pre}{label} ({dc} {ep_label}, hops={ph}, tiers={pt}, hubs={p_hub})")
        self._print_port_tree(child)
        if is_int:
            self.tree_text.insert(tk.END, "    (internal)\n")
        elif port_info:
            self._print_stability_port(port_info, port_info['verdicts'])

    def _print_port_tree(self, port_node):
        """Print tree for a single port (children of the port node)."""
        children = port_node.get('children', [])
        if children:
            self._print_tree(children, "    ")

    def _print_stability_port(self, port_info, stability_data):
        """Print stability verdicts for a single port."""
        for v in stability_data:
            status_char = '+' if v['color'] == 'green' else ('~' if v['color'] == 'orange' else '!')
            hubs_str = f"hubs {v['current_hubs']}/{v['max_hubs']}  " if 'current_hubs' in v else ""
            desc = v.get('description', v.get('name', ''))
            self.tree_text.insert(tk.END, f"    {status_char} {desc:<22s} ")
            self.tree_text.insert(tk.END, f"{v['status']:<9s} ")
            self.tree_text.insert(tk.END, f"hops {v['current_hops']}/{v['max_hops']}  ")
            self.tree_text.insert(tk.END, f"tiers {v['current_tiers']}/{v['max_tiers']}  ")
            self.tree_text.insert(tk.END, f"{hubs_str}\n")

    @staticmethod
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

    def _update_log(self):
        """Update the log widget from the queue (stdout redirection)."""
        try:
            while True:
                text = self.log_queue.get_nowait()
                self.log_text.insert(tk.END, text)
                self.log_text.see(tk.END)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._update_log)

    def _stop_monitoring(self):
        """Stop live USB monitoring if it's running."""
        if self.is_monitoring and self.usb_analyzer:
            self.usb_analyzer.stop_live_monitoring()
            self.is_monitoring = False

    def _reset_analysis(self):
        """Reset and restart the analysis."""
        if messagebox.askyesno("Reset", "Do you want to reset and restart the analysis?"):
            self._stop_monitoring()
            self.stop_btn.pack_forget()
            self.start_btn.pack(side=tk.RIGHT, padx=(10, 0))
            self._start_analysis()

    def _on_close(self):
        """Stop the background monitoring thread and close the window."""
        self._stop_monitoring()
        self.root.destroy()

    def _export_csv_limits(self):
        """Export hop and tier limits as CSV."""
        if not self.usb_analyzer:
            messagebox.showwarning("No analysis", "Run an analysis first!")
            return

        default_name = f"usb_data_{datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}.csv"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=default_name,
            title="Save USB data as"
        )

        if not file_path:
            return

        try:
            self.usb_analyzer.save_usb_data(file_path)
            messagebox.showinfo("Done", f"USB data saved:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save data: {e}")

    def _generate_report_dialog(self):
        """Show dialog for report generation with format and port selection."""
        if not self.current_data:
            messagebox.showwarning("No data", "Run an analysis first!")
            return

        stability = self.current_data['stability']
        ports_data = stability.get('ports', [])

        # Create dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Generate Report")
        dialog.geometry("600x550")
        dialog.resizable(False, False)
        dialog.configure(bg=self.colors['bg'])
        dialog.transient(self.root)
        dialog.grab_set()

        # Center on parent
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 600) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 550) // 2
        dialog.geometry(f"+{x}+{y}")

        # Title
        title_label = tk.Label(
            dialog,
            text="Generate Report",
            font=('Segoe UI', 12, 'bold'),
            bg=self.colors['bg'],
            fg=self.colors['blue']
        )
        title_label.pack(pady=(15, 5))

        # Format selection
        format_frame = tk.LabelFrame(dialog, text="Report Format", padding=15)
        format_frame.pack(fill=tk.X, padx=20, pady=10)

        self.report_format_var = tk.StringVar(value='html')
        html_radio = tk.Radiobutton(
            format_frame,
            text="HTML Report (opens in browser)",
            variable=self.report_format_var,
            value='html',
            font=('Segoe UI', 10),
            bg=self.colors['bg'],
            fg=self.colors['fg'],
            selectcolor=self.colors['bg_card'],
            activebackground=self.colors['bg'],
            activeforeground=self.colors['blue']
        )
        html_radio.pack(anchor=tk.W, pady=2)

        pdf_radio = tk.Radiobutton(
            format_frame,
            text="PDF Report (requires weasyprint)",
            variable=self.report_format_var,
            value='pdf',
            font=('Segoe UI', 10),
            bg=self.colors['bg'],
            fg=self.colors['fg'],
            selectcolor=self.colors['bg_card'],
            activebackground=self.colors['bg'],
            activeforeground=self.colors['blue']
        )
        pdf_radio.pack(anchor=tk.W, pady=2)

        # Port selection
        port_frame = tk.LabelFrame(dialog, text="Port Selection", padding=15)
        port_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # Create checkboxes for ports
        self.port_vars = {}
        self.port_vars['Full'] = tk.BooleanVar(value=True)

        # Full option
        full_frame = tk.Frame(port_frame, bg=self.colors['bg'])
        full_frame.pack(fill=tk.X, pady=2)
        full_cb = tk.Checkbutton(
            full_frame,
            text="Full - All ports combined (overall assessment)",
            variable=self.port_vars['Full'],
            font=('Segoe UI', 10),
            bg=self.colors['bg'],
            fg=self.colors['fg'],
            selectcolor=self.colors['bg_card'],
            activebackground=self.colors['bg'],
            activeforeground=self.colors['blue']
        )
        full_cb.pack(side=tk.LEFT)

        # Separator
        ttk.Separator(port_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=8)

        # Filter to only external (non-internal) ports
        # Get the original root children to check is_internal flag
        usb_tree = self.current_data['usb_tree']
        root_orig = usb_tree[0] if usb_tree else {}
        orig_children = list(root_orig.get('children', []))

        external_ports = []
        for i, port in enumerate(ports_data):
            if i < len(orig_children):
                child = orig_children[i]
                if not child.get('is_internal', False) and not child.get('is_display', False):
                    external_ports.append(port)

        # Individual ports (external only)
        for port in external_ports:
            label = port.get('label', f"Port {port.get('id', '?')}")
            devices = len(port.get('devices', []))
            hops = port.get('max_hops', 0)
            tiers = port.get('max_tiers', 0)
            hubs = port.get('external_hubs', 0)

            port_var = tk.BooleanVar(value=False)
            self.port_vars[label] = port_var

            port_row = tk.Frame(port_frame, bg=self.colors['bg'])
            port_row.pack(fill=tk.X, pady=2)

            ep_label = "endpoint" if devices == 1 else "endpoints"
            cb = tk.Checkbutton(
                port_row,
                text=f"{label}  ({devices} {ep_label}, hops={hops}, tiers={tiers}, hubs={hubs})",
                variable=port_var,
                font=('Segoe UI', 10),
                bg=self.colors['bg'],
                fg=self.colors['fg'],
                selectcolor=self.colors['bg_card'],
                activebackground=self.colors['bg'],
                activeforeground=self.colors['blue']
            )
            cb.pack(side=tk.LEFT)

        # Buttons
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=20, pady=15)

        generate_btn = tk.Button(
            btn_frame,
            text="Generate",
            font=('Segoe UI', 10, 'bold'),
            bg=self.colors['blue'],
            fg='white',
            padx=20,
            pady=8,
            command=lambda: self._do_generate_report(dialog),
            cursor='hand2'
        )
        generate_btn.pack(side=tk.RIGHT, padx=(10, 0))

        cancel_btn = tk.Button(
            btn_frame,
            text="Cancel",
            font=('Segoe UI', 10),
            bg='#2d2d44',
            fg='white',
            padx=20,
            pady=8,
            command=dialog.destroy,
            cursor='hand2'
        )
        cancel_btn.pack(side=tk.RIGHT)

    def _do_generate_report(self, dialog):
        """Generate the report based on dialog selections."""
        format_type = self.report_format_var.get()

        # Collect selected ports
        selected_ports = []
        for label, var in self.port_vars.items():
            if var.get():
                selected_ports.append(label)

        if not selected_ports:
            selected_ports = ['Full']

        dialog.destroy()

        # Ask for save location
        ext = '.html' if format_type == 'html' else '.pdf'
        default_name = f"proav-shoko-{datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}{ext}"
        file_types = [("HTML files", "*.html")] if format_type == 'html' else [("PDF files", "*.pdf")]

        file_path = filedialog.asksaveasfilename(
            defaultextension=ext,
            filetypes=file_types + [("All files", "*.*")],
            initialfile=default_name,
            title=f"Save {format_type.upper()} Report"
        )

        if not file_path:
            return

        try:
            self.status_label.config(text=f"Generating {format_type.upper()}...", foreground=self.colors['yellow'])
            self.report_generator = ReportGenerator()

            if format_type == 'html':
                html_path = self.report_generator.generate_html_report(
                    self.current_data['usb_tree'],
                    self.current_data['hops_data'],
                    self.current_data['stability'],
                    self.current_data['displays'],
                    self.current_data['platform_info'],
                    platform_notes=self.usb_analyzer.get_platform_notes() if self.usb_analyzer else None,
                    selected_ports=selected_ports,
                    custom_path=file_path
                )
                self.status_label.config(text="HTML saved", foreground=self.colors['green'])
                print(f"\n[+] HTML Report saved: {html_path}")
                messagebox.showinfo("Done", f"HTML report saved:\n{html_path}")
                self.report_generator.open_report(html_path)
            else:
                # For PDF, generate HTML first then convert
                import tempfile
                with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False) as tmp:
                    tmp_path = tmp.name

                html_path = self.report_generator.generate_html_report(
                    self.current_data['usb_tree'],
                    self.current_data['hops_data'],
                    self.current_data['stability'],
                    self.current_data['displays'],
                    self.current_data['platform_info'],
                    platform_notes=self.usb_analyzer.get_platform_notes() if self.usb_analyzer else None,
                    selected_ports=selected_ports,
                    custom_path=tmp_path
                )

                pdf_path = self.report_generator.generate_pdf_report(html_path, custom_path=file_path)

                try:
                    Path(tmp_path).unlink()
                except:
                    pass

                if pdf_path:
                    self.status_label.config(text="PDF saved", foreground=self.colors['green'])
                    print(f"\n[+] PDF Report saved: {pdf_path}")
                    messagebox.showinfo("Done", f"PDF report saved:\n{pdf_path}")
                    self.report_generator.open_report(pdf_path)
                else:
                    messagebox.showwarning("Warning", "PDF generation failed. Check that weasyprint is installed.")
                    self.status_label.config(text="PDF failed", foreground=self.colors['red'])

        except Exception as e:
            messagebox.showerror("Error", f"Could not generate report: {e}")
            self.status_label.config(text=f"Error: {e}", foreground=self.colors['red'])

    def save_logs(self, path):
        """Save analysis logs to a file."""
        try:
            with open(path, 'w', encoding='utf-8') as f:
                # Add header
                f.write(f"ProAV Shoko Log - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("=" * 60 + "\n")
                
                # Platform info
                f.write(f"Platform: {self.platform_info['os']} {self.platform_info['version']}\n")
                f.write(f"Architecture: {self.platform_info['architecture']}\n")
                if self.platform_info['is_apple_silicon']:
                    f.write("Apple Silicon: Yes\n")
                
                # Tree and logs
                f.write("\n" + "=" * 60 + "\n")
                f.write("USB TREE STRUCTURE\n")
                f.write("=" * 60 + "\n")
                if self.current_data and self.current_data['usb_tree']:
                    self._render_tree_to_file(f, self.current_data['usb_tree'], 0)
                else:
                    f.write("No USB devices found.\n")
                
                f.write("\n" + "=" * 60 + "\n")
                f.write("STABILITY ASSESSMENT\n")
                f.write("=" * 60 + "\n")
                if self.current_data and 'stability' in self.current_data:
                    f.write(self.usb_analyzer.get_stability_summary(self.current_data['stability']))
                
                f.write("\n" + "=" * 60 + "\n")
                f.write("DISPLAY INFORMATION\n")
                f.write("=" * 60 + "\n")
                if self.current_data and self.current_data.get('displays'):
                    for display in self.current_data['displays']:
                        primary = " (Primary)" if display.get('is_primary', False) else ""
                        f.write(f"  {display['resolution']}  {display['name']}{primary}\n")
                elif self.current_data and not self.current_data.get('displays'):
                    f.write("  No displays found.\n")
                
                print(f"[+] Log saved: {path}")
        except Exception as e:
            print(f"[!] Could not save log: {e}")

    def _render_tree_to_file(self, f, tree, level):
        """Recursively write USB tree to file."""
        indent = "  " * level
        for node in tree:
            is_hub = node.get('is_hub', False)
            hub_tag = " [HUB]" if is_hub else ""
            hops = node['devpath'].count('/') if node.get('devpath') else 0

            line = f"{indent}{node.get('model', 'Unknown')}{hub_tag}  hops: {hops}\n"
            f.write(line)

            if node.get('children'):
                self._render_tree_to_file(f, node['children'], level + 1)

    def _save_logs_dialog(self):
        """Show a dialog to save logs to a file."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".log",
            filetypes=[("Log files", "*.log"), ("All files", "*.*")],
            title="Save Logs"
        )
        if file_path:
            self.save_logs(file_path)

    def main():
        """Start the GUI application."""
        root = tk.Tk()
        app = ProAVShokoGUI(root)
        root.mainloop()


if __name__ == "__main__":
    main()