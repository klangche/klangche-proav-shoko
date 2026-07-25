"""
GUI for ProAV Shoko - live overview with tree and log
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
        self.root.geometry("1400x800")
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

        # Colors
        self.colors = {
            'bg': '#1a1a2e',
            'bg_light': '#16213e',
            'bg_card': '#1e2a4a',
            'fg': '#e0e0e0',
            'green': '#00cc66',
            'yellow': '#ffcc00',
            'orange': '#ff8800',
            'red': '#ff3333',
            'blue': '#00d4ff'
        }

        # Set theme
        self.root.configure(bg=self.colors['bg'])
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TFrame', background=self.colors['bg'])
        style.configure('TLabelframe', background=self.colors['bg'], foreground=self.colors['fg'])
        style.configure('TLabelframe.Label', background=self.colors['bg'], foreground=self.colors['fg'])
        style.configure('TPanedwindow', background=self.colors['bg'])

        # Build GUI
        self._build_gui()

        # React to window resizing (hide/show the log panel)
        self.root.bind('<Configure>', self._on_window_resize)

        # Clean shutdown: stop the USB monitor thread when the window closes
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # Start log updates
        self._update_log()

        # Start analysis immediately
        self.root.after(500, self._start_analysis)

    def _build_gui(self):
        """Build the interface."""
        # Main frame
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # === TOP: Reset button ===
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))

        # Platform info
        self.platform_label = ttk.Label(
            top_frame,
            text="Loading platform info...",
            font=('Segoe UI', 10),
            foreground=self.colors['blue']
        )
        self.platform_label.pack(side=tk.LEFT)

        # Reset button
        reset_btn = tk.Button(
            top_frame,
            text="Reset",
            font=('Segoe UI', 10, 'bold'),
            bg=self.colors['blue'],
            fg='white',
            padx=15,
            pady=5,
            command=self._reset_analysis,
            cursor='hand2'
        )
        reset_btn.pack(side=tk.RIGHT)

        # === MIDDLE: Tree (left, ~70%) + Log (right, ~30%), resizable ===
        self.middle_paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        self.middle_paned.pack(fill=tk.BOTH, expand=True)

        # Left: USB tree with stability - the dominant panel
        self.left_frame = ttk.LabelFrame(self.middle_paned, text="USB Tree & Stability", padding=10)

        self.tree_text = tk.Text(
            self.left_frame,
            bg=self.colors['bg_card'],
            fg=self.colors['fg'],
            font=('Courier New', 10),
            wrap=tk.NONE,
            relief=tk.FLAT,
            padx=10,
            pady=10
        )
        self.tree_text.pack(fill=tk.BOTH, expand=True)

        tree_scroll_y = ttk.Scrollbar(self.left_frame, orient=tk.VERTICAL, command=self.tree_text.yview)
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scroll_x = ttk.Scrollbar(self.left_frame, orient=tk.HORIZONTAL, command=self.tree_text.xview)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree_text.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)

        # Right: Live log - connect/disconnect/re-enumeration events as they happen
        self.right_frame = ttk.LabelFrame(self.middle_paned, text="Live Log", padding=10)

        self.log_text = tk.Text(
            self.right_frame,
            bg='#0a0a1a',
            fg='#88ccff',
            font=('Courier New', 9),
            wrap=tk.WORD,
            relief=tk.FLAT,
            padx=10,
            pady=10
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        log_scroll_y = ttk.Scrollbar(self.right_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        log_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=log_scroll_y.set)

        self.log_text.tag_configure('event_connect', foreground=self.colors['green'])
        self.log_text.tag_configure('event_disconnect', foreground=self.colors['orange'])
        self.log_text.tag_configure('timestamp', foreground='#666666')

        self.middle_paned.add(self.left_frame, weight=7)
        self.middle_paned.add(self.right_frame, weight=3)

        # === BOTTOM: Report buttons ===
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=(10, 0))

        # Status
        self.status_label = ttk.Label(
            bottom_frame,
            text="Ready",
            font=('Segoe UI', 9),
            foreground=self.colors['green']
        )
        self.status_label.pack(side=tk.LEFT)

        # Buttons
        btn_frame = ttk.Frame(bottom_frame)
        btn_frame.pack(side=tk.RIGHT)

        # Export CSV button
        csv_btn = tk.Button(
            btn_frame,
            text="Export CSV Limits",
            font=('Segoe UI', 9, 'bold'),
            bg='#2d2d44',
            fg='white',
            padx=10,
            pady=5,
            command=self._export_csv_limits,
            cursor='hand2'
        )
        csv_btn.pack(side=tk.LEFT, padx=(0, 10))

        html_btn = tk.Button(
            btn_frame,
            text="Print HTML Report",
            font=('Segoe UI', 10, 'bold'),
            bg='#2d2d44',
            fg='white',
            padx=15,
            pady=5,
            command=self._print_html_report,
            cursor='hand2'
        )
        html_btn.pack(side=tk.LEFT, padx=(0, 10))

        pdf_btn = tk.Button(
            btn_frame,
            text="Print PDF Report",
            font=('Segoe UI', 10, 'bold'),
            bg='#2d2d44',
            fg='white',
            padx=15,
            pady=5,
            command=self._print_pdf_report,
            cursor='hand2'
        )
        pdf_btn.pack(side=tk.LEFT, padx=(0, 10))

        log_btn = tk.Button(
            btn_frame,
            text="Save Logs",
            font=('Segoe UI', 10, 'bold'),
            bg='#2d2d44',
            fg='white',
            padx=15,
            pady=5,
            command=self._save_logs_dialog,
            cursor='hand2'
        )
        log_btn.pack(side=tk.LEFT)

    def _on_window_resize(self, event):
        """Hide the live log panel when the window gets narrow, show it again otherwise."""
        if event.widget != self.root:
            return

        width = self.root.winfo_width()
        should_show = width >= NARROW_WINDOW_THRESHOLD

        if should_show and not self.log_panel_visible:
            self.middle_paned.add(self.right_frame, weight=3)
            self.log_panel_visible = True
        elif not should_show and self.log_panel_visible:
            self.middle_paned.forget(self.right_frame)
            self.log_panel_visible = False

    def _start_analysis(self):
        """Start the analysis in a background thread."""
        if self.is_running:
            return

        self.is_running = True
        self.status_label.config(text="Analyzing...", foreground=self.colors['yellow'])
        self.log_text.delete(1.0, tk.END)
        self.tree_text.delete(1.0, tk.END)

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

            # 2. USB analysis - load limits from custom CSV if specified, otherwise default
            self.usb_analyzer = USBAnalyzer(self.csv_path) if self.csv_path else USBAnalyzer()
            if self.csv_path:
                print(f"[+] Loaded limits from: {self.csv_path}")
            elif Path("hop_limits.csv").exists():
                print(f"[+] Loaded limits from: hop_limits.csv")
            else:
                self.usb_analyzer.save_hop_limits_csv("hop_limits.csv")
                print(f"[+] Created default hop_limits.csv")

            print("\n[+] Scanning USB devices...")
            self._refresh_tree_and_stability()

            # 3. Display information
            print("\n[+] Scanning displays...")
            self.display_analyzer = DisplayAnalyzer()

            print("\n[+] Initial scan complete. Starting live monitoring...")
            self.root.after(0, lambda: self.status_label.config(
                text="Monitoring (live)",
                foreground=self.colors['green']
            ))

            # 4. Start live connect/disconnect monitoring. This keeps
            # running in usbmonitor's own background thread until
            # stop_live_monitoring() is called (on reset or window close).
            self.usb_analyzer.start_live_monitoring(
                on_connect=self._on_device_connect,
                on_disconnect=self._on_device_disconnect,
                check_every_seconds=1.0
            )
            self.is_monitoring = True

        except Exception as e:
            print(f"\n[!] Error: {e}")
            self.root.after(0, lambda: self.status_label.config(
                text=f"Error: {e}",
                foreground=self.colors['red']
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
        """Update the tree display."""
        self.tree_text.delete(1.0, tk.END)

        # Current hops/tiers
        self.tree_text.insert(tk.END, "\n")
        self.tree_text.insert(
            tk.END,
            f"Current: {hops_data['max_hops']} hops, {hops_data['max_tiers']} tiers\n",
            ('info',)
        )

        # Stability heading
        self.tree_text.insert(tk.END, "STABILITY VERDICT\n", ('header',))
        self.tree_text.insert(tk.END, "-" * 50 + "\n", ('header',))

        groups = stability.get('groups', {})
        for arch, verdicts in groups.items():
            self.tree_text.insert(tk.END, f"\n{arch}\n", ('arch_header',))
            for v in verdicts:
                emoji = v['emoji']
                name = v['name']
                status = v['status']
                color = v['color']
                tag = f'stability_{color}'
                self.tree_text.insert(tk.END, f"  {emoji} {name}  ", (tag,))
                self.tree_text.insert(tk.END, f"({status})  ", ('info',))
                self.tree_text.insert(
                    tk.END,
                    f"max {v['max_hops']} hops / {v['max_tiers']} tiers\n",
                    ('dim',)
                )

        # Warnings
        warnings = [v for v in stability.get('verdicts', []) if not v['is_stable'] or v['warning']]
        if warnings:
            self.tree_text.insert(tk.END, "\nWARNINGS:\n", ('warning_header',))
            for w in warnings:
                self.tree_text.insert(
                    tk.END,
                    f"  - {w['name']}: {w['warning']}\n",
                    ('warning',)
                )

        # USB tree
        self.tree_text.insert(tk.END, "\n\nUSB TREE STRUCTURE\n", ('header',))
        self.tree_text.insert(tk.END, "-" * 50 + "\n", ('header',))

        if usb_tree:
            self._render_tree_to_text(usb_tree, 0)
        else:
            self.tree_text.insert(tk.END, "  No USB devices found.\n", ('info',))

        # Displays
        self.tree_text.insert(tk.END, "\nDISPLAY INFORMATION\n", ('header',))
        self.tree_text.insert(tk.END, "-" * 50 + "\n", ('header',))

        if displays:
            for display in displays:
                primary = " (Primary)" if display.get('is_primary', False) else ""
                self.tree_text.insert(
                    tk.END,
                    f"  {display['resolution']}  {display['name']}{primary}\n",
                    ('info',)
                )
        else:
            self.tree_text.insert(tk.END, "  No displays found.\n", ('info',))

        # Configure tags for colors
        self.tree_text.tag_configure('header', font=('Courier New', 11, 'bold'), foreground=self.colors['blue'])
        self.tree_text.tag_configure('arch_header', font=('Courier New', 10, 'bold'), foreground=self.colors['blue'])
        self.tree_text.tag_configure('warning_header', font=('Courier New', 10, 'bold'), foreground=self.colors['orange'])
        self.tree_text.tag_configure('warning', font=('Courier New', 10), foreground=self.colors['orange'])
        self.tree_text.tag_configure('info', foreground=self.colors['fg'])
        self.tree_text.tag_configure('dim', foreground='#666666')
        self.tree_text.tag_configure('stability_green', foreground=self.colors['green'])
        self.tree_text.tag_configure('stability_yellow', foreground=self.colors['yellow'])
        self.tree_text.tag_configure('stability_orange', foreground=self.colors['orange'])
        self.tree_text.tag_configure('stability_red', foreground=self.colors['red'])

        self.tree_text.see(1.0)

    def _render_tree_to_text(self, tree, level):
        """Recursively render the USB tree into the text widget."""
        indent = "  " * level
        for node in tree:
            is_hub = node.get('is_hub', False)
            hub_tag = " [HUB]" if is_hub else ""
            hops = node['devpath'].count('/') if node.get('devpath') else 0

            line = f"{indent}{node.get('model', 'Unknown')}{hub_tag}  hops: {hops}\n"
            self.tree_text.insert(tk.END, line, ('info',))

            if node.get('children'):
                self._render_tree_to_text(node['children'], level + 1)

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
            self._start_analysis()

    def _on_close(self):
        """Stop the background monitoring thread cleanly before closing."""
        self._stop_monitoring()

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

    def _export_csv_limits(self):
        """Export hop and tier limits as CSV."""
        if not self.usb_analyzer:
            messagebox.showwarning("No analysis", "Run an analysis first!")
            return

        default_name = f"hop_limits_{datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}.csv"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=default_name,
            title="Save hop/tier limits as CSV"
        )

        if not file_path:
            return

        try:
            self.usb_analyzer.save_hop_limits_csv(file_path)
            messagebox.showinfo("Done", f"CSV file saved:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save CSV: {e}")

    def _print_html_report(self):
        """Generate and save the HTML report."""
        if not self.current_data:
            messagebox.showwarning("No data", "Run an analysis first!")
            return

        default_name = f"proav-shoko-{datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}.html"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".html",
            filetypes=[("HTML files", "*.html"), ("All files", "*.*")],
            initialfile=default_name,
            title="Save HTML report"
        )

        if not file_path:
            return

        try:
            self.status_label.config(text="Generating HTML...", foreground=self.colors['yellow'])
            self.report_generator = ReportGenerator()
            html_path = self.report_generator.generate_html_report(
                self.current_data['usb_tree'],
                self.current_data['hops_data'],
                self.current_data['stability'],
                self.current_data['displays'],
                self.current_data['platform_info'],
                platform_notes=self.usb_analyzer.get_platform_notes() if self.usb_analyzer else None,
                custom_path=file_path
            )
            self.status_label.config(text="HTML saved", foreground=self.colors['green'])
            print(f"\n[+] HTML Report saved: {html_path}")
            messagebox.showinfo("Done", f"HTML report saved:\n{html_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not generate HTML report: {e}")
            self.status_label.config(text=f"Error: {e}", foreground=self.colors['red'])

    def _print_pdf_report(self):
        """Generate and save the PDF report."""
        if not self.current_data:
            messagebox.showwarning("No data", "Run an analysis first!")
            return

        default_name = f"proav-shoko-{datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}.pdf"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialfile=default_name,
            title="Save PDF report"
        )

        if not file_path:
            return

        try:
            self.status_label.config(text="Generating PDF...", foreground=self.colors['yellow'])

            import tempfile
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False) as tmp:
                tmp_path = tmp.name

            self.report_generator = ReportGenerator()
            html_path = self.report_generator.generate_html_report(
                self.current_data['usb_tree'],
                self.current_data['hops_data'],
                self.current_data['stability'],
                self.current_data['displays'],
                self.current_data['platform_info'],
                platform_notes=self.usb_analyzer.get_platform_notes() if self.usb_analyzer else None,
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
            else:
                messagebox.showwarning("Warning", "PDF generation failed. Check that weasyprint is installed.")
        except Exception as e:
            messagebox.showerror("Error", f"Could not generate PDF report: {e}")
            self.status_label.config(text=f"Error: {e}", foreground=self.colors['red'])


def def _save_logs_dialog(self):
        """Show a dialog to save logs to a file."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".log",
            filetypes=[("Log files", "*.log"), ("All files", "*.*")],
            title="Save Logs"
        )
        if file_path:
            self.save_logs(file_path)

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

    def main():
        """Start the GUI application."""
        root = tk.Tk()
        app = ProAVShokoGUI(root)
        root.mainloop()


if __name__ == "__main__":
    main()
