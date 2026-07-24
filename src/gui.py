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

    def __init__(self, root):
        self.root = root
        self.root.title("ProAV Shoko - USB Detective")
        self.root.geometry("1400x800")
        self.root.minsize(1200, 600)

        # Variables
        self.is_running = False
        self.log_queue = queue.Queue()
        self.usb_analyzer = None
        self.display_analyzer = None
        self.report_generator = None
        self.current_data = None
        self.platform_info = None

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

        # Build GUI
        self._build_gui()

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

        # === MIDDLE: Tree (left) + Log (right) ===
        middle_frame = ttk.Frame(main_frame)
        middle_frame.pack(fill=tk.BOTH, expand=True)

        # Left: USB tree with stability
        left_frame = ttk.LabelFrame(middle_frame, text="USB Tree & Stability", padding=10)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        self.tree_text = tk.Text(
            left_frame,
            bg=self.colors['bg_card'],
            fg=self.colors['fg'],
            font=('Courier New', 10),
            wrap=tk.NONE,
            relief=tk.FLAT,
            padx=10,
            pady=10
        )
        self.tree_text.pack(fill=tk.BOTH, expand=True)

        # Scrollbars for the tree
        tree_scroll_y = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.tree_text.yview)
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scroll_x = ttk.Scrollbar(left_frame, orient=tk.HORIZONTAL, command=self.tree_text.xview)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree_text.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)

        # Right: Live log
        right_frame = ttk.LabelFrame(middle_frame, text="Live Log", padding=10)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))

        self.log_text = tk.Text(
            right_frame,
            bg='#0a0a1a',
            fg='#88ccff',
            font=('Courier New', 9),
            wrap=tk.WORD,
            relief=tk.FLAT,
            padx=10,
            pady=10
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        log_scroll_y = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        log_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=log_scroll_y.set)

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
        pdf_btn.pack(side=tk.LEFT)

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
        """Run the analysis in the background."""
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

            # Update the platform label in the GUI thread
            self.root.after(0, self._update_platform_label)

            # 2. USB analysis - load hop limits from CSV if it exists
            config_path = Path("hop_limits.csv")
            if config_path.exists():
                self.usb_analyzer = USBAnalyzer(str(config_path))
                print(f"[+] Loaded hop limits from: {config_path}")
            else:
                self.usb_analyzer = USBAnalyzer()
                # Save default CSV
                self.usb_analyzer.save_hop_limits_csv("hop_limits.csv")
                print(f"[+] Created default hop_limits.csv")

            print("\n[+] Scanning USB devices...")
            usb_tree = self.usb_analyzer.build_tree()
            hops_data = self.usb_analyzer.calculate_hops_and_tiers(usb_tree)

            # 3. Stability assessment for all platforms
            stability = self.usb_analyzer.assess_stability(hops_data)

            # 4. Display information
            print("\n[+] Scanning displays...")
            self.display_analyzer = DisplayAnalyzer()
            displays = self.display_analyzer.get_display_info()

            # Save data
            self.current_data = {
                'usb_tree': usb_tree,
                'hops_data': hops_data,
                'stability': stability,
                'displays': displays,
                'platform_info': self.platform_info
            }

            # 5. Update the tree in the GUI
            self.root.after(0, self._update_tree_display, usb_tree, hops_data, stability, displays)

            # 6. Print the stability summary to the log
            print(self.usb_analyzer.get_stability_summary(stability))

            print("\n[+] Analysis complete!")
            self.root.after(0, lambda: self.status_label.config(
                text="Analysis complete",
                foreground=self.colors['green']
            ))

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

        # Stability heading
        self.tree_text.insert(tk.END, "\n")
        self.tree_text.insert(tk.END, "STABILITY VERDICT\n", ('header',))
        self.tree_text.insert(tk.END, "-" * 50 + "\n", ('header',))

        # Show all platforms grouped
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
                self.tree_text.insert(tk.END, f"max {v['max_hops']} hops\n", ('dim',))

        # Warnings
        warnings = [v for v in stability.get('verdicts', []) if not v['is_stable']]
        if warnings:
            self.tree_text.insert(tk.END, "\nWARNINGS:\n", ('warning_header',))
            for w in warnings:
                self.tree_text.insert(
                    tk.END,
                    f"  - {w['name']}: {w['warning']} (current hops: {w['current_hops']})\n",
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

        # Scroll to top
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
        """Update the log widget from the queue."""
        try:
            while True:
                text = self.log_queue.get_nowait()
                self.log_text.insert(tk.END, text)
                self.log_text.see(tk.END)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._update_log)

    def _reset_analysis(self):
        """Reset and restart the analysis."""
        if messagebox.askyesno("Reset", "Do you want to reset and restart the analysis?"):
            self._start_analysis()

    def _export_csv_limits(self):
        """Export hop limits as CSV."""
        if not self.usb_analyzer:
            messagebox.showwarning("No analysis", "Run an analysis first!")
            return

        default_name = f"hop_limits_{datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}.csv"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=default_name,
            title="Save hop limits as CSV"
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


def main():
    """Start the GUI application."""
    root = tk.Tk()
    app = ProAVShokoGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
