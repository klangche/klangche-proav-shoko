"""
GUI för ProAV Shōko - live-överblick med träd och logg
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
    """Omdirigerar stdout/stderr till GUI-logg."""

    def __init__(self, text_widget, queue):
        self.text_widget = text_widget
        self.queue = queue
        self.original_stdout = sys.stdout

    def write(self, text):
        """Skriv till både original stdout och GUI."""
        self.original_stdout.write(text)
        self.queue.put(text)

    def flush(self):
        self.original_stdout.flush()


class ProAVShokoGUI:
    """Huvud-GUI för ProAV Shōko."""

    def __init__(self, root):
        self.root = root
        self.root.title("🔍 ProAV Shōko - USB-detektiv")
        self.root.geometry("1400x800")
        self.root.minsize(1200, 600)

        # Variabler
        self.is_running = False
        self.log_queue = queue.Queue()
        self.usb_analyzer = None
        self.display_analyzer = None
        self.report_generator = None
        self.current_data = None
        self.platform_info = None

        # Färger
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

        # Sätt tema
        self.root.configure(bg=self.colors['bg'])
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TFrame', background=self.colors['bg'])
        style.configure('TLabelframe', background=self.colors['bg'], foreground=self.colors['fg'])
        style.configure('TLabelframe.Label', background=self.colors['bg'], foreground=self.colors['fg'])

        # Bygg GUI
        self._build_gui()

        # Starta logg-uppdatering
        self._update_log()

        # Börja analys direkt
        self.root.after(500, self._start_analysis)

    def _build_gui(self):
        """Bygg gränssnittet."""
        # Huvudram
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # === TOPP: Reset-knapp ===
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))

        # Plattformsinfo
        self.platform_label = ttk.Label(
            top_frame,
            text="🖥️ Laddar plattformsinfo...",
            font=('Segoe UI', 10),
            foreground=self.colors['blue']
        )
        self.platform_label.pack(side=tk.LEFT)

        # Reset-knapp
        reset_btn = tk.Button(
            top_frame,
            text="🔄 Reset",
            font=('Segoe UI', 10, 'bold'),
            bg=self.colors['blue'],
            fg='white',
            padx=15,
            pady=5,
            command=self._reset_analysis,
            cursor='hand2'
        )
        reset_btn.pack(side=tk.RIGHT)

        # === MITTEN: Träd (vänster) + Logg (höger) ===
        middle_frame = ttk.Frame(main_frame)
        middle_frame.pack(fill=tk.BOTH, expand=True)

        # Vänster: USB-träd med stabilitet
        left_frame = ttk.LabelFrame(middle_frame, text="🌳 USB-träd & Stabilitet", padding=10)
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

        # Scrollbars för träd
        tree_scroll_y = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.tree_text.yview)
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scroll_x = ttk.Scrollbar(left_frame, orient=tk.HORIZONTAL, command=self.tree_text.xview)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree_text.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)

        # Höger: Live-logg
        right_frame = ttk.LabelFrame(middle_frame, text="📋 Live-logg", padding=10)
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

        # === BOTTEN: Knappar för rapporter ===
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=(10, 0))

        # Status
        self.status_label = ttk.Label(
            bottom_frame,
            text="✅ Redo",
            font=('Segoe UI', 9),
            foreground=self.colors['green']
        )
        self.status_label.pack(side=tk.LEFT)

        # Knappar
        btn_frame = ttk.Frame(bottom_frame)
        btn_frame.pack(side=tk.RIGHT)

        html_btn = tk.Button(
            btn_frame,
            text="📄 Print HTML Report",
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
            text="📄 Print PDF Report",
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
        """Starta analys i bakgrundstråd."""
        if self.is_running:
            return

        self.is_running = True
        self.status_label.config(text="⏳ Analyserar...", foreground=self.colors['yellow'])
        self.log_text.delete(1.0, tk.END)
        self.tree_text.delete(1.0, tk.END)

        # Omdirigera stdout till logg
        sys.stdout = LogRedirector(self.log_text, self.log_queue)

        # Starta tråd
        thread = threading.Thread(target=self._run_analysis, daemon=True)
        thread.start()

    def _run_analysis(self):
        """Kör analys i bakgrunden."""
        try:
            # 1. Plattform
            self.platform_info = PlatformUtils.get_platform_info()
            print("=" * 60)
            print("  🔍 KANGCHE PROAV SHOKO - USB DETECTIVE")
            print("=" * 60)
            print(f"\n[+] Platform: {self.platform_info['os']} {self.platform_info['version']}")
            print(f"[+] Architecture: {self.platform_info['architecture']}")
            if self.platform_info['is_apple_silicon']:
                print("[+] Apple Silicon detected!")

            # Uppdatera plattformsetikett i GUI-tråden
            self.root.after(0, self._update_platform_label)

            # 2. USB-analys
            print("\n[+] Scanning USB devices...")
            self.usb_analyzer = USBAnalyzer()
            usb_tree = self.usb_analyzer.build_tree()
            hops_data = self.usb_analyzer.calculate_hops_and_tiers(usb_tree)
            stability = self.usb_analyzer.assess_stability(
                hops_data,
                self.platform_info['is_apple_silicon']
            )

            # 3. Skärminformation
            print("\n[+] Scanning displays...")
            self.display_analyzer = DisplayAnalyzer()
            displays = self.display_analyzer.get_display_info()

            # Spara data
            self.current_data = {
                'usb_tree': usb_tree,
                'hops_data': hops_data,
                'stability': stability,
                'displays': displays,
                'platform_info': self.platform_info
            }

            # 4. Uppdatera träd i GUI
            self.root.after(0, self._update_tree_display, usb_tree, hops_data, stability, displays)

            print("\n[+] Analysis complete!")
            self.root.after(0, lambda: self.status_label.config(
                text="✅ Analys klar",
                foreground=self.colors['green']
            ))

        except Exception as e:
            print(f"\n❌ Error: {e}")
            self.root.after(0, lambda: self.status_label.config(
                text=f"❌ Fel: {e}",
                foreground=self.colors['red']
            ))
        finally:
            self.is_running = False
            # Återställ stdout
            if hasattr(sys.stdout, 'original_stdout'):
                sys.stdout = sys.stdout.original_stdout

    def _update_platform_label(self):
        """Uppdatera plattformsetiketten."""
        if self.platform_info:
            text = f"🖥️ {self.platform_info['os']} {self.platform_info['architecture']}"
            if self.platform_info['is_apple_silicon']:
                text += " ⚠️ Apple Silicon"
            self.platform_label.config(text=text)

    def _update_tree_display(self, usb_tree, hops_data, stability, displays):
        """Uppdatera trädvisningen."""
        self.tree_text.delete(1.0, tk.END)

        # Stabilitetsrubrik
        self.tree_text.insert(tk.END, "\n")
        self.tree_text.insert(tk.END, "📊 STABILITY VERDICT\n", ('header',))
        self.tree_text.insert(tk.END, "-" * 50 + "\n", ('header',))

        # Färgad stabilitetsstatus
        color_map = {
            'green': 'stability_green',
            'yellow': 'stability_yellow',
            'orange': 'stability_orange',
            'red': 'stability_red'
        }
        tag = color_map.get(stability['color'], 'stability_green')
        self.tree_text.insert(tk.END, f"{stability['label']}\n", (tag,))

        if stability['warning']:
            self.tree_text.insert(tk.END, f"\n⚠️  {stability['warning']}\n", ('warning',))

        self.tree_text.insert(tk.END, f"\nMax Hops: {hops_data['max_hops']}\n", ('info',))
        self.tree_text.insert(tk.END, f"Total Tiers: {hops_data['max_tiers']}\n\n", ('info',))

        # USB-träd
        self.tree_text.insert(tk.END, "🌳 USB TREE STRUCTURE\n", ('header',))
        self.tree_text.insert(tk.END, "-" * 50 + "\n", ('header',))

        if usb_tree:
            self._render_tree_to_text(usb_tree, 0)
        else:
            self.tree_text.insert(tk.END, "  Inga USB-enheter hittades.\n", ('info',))

        # Skärmar
        self.tree_text.insert(tk.END, "\n🖥️ DISPLAY INFORMATION\n", ('header',))
        self.tree_text.insert(tk.END, "-" * 50 + "\n", ('header',))

        if displays:
            for display in displays:
                primary = " ⭐" if display.get('is_primary', False) else ""
                self.tree_text.insert(
                    tk.END,
                    f"  {display['resolution']}  {display['name']}{primary}\n",
                    ('info',)
                )
        else:
            self.tree_text.insert(tk.END, "  Inga skärmar hittades.\n", ('info',))

        # Konfigurera taggar för färger
        self.tree_text.tag_configure('header', font=('Courier New', 11, 'bold'), foreground=self.colors['blue'])
        self.tree_text.tag_configure('warning', font=('Courier New', 10, 'bold'), foreground=self.colors['orange'])
        self.tree_text.tag_configure('info', foreground=self.colors['fg'])
        self.tree_text.tag_configure('stability_green', foreground=self.colors['green'], font=('Courier New', 12, 'bold'))
        self.tree_text.tag_configure('stability_yellow', foreground=self.colors['yellow'], font=('Courier New', 12, 'bold'))
        self.tree_text.tag_configure('stability_orange', foreground=self.colors['orange'], font=('Courier New', 12, 'bold'))
        self.tree_text.tag_configure('stability_red', foreground=self.colors['red'], font=('Courier New', 12, 'bold'))

        # Flytta till toppen
        self.tree_text.see(1.0)

    def _render_tree_to_text(self, tree, level):
        """Rekursivt rendera USB-träd till textwidget."""
        indent = "  " * level
        for node in tree:
            is_hub = node.get('is_hub', False)
            icon = "📌" if is_hub else "🖥️"
            hub_tag = " [HUB]" if is_hub else ""
            hops = node['devpath'].count('/') if node.get('devpath') else 0

            line = f"{indent}{icon} {node.get('model', 'Okänd')}{hub_tag}  hops: {hops}\n"
            self.tree_text.insert(tk.END, line, ('info',))

            # Barn
            if node.get('children'):
                self._render_tree_to_text(node['children'], level + 1)

    def _update_log(self):
        """Uppdatera loggwidget från kö."""
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
        """Återställ och starta om analys."""
        if messagebox.askyesno("Återställ", "Vill du återställa och starta om analysen?"):
            self._start_analysis()

    def _print_html_report(self):
        """Generera och spara HTML-rapport."""
        if not self.current_data:
            messagebox.showwarning("Ingen data", "Kör först en analys!")
            return

        # Fråga om filnamn
        default_name = f"proav-shoko-{datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}.html"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".html",
            filetypes=[("HTML files", "*.html"), ("All files", "*.*")],
            initialfile=default_name,
            title="Spara HTML-rapport"
        )

        if not file_path:
            return

        try:
            self.status_label.config(text="⏳ Genererar HTML...", foreground=self.colors['yellow'])
            self.report_generator = ReportGenerator()
            html_path = self.report_generator.generate_html_report(
                self.current_data['usb_tree'],
                self.current_data['hops_data'],
                self.current_data['stability'],
                self.current_data['displays'],
                self.current_data['platform_info'],
                custom_path=file_path
            )
            self.status_label.config(text=f"✅ HTML sparad", foreground=self.colors['green'])
            print(f"\n[+] HTML Report saved: {html_path}")
            messagebox.showinfo("Klart", f"HTML-rapport sparad:\n{html_path}")
        except Exception as e:
            messagebox.showerror("Fel", f"Kunde inte generera HTML-rapport: {e}")
            self.status_label.config(text=f"❌ Fel: {e}", foreground=self.colors['red'])

    def _print_pdf_report(self):
        """Generera och spara PDF-rapport."""
        if not self.current_data:
            messagebox.showwarning("Ingen data", "Kör först en analys!")
            return

        # Fråga om filnamn
        default_name = f"proav-shoko-{datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}.pdf"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialfile=default_name,
            title="Spara PDF-rapport"
        )

        if not file_path:
            return

        try:
            self.status_label.config(text="⏳ Genererar PDF...", foreground=self.colors['yellow'])

            # Skapa temporär HTML
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
                custom_path=tmp_path
            )

            # Konvertera till PDF
            pdf_path = self.report_generator.generate_pdf_report(html_path, custom_path=file_path)

            # Städa bort temporär HTML
            try:
                Path(tmp_path).unlink()
            except:
                pass

            if pdf_path:
                self.status_label.config(text=f"✅ PDF sparad", foreground=self.colors['green'])
                print(f"\n[+] PDF Report saved: {pdf_path}")
                messagebox.showinfo("Klart", f"PDF-rapport sparad:\n{pdf_path}")
            else:
                messagebox.showwarning("Varning", "PDF-generering misslyckades. Kontrollera att weasyprint är installerat.")
        except Exception as e:
            messagebox.showerror("Fel", f"Kunde inte generera PDF-rapport: {e}")
            self.status_label.config(text=f"❌ Fel: {e}", foreground=self.colors['red'])


def main():
    """Starta GUI-applikationen."""
    root = tk.Tk()
    app = ProAVShokoGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
