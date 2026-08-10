#!/usr/bin/env python3
"""
ProAV Shoko - USB Detective GUI

Uses customtkinter to follow system dark/light theme.
Shows USB tree and live logs, matching CLI format exactly.
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import queue
import sys
import time
from datetime import datetime
from pathlib import Path

from src.usb_analyzer import USBAnalyzer
from src.display_analyzer import DisplayAnalyzer
from src.report_generator import ReportGenerator
from src.platform_utils import PlatformUtils

NARROW_WINDOW_THRESHOLD = 900


class LogRedirector:
    """Redirects stdout/stderr to the GUI log."""

    def __init__(self, text_widget, queue):
        self.text_widget = text_widget
        self.queue = queue
        self.original_stdout = sys.stdout

    def write(self, text):
        if self.original_stdout is not None:
            self.original_stdout.write(text)
        self.queue.put(text)

    def flush(self):
        if self.original_stdout is not None:
            self.original_stdout.flush()


class ProAVShokoGUI:
    """Main GUI for ProAV Shoko."""

    def __init__(self, root, csv_path=None):
        self.root = root
        self.root.title("ProAV Shoko - USB Detective")

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = 1400
        window_height = 800
        x_position = (screen_width - window_width) // 2
        y_position = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
        self.root.minsize(700, 500)

        self.is_running = False
        self.is_monitoring = False
        self.log_queue = queue.Queue()
        self.ui_queue = queue.Queue()
        self.usb_analyzer = None
        self.display_analyzer = None
        self.report_generator = None
        self.current_data = None
        self.platform_info = None
        self.log_panel_visible = True
        self.csv_path = csv_path
        self.recent_disconnects = {}
        self.event_log = []
        self.jitter_warned_models = set()

        self._build_gui()
        self.root.bind('<Configure>', self._on_window_resize)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._update_log()
        self._process_ui_queue()

    def _build_gui(self):
        main_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        main_frame.pack(fill=ctk.BOTH, expand=True, padx=20, pady=20)

        top_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        top_frame.pack(fill=ctk.X, pady=(0, 15))

        self.platform_label = ctk.CTkLabel(
            top_frame, text="Click Start to begin", font=('Segoe UI', 12)
        )
        self.platform_label.pack(side=ctk.LEFT)

        btn_right_frame = ctk.CTkFrame(top_frame, fg_color="transparent")
        btn_right_frame.pack(side=ctk.RIGHT)

        self.reset_btn = ctk.CTkButton(
            btn_right_frame, text="Reset",
            font=('Segoe UI', 11, 'bold'),
            command=self._reset_analysis, width=100
        )

        self.stop_btn = ctk.CTkButton(
            btn_right_frame, text="Stop",
            font=('Segoe UI', 11, 'bold'),
            fg_color="#ef4444", hover_color="#dc2626",
            command=self._on_stop_clicked, width=100
        )

        self.start_btn = ctk.CTkButton(
            btn_right_frame, text="Start Analysis",
            font=('Segoe UI', 11, 'bold'),
            fg_color="#10b981", hover_color="#059669",
            command=self._on_start_clicked, width=130
        )
        self.start_btn.pack(side=ctk.RIGHT, padx=(10, 0))

        middle_container = ctk.CTkFrame(main_frame, fg_color="transparent")
        middle_container.pack(fill=ctk.BOTH, expand=True, pady=(0, 15))

        self.left_frame = ctk.CTkFrame(middle_container)
        self.left_frame.pack(side=ctk.LEFT, fill=ctk.BOTH, expand=True, padx=(0, 10))

        log_header = ctk.CTkFrame(self.left_frame, height=40, fg_color="transparent")
        log_header.pack(fill=ctk.X)
        log_header_label = ctk.CTkLabel(
            log_header, text="Live Log", font=('Segoe UI', 13, 'bold')
        )
        log_header_label.pack(side=ctk.LEFT, padx=15, pady=8)

        log_frame = ctk.CTkFrame(self.left_frame, fg_color="transparent")
        log_frame.pack(fill=ctk.BOTH, expand=True, padx=15, pady=15)

        self.log_text = ctk.CTkTextbox(
            log_frame, font=('Consolas', 10), wrap=ctk.WORD
        )
        self.log_text.pack(fill=ctk.BOTH, expand=True)
        self.log_text.configure(state='disabled')

        self.right_frame = ctk.CTkFrame(middle_container)
        self.right_frame.pack(side=ctk.RIGHT, fill=ctk.BOTH, expand=True)

        tree_header = ctk.CTkFrame(self.right_frame, height=40, fg_color="transparent")
        tree_header.pack(fill=ctk.X)
        tree_header_label = ctk.CTkLabel(
            tree_header, text="Tree layout", font=('Segoe UI', 13, 'bold')
        )
        tree_header_label.pack(side=ctk.LEFT, padx=15, pady=8)

        tree_frame = ctk.CTkFrame(self.right_frame, fg_color="transparent")
        tree_frame.pack(fill=ctk.BOTH, expand=True, padx=15, pady=15)

        self.tree_text = ctk.CTkTextbox(
            tree_frame, font=('Consolas', 11), wrap=ctk.WORD
        )
        self.tree_text.pack(fill=ctk.BOTH, expand=True)
        self.tree_text.configure(state='disabled')

        self.log_text.tag_config('event_connect', foreground="#22c55e")
        self.log_text.tag_config('event_disconnect', foreground="#fb923c")
        self.log_text.tag_config('timestamp', foreground="#94a3b8")

        bottom_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        bottom_frame.pack(fill=ctk.X)

        self.status_label = ctk.CTkLabel(
            bottom_frame, text="Ready", font=('Segoe UI', 11)
        )
        self.status_label.pack(side=ctk.LEFT)

        btn_frame = ctk.CTkFrame(bottom_frame, fg_color="transparent")
        btn_frame.pack(side=ctk.RIGHT)

        self.room_entry = ctk.CTkEntry(
            btn_frame, width=120, font=('Segoe UI', 11),
            placeholder_text="Room name"
        )
        self.room_entry.pack(side=ctk.LEFT, padx=(0, 8))

        pdf_btn = ctk.CTkButton(
            btn_frame, text="PDF Report",
            font=('Segoe UI', 11, 'bold'),
            fg_color="#dc2626", hover_color="#b91c1c",
            command=self._generate_pdf_report, width=130
        )
        pdf_btn.pack(side=ctk.LEFT, padx=(0, 8))

        html_btn = ctk.CTkButton(
            btn_frame, text="HTML Report",
            font=('Segoe UI', 11, 'bold'),
            fg_color="#2563eb", hover_color="#1d4ed8",
            command=self._generate_html_report, width=130
        )
        html_btn.pack(side=ctk.LEFT)

    def _on_window_resize(self, event):
        if event.widget != self.root:
            return
        width = self.root.winfo_width()
        should_show = width >= NARROW_WINDOW_THRESHOLD
        if should_show and not self.log_panel_visible:
            self.right_frame.pack(side=ctk.RIGHT, fill=ctk.BOTH, expand=True)
            self.left_frame.pack(side=ctk.LEFT, fill=ctk.BOTH, expand=True, padx=(0, 10))
            self.log_panel_visible = True
        elif not should_show and self.log_panel_visible:
            self.right_frame.pack_forget()
            self.log_panel_visible = False

    def _on_start_clicked(self):
        self.start_btn.pack_forget()
        self.stop_btn.pack(side=ctk.RIGHT, padx=(10, 0))
        self.reset_btn.pack(side=ctk.RIGHT)
        self._start_analysis()

    def _on_stop_clicked(self):
        self._stop_monitoring()
        self.stop_btn.pack_forget()
        self.status_label.configure(text="Stopped", text_color="#ef4444")

    def _start_analysis(self):
        if self.is_running:
            return
        self.is_running = True
        self.status_label.configure(text="Analyzing...", text_color="#f59e0b")
        self.log_text.configure(state='normal')
        self.log_text.delete('0.0', "end")
        self.log_text.configure(state='disabled')
        self.tree_text.configure(state='normal')
        self.tree_text.delete('0.0', "end")
        self.tree_text.configure(state='disabled')

        sys.stdout = LogRedirector(self.log_text, self.log_queue)

        thread = threading.Thread(target=self._run_analysis, daemon=True)
        thread.start()

    def _run_analysis(self):
        try:
            self.platform_info = PlatformUtils.get_platform_info()
            print(f"[+]System OS: {self.platform_info['os']}")
            print(f"[+]OS version: {self.platform_info['version']}")
            self.ui_queue.put(self._update_platform_label)

            self.usb_analyzer = USBAnalyzer(self.csv_path) if self.csv_path else USBAnalyzer()
            print("[+]Loading Connected USB Devices...")

            self.display_analyzer = DisplayAnalyzer()
            print("[+]Loading Connected displays...")
            print()

            self._refresh_tree_and_stability()

            print("[+]Live Logging")
            self.ui_queue.put(lambda: self.status_label.configure(
                text="Monitoring (live)", text_color="#10b981"
            ))

            self.usb_analyzer.start_live_monitoring(
                on_connect=self._on_device_connect,
                on_disconnect=self._on_device_disconnect,
                check_every_seconds=0.1
            )
            self.is_monitoring = True

        except Exception as e:
            err_msg = str(e)
            self.ui_queue.put(lambda m=err_msg: self._log_event(f"Error: {m}", None))
            self.ui_queue.put(lambda m=err_msg: self.status_label.configure(
                text=f"Error: {m}", text_color="#ef4444"
            ))
        finally:
            self.is_running = False
            if hasattr(sys.stdout, 'original_stdout'):
                sys.stdout = sys.stdout.original_stdout

    def _refresh_tree_and_stability(self):
        usb_tree = self.usb_analyzer.build_tree()
        hops_data = self.usb_analyzer.calculate_hops_and_tiers(usb_tree)
        stability = self.usb_analyzer.assess_stability(hops_data, usb_tree)

        displays = self.display_analyzer.get_display_info() if self.display_analyzer else []

        self.current_data = {
            'usb_tree': usb_tree,
            'hops_data': hops_data,
            'stability': stability,
            'displays': displays,
            'platform_info': self.platform_info
        }

        self.ui_queue.put(lambda: self._update_tree_display(usb_tree, hops_data, stability, displays))
        self._log_external_ports(usb_tree, stability)

    def _log_external_ports(self, usb_tree, stability):
        root_orig = usb_tree[0] if usb_tree else {}
        orig_children = [c for c in root_orig.get('children', []) if not c.get('is_display')]
        ports_data = stability.get('ports', [])

        warned_ports = []
        for idx, child in enumerate(orig_children):
            if child.get('is_internal', False):
                continue
            port_info = next((p for p in ports_data if p.get('id') == idx + 1), None)
            if not port_info:
                continue
            port_warnings = [v for v in port_info.get('verdicts', []) if v.get('warning')]
            if port_warnings:
                warned_ports.append((idx, child, port_info, port_warnings))

        if not warned_ports:
            return

        print("[+]Warnings")

        for idx, child, port_info, port_warnings in warned_ports:
            port_tag = f"external.port{idx + 1}"
            label = port_info.get('label', child.get('model', 'Port'))
            dc = len(port_info.get('devices', []))
            ph = port_info.get('max_hops', 0)
            pt = port_info.get('max_tiers', 0)
            p_hub = port_info.get('external_hubs', 0)
            ep_label = "endpoint" if dc == 1 else "endpoints"

            print(f"{port_tag}.tree")
            print(f"  {label} ({dc} {ep_label}, hops={ph}, tiers={pt}, hubs={p_hub})")

            children = child.get('children', [])
            if children:
                self._print_tree_stdout(children, "    ")

            print()
            print(f"{port_tag}.warnings")
            for w in port_warnings:
                print(f"    ! {w['name']}: {w['warning']}")
            print()

    def _print_tree_stdout(self, nodes, prefix="", _show_internal=False, _parent_is_internal=False):
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
            print(f"{prefix}{connector}{badge_str}{label}{port_info_str}")

            if node.get('children'):
                child_prefix = prefix + ("    " if is_last else "│   ")
                new_parent_int = _parent_is_internal or node.get('is_internal', False)
                self._print_tree_stdout(node['children'], child_prefix, _show_internal, new_parent_int)

    def _on_device_connect(self, device_id, device_info):
        now = time.time()
        model = device_info.get('ID_MODEL', device_id)
        self.event_log.append((model, 'connect', now))
        cutoff = now - 30.0
        self.event_log = [e for e in self.event_log if e[2] >= cutoff]

        last_disc = self.recent_disconnects.pop(model, 0.0)
        if last_disc and (now - last_disc) <= 5.0:
            msg = f"RE-CONNECTED: {model}"
        else:
            msg = f"CONNECTED: {model}"
        self.ui_queue.put(lambda m=msg: self._log_event(m, 'event_connect'))

        if not self._jitter_warned(model, now):
            self._check_jitter(model, now)

    def _on_device_disconnect(self, device_id, device_info):
        now = time.time()
        model = device_info.get('ID_MODEL', device_id)
        self.recent_disconnects[model] = now
        self.event_log.append((model, 'disconnect', now))
        cutoff = now - 30.0
        self.event_log = [e for e in self.event_log if e[2] >= cutoff]

        self.ui_queue.put(lambda m=model: self._log_event(f"DISCONNECTED: {m}", 'event_disconnect'))

        if not self._jitter_warned(model, now):
            self._check_jitter(model, now)

    def _jitter_warned(self, model, now):
        return model in self.jitter_warned_models

    def _check_jitter(self, model, now):
        cutoff = now - 30.0
        recent = [e for e in self.event_log if e[2] >= cutoff and e[0] == model]
        connects = sum(1 for _, t, _ in recent if t == 'connect')
        disconnects = sum(1 for _, t, _ in recent if t == 'disconnect')
        if connects >= 2 and disconnects >= 2:
            self.jitter_warned_models.add(model)
            self.ui_queue.put(lambda m=model: self._log_event(
                f"JITTER: {m} — rapid connect/disconnect", 'event_disconnect'))

    def _log_event(self, message, tag):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state='normal')
        self.log_text.insert("end", f"[{timestamp}] ", ("timestamp",))
        tags = (tag,) if tag else ()
        self.log_text.insert("end", f"{message}\n", tags)
        self.log_text.configure(state='disabled')
        self.log_text.see("end")

    def _update_platform_label(self):
        if self.platform_info:
            text = f"{self.platform_info['os']} {self.platform_info['architecture']}"
            if self.platform_info['is_apple_silicon']:
                text += " (Apple Silicon)"
            self.platform_label.configure(text=text)

    def _update_tree_display(self, usb_tree, hops_data, stability, displays):
        self.tree_text.configure(state='normal')
        self.tree_text.delete('0.0', "end")

        self._print_section_header("Full USB & Display Tree")
        self._print_tag("overall.tree")

        if usb_tree:
            root_node = usb_tree[0]
            if displays:
                for d in displays:
                    prim = " (Primary)" if d.get('is_primary', False) else ""
                    int_disp = d.get('is_internal', False)
                    root_node.setdefault('children', []).append({
                        'model': f"{d['resolution']}  {d['name']}{prim}",
                        'name': d['name'], 'children': [], 'hops': 1,
                        'is_hub': False, 'is_internal': int_disp, 'is_display': True, 'port': 0
                    })
            self._print_tree(usb_tree, "  ", _show_internal=True)
        else:
            self.tree_text.insert("end", "  No USB devices found.\n")

        self._print_section_header("Overall rating")
        self._print_tag("overall.verdict")
        for v in stability.get('verdicts', []):
            self._print_verdict(v)

        overall_warnings = [v for v in stability.get('verdicts', []) if v.get('warning')]
        if overall_warnings:
            self._print_tag("overall.warnings")
            for w in overall_warnings:
                self.tree_text.insert("end", f"    ! {w['name']}: {w['warning']}\n")
        self.tree_text.insert("end", "\n")
        self.tree_text.insert("end", "\n")
        self.tree_text.insert("end", "PER PORT" + "=" * 31)
        self.tree_text.insert("end", "\n\n")

        ports_data = stability.get('ports', [])
        root_orig = usb_tree[0] if usb_tree else {}
        orig_children = [c for c in root_orig.get('children', []) if not c.get('is_display')]

        sep = "  " + "- " * 35

        def _print_port_label(child, idx):
            port_info = next((p for p in ports_data if p.get('id') == idx + 1), None)
            label = port_info['label'] if port_info else child.get('model', 'Port')
            dc = len(port_info['devices']) if port_info else 0
            ph = port_info['max_hops'] if port_info else 0
            pt = port_info['max_tiers'] if port_info else 0
            p_hub = port_info.get('external_hubs', 0) if port_info else 0
            is_int = child.get('is_internal', False)
            int_pre = "[INTERNAL] " if is_int else ""
            ep_label = "endpoint" if dc == 1 else "endpoints"
            self.tree_text.insert("end", f"  {int_pre}{label} ({dc} {ep_label}, hops={ph}, tiers={pt}, hubs={p_hub})\n")
            children = child.get('children', [])
            if children:
                self._print_tree(children, "    ")

        def print_section(header, is_internal_filter, tag_prefix):
            self.tree_text.insert("end", header + "-" * 31 + "\n")
            self._print_tag(f"{tag_prefix}.section")
            first = True
            for idx, child in enumerate(orig_children):
                if not is_internal_filter(child):
                    continue
                if not first:
                    self.tree_text.insert("end", "\n")
                port_tag = f"{tag_prefix}.port{idx + 1}"
                self._print_tag(f"{port_tag}.tree")
                _print_port_label(child, idx)
                self.tree_text.insert("end", "\n")
                port_info = next((p for p in ports_data if p.get('id') == idx + 1), None)
                if port_info and tag_prefix != "internal":
                    self._print_tag(f"{port_tag}.verdict")
                    self._print_stability_port(port_info, port_info['verdicts'])
                    port_warnings = [v for v in port_info['verdicts'] if v.get('warning')]
                    if port_warnings:
                        self._print_tag(f"{port_tag}.warnings")
                        for w in port_warnings:
                            self.tree_text.insert("end", f"    ! {w['name']}: {w['warning']}\n")
                elif tag_prefix != "internal":
                    self.tree_text.insert("end", "    No port data available.\n")
                self.tree_text.insert("end", "\n")
                self.tree_text.insert("end", sep + "\n")
                first = False

        print_section("EXTERNAL", lambda c: not c.get('is_internal', False), "external")
        print_section("INTERNAL", lambda c: c.get('is_internal', False), "internal")
        self.tree_text.configure(state='disabled')

    def _print_tree(self, nodes, prefix="", _show_internal=False, _parent_is_internal=False):
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
            self.tree_text.insert("end", f"{prefix}{connector}{badge_str}{label}{port_info_str}\n")

            if node.get('children'):
                child_prefix = prefix + ("    " if is_last else "│   ")
                new_parent_int = _parent_is_internal or node.get('is_internal', False)
                self._print_tree(node['children'], child_prefix, _show_internal, new_parent_int)

    @staticmethod
    def _node_label(node):
        model = node.get('model', node.get('name', 'Unknown'))
        device_info = node.get('device_info', '')
        iface_desc = node.get('interface_desc', '')
        iface_num = node.get('interface_number')
        if node.get('is_composite_interface'):
            mi = f"MI_{iface_num:02d}" if iface_num is not None else ""
            suffix = f" ({device_info})" if device_info else ""
            if model and 'USB-enhet' not in model and 'sammansatt' not in model and 'Composite' not in model:
                label = model
                if mi:
                    label += f" {mi}"
                return f"{label}{suffix}"
            if iface_desc:
                iface_tag = "HID Keyboard" if "Keyboard" in iface_desc else \
                            "HID Mouse" if "Mouse" in iface_desc else \
                            iface_desc
                return f"{iface_tag} {mi}{suffix}".strip()
        return f"{model} ({device_info})" if device_info else model

    def _print_tag(self, tag: str):
        self.tree_text.insert("end", f"\n    {tag}\n")

    def _print_stability_port(self, port_info, verdicts):
        for v in verdicts:
            self._print_verdict(v)

    def _print_verdict(self, v):
        status_char = '+' if v['color'] == 'green' else ('~' if v['color'] == 'orange' else '!')
        hubs_str = f"hubs {v['current_hubs']}/{v['max_hubs']}  " if 'current_hubs' in v else ""
        desc = v.get('description', v.get('name', ''))
        self.tree_text.insert("end",
            f"    {status_char} {desc:<22s} "
            f"{v['status']:<9s} "
            f"hops {v['current_hops']}/{v['max_hops']}  "
            f"tiers {v['current_tiers']}/{v['max_tiers']}  "
            f"{hubs_str}\n")

    def _print_section_header(self, title):
        self.tree_text.insert("end", f"\n{title}\n")
        self.tree_text.insert("end", "-" * 70)
        self.tree_text.insert("end", "\n")

    def _update_log(self):
        try:
            while True:
                text = self.log_queue.get_nowait()
                self.log_text.configure(state='normal')
                self.log_text.insert("end", text)
                self.log_text.configure(state='disabled')
                self.log_text.see("end")
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._update_log)

    def _process_ui_queue(self):
        """Run UI updates submitted from background threads on the main thread.

        Tkinter is not thread-safe, and calling root.after() from a worker
        thread silently drops the callback on macOS. All worker-thread UI
        work is pushed here as callables and executed on the main thread.
        """
        try:
            while True:
                callback = self.ui_queue.get_nowait()
                try:
                    callback()
                except Exception as e:
                    print(f"UI update error: {e}")
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._process_ui_queue)

    def _stop_monitoring(self):
        if self.is_monitoring and self.usb_analyzer:
            self.usb_analyzer.stop_live_monitoring()
            self.is_monitoring = False

    def _reset_analysis(self):
        if messagebox.askyesno("Reset", "Do you want to reset and restart the analysis?"):
            self._stop_monitoring()
            self.stop_btn.pack_forget()
            self.start_btn.pack(side=ctk.RIGHT, padx=(10, 0))
            self._start_analysis()

    def _on_close(self):
        self._stop_monitoring()
        self.root.destroy()

    def _generate_html_report(self):
        if not self.current_data:
            messagebox.showwarning("No data", "Run an analysis first!")
            return

        room_name = self.room_entry.get().strip()
        ts = datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
        room_part = f"{room_name}-" if room_name else ""
        default_name = f"proav-shoko-{room_part}{ts}.html"
        file_path = filedialog.asksaveasfilename(
            defaultextension='.html',
            filetypes=[("HTML files", "*.html"), ("All files", "*.*")],
            initialfile=default_name,
            title="Save HTML Report"
        )
        if not file_path:
            return

        try:
            self.status_label.configure(text="Generating HTML...", text_color="#f59e0b")
            self.report_generator = ReportGenerator(room_name=room_name)
            html_path = self.report_generator.generate_html_report(
                self.current_data['usb_tree'],
                self.current_data['hops_data'],
                self.current_data['stability'],
                self.current_data['displays'],
                self.current_data['platform_info'],
                platform_notes=self.usb_analyzer.get_platform_notes() if self.usb_analyzer else None,
                custom_path=file_path
            )
            self.status_label.configure(text="HTML saved", text_color="#10b981")
            print(f"\n[+] HTML Report saved: {html_path}")
            messagebox.showinfo("Done", f"HTML report saved:\n{html_path}")
            self.report_generator.open_report(html_path)
        except Exception as e:
            messagebox.showerror("Error", f"Could not generate HTML report: {e}")
            self.status_label.configure(text=f"Error: {e}", text_color="#ef4444")

    def _generate_pdf_report(self):
        if not self.current_data:
            messagebox.showwarning("No data", "Run an analysis first!")
            return

        room_name = self.room_entry.get().strip()
        ts = datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
        room_part = f"{room_name}-" if room_name else ""
        default_name = f"proav-shoko-{room_part}{ts}.pdf"
        file_path = filedialog.asksaveasfilename(
            defaultextension='.pdf',
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialfile=default_name,
            title="Save PDF Report"
        )
        if not file_path:
            return

        try:
            self.status_label.configure(text="Generating PDF...", text_color="#f59e0b")
            self.report_generator = ReportGenerator(room_name=room_name)

            pdf_path = self.report_generator.generate_pdf_report(
                custom_path=file_path,
                usb_tree=self.current_data['usb_tree'],
                hops_data=self.current_data['hops_data'],
                stability=self.current_data['stability'],
                displays=self.current_data['displays'],
                platform_info=self.current_data['platform_info'],
                platform_notes=self.usb_analyzer.get_platform_notes() if self.usb_analyzer else None,
            )

            if pdf_path:
                self.status_label.configure(text="PDF saved", text_color="#10b981")
                print(f"\n[+] PDF Report saved: {pdf_path}")
                messagebox.showinfo("Done", f"PDF report saved:\n{pdf_path}")
                self.report_generator.open_report(pdf_path)
            else:
                self.status_label.configure(text="PDF failed", text_color="#ef4444")

        except Exception as e:
            messagebox.showerror("Error", f"Could not generate PDF report: {e}")
            self.status_label.configure(text=f"Error: {e}", text_color="#ef4444")


def main():
    """Start the GUI application."""
    ctk.set_appearance_mode("system")
    ctk.set_default_color_theme("blue")
    root = ctk.CTk()
    app = ProAVShokoGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
