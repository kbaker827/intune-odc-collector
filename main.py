#!/usr/bin/env python3
"""
Intune ODC Log Collector - Native Python Implementation

Collects Intune diagnostic logs directly in Python without PowerShell.
Downloads Intune.XML and processes it natively.
"""

import glob
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime

import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk


def _hostname():
    """Return this machine's name with safe fallbacks."""
    return os.environ.get('COMPUTERNAME') or os.environ.get('HOSTNAME') or 'Unknown'


class ODCLogCollector:
    """Main application class."""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Intune ODC Log Collector")
        self.root.geometry("750x600")
        self.root.minsize(750, 500)

        if sys.platform == 'win32':
            try:
                from ctypes import windll
                windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                pass

        self.is_running = False
        self.log_dir = r"C:\IntuneODCLogs"
        self.result_dir = None
        self.setup_ui()

    # ------------------------------------------------------------------ #
    # UI construction                                                      #
    # ------------------------------------------------------------------ #

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        self._create_header(main_frame)
        self._create_info_section(main_frame)
        self._create_mode_section(main_frame)
        self._create_actions_section(main_frame)
        self._create_progress_section(main_frame)
        self._create_output_section(main_frame)

    def _create_header(self, parent):
        header = ttk.Frame(parent)
        header.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(
            header,
            text="Intune ODC Log Collector",
            font=('Segoe UI', 16, 'bold'),
        ).pack(anchor=tk.W)

        ttk.Label(
            header,
            text="Collect diagnostic logs for Microsoft Intune troubleshooting",
            font=('Segoe UI', 9),
        ).pack(anchor=tk.W)

        ttk.Separator(parent, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=5)

    def _create_info_section(self, parent):
        info_frame = ttk.LabelFrame(parent, text="What This Tool Does", padding="10")
        info_frame.pack(fill=tk.X, pady=(0, 10))

        info_text = (
            "This tool collects Intune One Data Collector (ODC) logs which are used by "
            "Microsoft Support to diagnose Intune issues.\n\n"
            "Choose collection method:\n"
            "• Native Python - Uses built-in Python collection (recommended, faster)\n"
            "• Microsoft Tool - Downloads and runs official Microsoft PowerShell script\n\n"
            "The process takes ~10 minutes and creates a compressed ZIP file with all data.\n\n"
            "Note: This tool must be run as Administrator."
        )

        ttk.Label(
            info_frame,
            text=info_text,
            font=('Segoe UI', 9),
            justify=tk.LEFT,
            wraplength=650,
        ).pack(anchor=tk.W)

    def _create_mode_section(self, parent):
        mode_frame = ttk.LabelFrame(parent, text="Collection Mode", padding="10")
        mode_frame.pack(fill=tk.X, pady=(0, 10))

        self.collection_mode = tk.StringVar(value='native')

        native_frame = ttk.Frame(mode_frame)
        native_frame.pack(fill=tk.X, pady=2)
        ttk.Radiobutton(
            native_frame, text="Native Python",
            variable=self.collection_mode, value='native',
        ).pack(side=tk.LEFT)
        ttk.Label(
            native_frame,
            text="(Recommended - Faster, works offline with cached XML)",
            font=('Segoe UI', 8), foreground='green',
        ).pack(side=tk.LEFT, padx=(5, 0))

        ms_frame = ttk.Frame(mode_frame)
        ms_frame.pack(fill=tk.X, pady=2)
        ttk.Radiobutton(
            ms_frame, text="Microsoft Tool",
            variable=self.collection_mode, value='microsoft',
        ).pack(side=tk.LEFT)
        ttk.Label(
            ms_frame,
            text="(Uses official Microsoft PowerShell script)",
            font=('Segoe UI', 8), foreground='gray',
        ).pack(side=tk.LEFT, padx=(5, 0))

        self.cache_xml_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            mode_frame,
            text="Cache Intune.XML locally (faster subsequent runs)",
            variable=self.cache_xml_var,
        ).pack(anchor=tk.W, pady=(5, 0))

    def _create_actions_section(self, parent):
        action_frame = ttk.Frame(parent)
        action_frame.pack(fill=tk.X, pady=(0, 10))

        self.collect_btn = tk.Button(
            action_frame, text="Start", command=self.start_collection,
            bg='#007bff', fg='white', disabledforeground='white',
            font=('Segoe UI', 11, 'bold'), width=10, cursor='hand2',
        )
        self.collect_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.open_btn = tk.Button(
            action_frame, text="Open Folder", command=self.open_log_folder,
            bg='#6c757d', fg='white', disabledforeground='white',
            font=('Segoe UI', 10), width=12, cursor='hand2', state=tk.DISABLED,
        )
        self.open_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.cancel_btn = tk.Button(
            action_frame, text="Cancel", command=self.cancel_collection,
            bg='#dc3545', fg='white', disabledforeground='white',
            font=('Segoe UI', 10), width=10, cursor='hand2', state=tk.DISABLED,
        )
        self.cancel_btn.pack(side=tk.LEFT)

    def _create_progress_section(self, parent):
        progress_frame = ttk.LabelFrame(parent, text="Progress", padding="10")
        progress_frame.pack(fill=tk.X, pady=(0, 10))

        self.progress_var = tk.DoubleVar(value=0)
        ttk.Progressbar(
            progress_frame, variable=self.progress_var,
            maximum=100, mode='determinate',
        ).pack(fill=tk.X, pady=(0, 5))

        self.time_var = tk.StringVar(value="Estimated time: ~10 minutes")
        ttk.Label(progress_frame, textvariable=self.time_var, font=('Segoe UI', 9)).pack(anchor=tk.W)

        self.status_var = tk.StringVar(value="Ready to start")
        ttk.Label(
            progress_frame, textvariable=self.status_var,
            font=('Segoe UI', 9, 'bold'),
        ).pack(anchor=tk.W, pady=(5, 0))

    def _create_output_section(self, parent):
        output_frame = ttk.LabelFrame(parent, text="Output", padding="10")
        output_frame.pack(fill=tk.BOTH, expand=True)

        self.output_text = scrolledtext.ScrolledText(
            output_frame, wrap=tk.WORD, font=('Consolas', 9),
            height=12, bg='#f5f5f5',
        )
        self.output_text.pack(fill=tk.BOTH, expand=True)

        self._log("Ready to collect Intune ODC logs.")
        self._log("Click 'Start' to begin.")
        self._log("")
        self._log("Note: This tool must be run as Administrator.")

    # ------------------------------------------------------------------ #
    # Thread-safe UI helpers                                               #
    # ------------------------------------------------------------------ #

    def _log(self, message):
        """Append a timestamped line to the output pane (thread-safe)."""
        def _do():
            self.output_text.config(state=tk.NORMAL)
            self.output_text.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {message}\n")
            self.output_text.see(tk.END)
            self.output_text.config(state=tk.DISABLED)
        self.root.after(0, _do)

    def _update_status(self, message):
        """Update the status label and log the message (thread-safe)."""
        self.root.after(0, self.status_var.set, message)
        self._log(message)

    def _set_progress(self, value):
        """Update the progress bar (thread-safe)."""
        self.root.after(0, self.progress_var.set, value)

    def _ui_call(self, func, *args, **kwargs):
        """Schedule a GUI call on the main thread (thread-safe)."""
        self.root.after(0, lambda: func(*args, **kwargs))

    # ------------------------------------------------------------------ #
    # XML helpers                                                          #
    # ------------------------------------------------------------------ #

    @staticmethod
    def _find_xml_elems(parent, tag, ns_map):
        """Return all elements matching tag, handling an optional XML namespace."""
        if ns_map:
            elems = list(parent.iter(f'{{{ns_map["ns"]}}}{tag}'))
            if elems:
                return elems
        return list(parent.iter(tag))

    @staticmethod
    def _find_child(parent, tag_name, ns_map):
        """Return the first direct child matching tag_name (namespace-aware)."""
        if ns_map:
            child = parent.find(f'ns:{tag_name}', ns_map)
            if child is not None:
                return child
        child = parent.find(tag_name)
        if child is not None:
            return child
        for child in parent:
            if child.tag.endswith(tag_name):
                return child
        return None

    # ------------------------------------------------------------------ #
    # Admin check                                                          #
    # ------------------------------------------------------------------ #

    def is_admin(self):
        """Return True if the process is running with administrator privileges."""
        try:
            import ctypes
            return bool(ctypes.windll.shell32.IsUserAnAdmin())
        except Exception:
            return False

    # ------------------------------------------------------------------ #
    # Download                                                             #
    # ------------------------------------------------------------------ #

    def _download_file(self, url, dest, timeout=60):
        """Download *url* to *dest* with a connect/read timeout."""
        with urllib.request.urlopen(url, timeout=timeout) as response:
            with open(dest, 'wb') as f:
                shutil.copyfileobj(response, f)

    def download_xml(self):
        """Download Intune.XML from GitHub, using a local cache when fresh."""
        xml_path = os.path.join(self.log_dir, "Intune.xml")
        cache_path = os.path.join(self.log_dir, "Intune.xml.cached")
        url = "https://raw.githubusercontent.com/markstan/IntuneOneDataCollector/master/Intune.xml"

        if self.cache_xml_var.get() and os.path.exists(cache_path):
            cache_age = time.time() - os.path.getmtime(cache_path)
            if cache_age < 7 * 24 * 3600:
                self._log(f"Using cached Intune.XML (age: {cache_age / 3600:.1f} hours)")
                shutil.copy(cache_path, xml_path)
                return xml_path
            self._log("Cache expired, downloading fresh copy...")

        self._update_status("Downloading Intune.XML...")
        try:
            self._download_file(url, xml_path)
            self._log(f"Downloaded: {xml_path}")
            if self.cache_xml_var.get():
                shutil.copy(xml_path, cache_path)
                self._log("Cached Intune.XML for future use")
            return xml_path
        except Exception as e:
            self._log(f"Error downloading XML: {e}")
            if os.path.exists(cache_path):
                self._log("Using expired cached XML as fallback...")
                shutil.copy(cache_path, xml_path)
                return xml_path
            raise

    # ------------------------------------------------------------------ #
    # XML parsing                                                          #
    # ------------------------------------------------------------------ #

    def parse_xml(self, xml_path):
        """Parse Intune.XML and return a list of (pkg_id, element) tuples."""
        self._update_status("Parsing Intune.XML...")
        try:
            root = ET.parse(xml_path).getroot()
            self._log(f"XML Root: {root.tag}")

            ns = {}
            if '}' in root.tag:
                ns_uri = root.tag.split('}')[0].strip('{')
                ns = {'ns': ns_uri}
                self._log(f"Detected namespace: {ns_uri}")

            packages = []
            if ns:
                packages = [(p.get('ID', 'Unknown'), p) for p in root.findall('.//ns:Package', ns)]
            if not packages:
                packages = [(p.get('ID', 'Unknown'), p) for p in root.findall('.//Package')]
            if not packages:
                self._log("Trying recursive search...")
                packages = [(e.get('ID', 'Unknown'), e) for e in root.iter() if e.tag.endswith('Package')]

            for pkg_id, _ in packages:
                self._log(f"  Found package: {pkg_id}")
            self._log(f"Total packages found: {len(packages)}")
            return packages
        except Exception as e:
            self._log(f"Error parsing XML: {e}")
            self._log(traceback.format_exc())
            raise

    # ------------------------------------------------------------------ #
    # Microsoft tool mode                                                  #
    # ------------------------------------------------------------------ #

    def run_microsoft_tool(self):
        """Download and run Microsoft's official Intune ODC PowerShell script."""
        self._update_status("Downloading Microsoft Intune ODC script...")

        ps1_path = os.path.join(self.log_dir, "IntuneODCStandAlone.ps1")
        xml_path = os.path.join(self.log_dir, "Intune.xml")

        try:
            self._log("Downloading from https://aka.ms/intuneps1")
            self._download_file("https://aka.ms/intuneps1", ps1_path)
            self._log(f"Downloaded: {ps1_path}")
            self._set_progress(20)

            self._log("Downloading from https://aka.ms/intunexml")
            self._download_file("https://aka.ms/intunexml", xml_path)
            self._log(f"Downloaded: {xml_path}")
            self._set_progress(30)

            self._update_status("Running Microsoft Intune ODC collection script...")
            self._log("This may take 10-15 minutes...")

            result = subprocess.run(
                ['powershell', '-ExecutionPolicy', 'Bypass', '-File', ps1_path],
                capture_output=True, text=True, timeout=900, cwd=self.log_dir,
            )

            self._log("Microsoft script output:")
            for line in result.stdout.splitlines():
                if line.strip():
                    self._log(f"  {line.strip()}")
            for line in result.stderr.splitlines():
                if line.strip():
                    self._log(f"  ! {line.strip()}")

            self._set_progress(90)

            zip_files = [f for f in os.listdir(self.log_dir) if f.endswith('.zip')]
            if zip_files:
                self._log(f"Created: {zip_files[0]}")

            self._set_progress(100)
            self._update_status("Microsoft tool collection complete!")

        except subprocess.TimeoutExpired:
            self._log("Microsoft script timed out (this is normal, it takes a while)")
            zip_files = [f for f in os.listdir(self.log_dir) if f.endswith('.zip')]
            if zip_files:
                self._log(f"ZIP file was created: {zip_files[0]}")
                self._set_progress(100)
            else:
                raise Exception("Collection timed out and no ZIP file was created")
        except Exception as e:
            self._log(f"Error running Microsoft tool: {e}")
            raise

    # ------------------------------------------------------------------ #
    # Collection orchestration                                             #
    # ------------------------------------------------------------------ #

    def start_collection(self):
        """Validate prerequisites, confirm with the user, then start collection."""
        if not self.is_admin():
            messagebox.showerror(
                "Administrator Required",
                "This tool must be run as Administrator.\n\n"
                "Please right-click and select 'Run as administrator'.",
            )
            return

        if not messagebox.askyesno(
            "Start Collection",
            "This will collect Intune diagnostic logs.\n\n"
            f"The process takes approximately 10 minutes.\n"
            f"Output will be saved to: {self.log_dir}\n\n"
            "Do you want to continue?",
        ):
            return

        self.is_running = True
        self.collect_btn.config(state=tk.DISABLED)
        self.cancel_btn.config(state=tk.NORMAL)
        self.open_btn.config(state=tk.DISABLED)
        self.progress_var.set(0)
        self.output_text.config(state=tk.NORMAL)
        self.output_text.delete(1.0, tk.END)

        threading.Thread(target=self._collection_thread, daemon=True).start()

    def _collection_thread(self):
        """Background thread: run the selected collection mode end-to-end."""
        try:
            self._update_status("Creating directories...")
            os.makedirs(self.log_dir, exist_ok=True)

            mode = self.collection_mode.get()
            self._log(f"Collection mode: {mode}")

            if mode == 'microsoft':
                self.run_microsoft_tool()
            else:
                self._run_native_collection()

            self._ui_call(self.open_btn.config, state=tk.NORMAL)
            self._ui_call(
                messagebox.showinfo,
                "Collection Complete",
                f"Intune ODC logs have been collected!\n\nLocation: {self.log_dir}",
            )
        except Exception as e:
            self._update_status(f"Error: {e}")
            self._ui_call(messagebox.showerror, "Error", str(e))
        finally:
            self.is_running = False
            self._ui_call(self.collect_btn.config, state=tk.NORMAL)
            self._ui_call(self.cancel_btn.config, state=tk.DISABLED)

    def _run_native_collection(self):
        """Orchestrate native Python collection: download XML, parse, collect, zip."""
        self.result_dir = os.path.join(os.environ.get('TEMP', r'C:\Temp'), 'CollectedData')
        if os.path.exists(self.result_dir):
            shutil.rmtree(self.result_dir)
        os.makedirs(self.result_dir, exist_ok=True)
        self._set_progress(10)

        xml_path = self.download_xml()
        self._set_progress(25)

        packages = self.parse_xml(xml_path)
        self._set_progress(30)

        # Detect namespace from first package element
        ns_map = {}
        if packages:
            first_tag = packages[0][1].tag
            if '}' in first_tag:
                ns_map = {'ns': first_tag.split('}')[0].strip('{')}

        total = len(packages)
        for i, (pkg_id, package) in enumerate(packages):
            if not self.is_running:
                return

            self._update_status(f"Processing package: {pkg_id}")

            files_elem = self._find_child(package, 'Files', ns_map)
            if files_elem is not None:
                count = self._collect_files(pkg_id, files_elem, ns_map)
                if count:
                    self._log(f"  Collected {count} files")

            reg_elem = self._find_child(package, 'Registries', ns_map)
            if reg_elem is not None:
                count = self._collect_registry(pkg_id, reg_elem, ns_map)
                if count:
                    self._log(f"  Collected {count} registry keys")

            evt_elem = self._find_child(package, 'EventLogs', ns_map)
            if evt_elem is not None:
                count = self._collect_eventlogs(pkg_id, evt_elem, ns_map)
                if count:
                    self._log(f"  Collected {count} event logs")

            cmd_elem = self._find_child(package, 'Commands', ns_map)
            if cmd_elem is not None:
                count = self._collect_commands(pkg_id, cmd_elem, ns_map)
                if count:
                    self._log(f"  Collected {count} command outputs")

            self._set_progress(int(30 + (i + 1) / total * 60))

        self._create_zip()
        self._set_progress(100)

        if os.path.exists(self.result_dir):
            shutil.rmtree(self.result_dir)

        self._update_status("Collection complete!")

    # ------------------------------------------------------------------ #
    # Collectors                                                           #
    # ------------------------------------------------------------------ #

    def _collect_files(self, package_id, files_element, ns_map):
        """Copy files listed in the XML <Files> element to the result directory."""
        file_elems = self._find_xml_elems(files_element, 'File', ns_map)
        self._log(f"  Processing {len(file_elems)} file entries...")
        collected = 0

        for elem in file_elems:
            if not self.is_running:
                return collected

            raw = elem.text
            if not raw:
                continue

            file_path = os.path.expandvars(raw).replace('"', '')
            team = elem.get('Team', 'General')
            dest_dir = os.path.join(self.result_dir, package_id, "Files", team)

            try:
                # Check for wildcards BEFORE os.path.exists (which returns False for globs)
                if '*' in file_path:
                    matched = glob.glob(file_path)
                elif os.path.isfile(file_path):
                    matched = [file_path]
                else:
                    matched = []

                for src in matched:
                    if not os.path.isfile(src):
                        continue
                    os.makedirs(dest_dir, exist_ok=True)
                    dest = os.path.join(dest_dir, f"{_hostname()}_{os.path.basename(src)}")
                    shutil.copy2(src, dest)
                    self._log(f"  Collected file: {os.path.basename(src)}")
                    collected += 1
            except Exception as e:
                self._log(f"  Error collecting file {file_path}: {e}")

        return collected

    def _collect_registry(self, package_id, reg_element, ns_map):
        """Export registry keys listed in the XML <Registries> element."""
        reg_elems = self._find_xml_elems(reg_element, 'Registry', ns_map)
        self._log(f"  Processing {len(reg_elems)} registry entries...")
        collected = 0

        for elem in reg_elems:
            if not self.is_running:
                return collected

            reg_path = elem.text
            if not reg_path:
                continue

            # Strip trailing wildcard (e.g. "HKLM\Software\*" -> "HKLM\Software")
            reg_path = reg_path.rstrip('\\*').strip()
            team = elem.get('Team', 'General')
            output_file = elem.get('OutputFileName') or reg_path.replace('\\', '_')

            dest_dir = os.path.join(self.result_dir, package_id, "RegistryKeys", team)
            os.makedirs(dest_dir, exist_ok=True)
            # Use .reg extension — reg.exe exports in REG file format
            dest = os.path.join(dest_dir, f"{_hostname()}_{output_file}.reg")

            try:
                result = subprocess.run(
                    ['reg', 'export', reg_path, dest, '/y', '/reg:64'],
                    capture_output=True, text=True,
                )
                if result.returncode == 0:
                    self._log(f"  Collected registry: {reg_path}")
                    collected += 1
                else:
                    self._log(f"  Skip registry (not found or access denied): {reg_path}")
            except Exception as e:
                self._log(f"  Error collecting registry {reg_path}: {e}")

        return collected

    def _collect_eventlogs(self, package_id, evt_element, ns_map):
        """Copy event log files listed in the XML <EventLogs> element."""
        evt_elems = self._find_xml_elems(evt_element, 'EventLog', ns_map)
        self._log(f"  Processing {len(evt_elems)} event log entries...")
        collected = 0

        for elem in evt_elems:
            if not self.is_running:
                return collected

            raw = elem.text
            if not raw:
                continue

            log_path = os.path.expandvars(raw)
            team = elem.get('Team', 'General')
            dest_dir = os.path.join(self.result_dir, package_id, "EventLogs", team)

            try:
                # Check for wildcards BEFORE os.path.exists (which returns False for globs)
                if '*' in log_path:
                    matched = glob.glob(log_path)
                elif os.path.isfile(log_path):
                    matched = [log_path]
                else:
                    matched = []

                for src in matched:
                    if not os.path.isfile(src):
                        continue
                    os.makedirs(dest_dir, exist_ok=True)
                    dest = os.path.join(dest_dir, f"{_hostname()}_{os.path.basename(src)}")
                    shutil.copy2(src, dest)
                    self._log(f"  Collected event log: {os.path.basename(src)}")
                    collected += 1
            except Exception as e:
                self._log(f"  Error collecting event log {log_path}: {e}")

        return collected

    def _collect_commands(self, package_id, cmd_element, ns_map):
        """Execute commands listed in the XML <Commands> element and save output."""
        cmd_elems = self._find_xml_elems(cmd_element, 'Command', ns_map)
        self._log(f"  Processing {len(cmd_elems)} command entries...")
        collected = 0

        # PowerShell helper injected before every PS command block
        RUNCOMMAND_PS = r'''
function RunCommand($cmdToRun) {
    Write-Host "=== Executing: $cmdToRun ==="
    try {
        if ($cmdToRun -match "^\s*[a-zA-Z0-9_-]+\.exe" -or
            $cmdToRun -match "^\s*[a-zA-Z0-9_-]+\.cmd" -or
            $cmdToRun -match "^\s*[a-zA-Z0-9_-]+\.bat") {
            $output = cmd /c $cmdToRun 2>&1
        } else {
            $output = Invoke-Expression $cmdToRun 2>&1
        }
        Write-Output ($output | Out-String)
    } catch {
        Write-Error "Error executing command: $_"
    }
}
'''

        for elem in cmd_elems:
            if not self.is_running:
                return collected

            cmd_type = elem.get('Type', 'PS').upper()
            cmd_text = elem.text
            if not cmd_text:
                continue

            output_file = elem.get('OutputFileName', 'output')
            if output_file == 'NA':
                continue

            team = elem.get('Team', 'General')
            dest_dir = os.path.join(self.result_dir, package_id, "Commands", team)
            os.makedirs(dest_dir, exist_ok=True)
            dest = os.path.join(dest_dir, f"{_hostname()}_{output_file}.txt")

            try:
                if cmd_type == 'PS':
                    fd, tmp = tempfile.mkstemp(suffix='.ps1')
                    try:
                        with os.fdopen(fd, 'w', encoding='utf-8') as f:
                            f.write(RUNCOMMAND_PS)
                            f.write('\n# Execute command(s)\n')
                            f.write(cmd_text)
                            f.write('\n')
                        result = subprocess.run(
                            ['powershell', '-ExecutionPolicy', 'Bypass', '-File', tmp],
                            capture_output=True, text=True, timeout=120,
                        )
                    finally:
                        try:
                            os.remove(tmp)
                        except Exception:
                            pass

                elif cmd_type == 'CMD':
                    fd, tmp = tempfile.mkstemp(suffix='.cmd')
                    try:
                        with os.fdopen(fd, 'w', encoding='utf-8') as f:
                            f.write('@echo off\n')
                            f.write(cmd_text)
                        result = subprocess.run(
                            ['cmd', '/c', tmp],
                            capture_output=True, text=True, timeout=120,
                        )
                    finally:
                        try:
                            os.remove(tmp)
                        except Exception:
                            pass

                else:
                    continue

                with open(dest, 'w', encoding='utf-8') as f:
                    f.write(result.stdout + result.stderr)

                self._log(f"  Collected command output: {cmd_text[:50].strip()}...")
                collected += 1

            except subprocess.TimeoutExpired:
                self._log(f"  Timeout running command: {cmd_text[:50].strip()}")
            except Exception as e:
                self._log(f"  Error running command: {e}")

        return collected

    # ------------------------------------------------------------------ #
    # ZIP                                                                  #
    # ------------------------------------------------------------------ #

    def _create_zip(self):
        """Bundle the collected data directory into a timestamped ZIP file."""
        self._update_status("Creating ZIP file...")
        timestamp = datetime.utcnow().strftime("%m_%d_%Y_%H_%M_UTC")
        zip_name = f"{_hostname()}_CollectedData_{timestamp}.zip"
        zip_path = os.path.join(self.log_dir, zip_name)

        try:
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root, _dirs, files in os.walk(self.result_dir):
                    for name in files:
                        src = os.path.join(root, name)
                        zf.write(src, os.path.relpath(src, self.result_dir))
            self._log(f"Created ZIP: {zip_name}")
            return zip_path
        except Exception as e:
            self._log(f"Error creating ZIP: {e}")
            raise

    # ------------------------------------------------------------------ #
    # UI actions                                                           #
    # ------------------------------------------------------------------ #

    def cancel_collection(self):
        """Ask the user to confirm cancellation, then set the stop flag."""
        if self.is_running and messagebox.askyesno(
            "Cancel Collection", "Are you sure you want to cancel?"
        ):
            self.is_running = False
            self._update_status("Collection cancelled")

    def open_log_folder(self):
        """Open the output folder in Windows Explorer."""
        if os.path.exists(self.log_dir):
            os.startfile(self.log_dir)
        else:
            messagebox.showwarning("Not Found", f"Folder not found: {self.log_dir}")

    def run(self):
        self.root.mainloop()


def main():
    app = ODCLogCollector()
    app.run()


if __name__ == '__main__':
    main()
