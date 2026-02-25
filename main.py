#!/usr/bin/env python3
"""
Intune ODC Log Collector - Native Python Implementation

Collects Intune diagnostic logs directly in Python without PowerShell.
Downloads Intune.XML and processes it natively.
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import os
import sys
import shutil
import zipfile
import subprocess
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime
import time


class ODCLogCollector:
    """Main application class"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Intune ODC Log Collector")
        self.root.geometry("750x600")
        self.root.minsize(750, 500)
        
        # Set DPI awareness on Windows
        if sys.platform == 'win32':
            try:
                from ctypes import windll
                windll.shcore.SetProcessDpiAwareness(1)
            except:
                pass
        
        self.is_running = False
        self.log_dir = r"C:\IntuneODCLogs"
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the user interface"""
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        self._create_header(main_frame)
        self._create_info_section(main_frame)
        self._create_mode_section(main_frame)
        self._create_actions_section(main_frame)
        self._create_progress_section(main_frame)
        self._create_output_section(main_frame)
        
    def _create_header(self, parent):
        """Create header"""
        header = ttk.Frame(parent)
        header.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(
            header,
            text="Intune ODC Log Collector",
            font=('Segoe UI', 16, 'bold')
        ).pack(anchor=tk.W)
        
        ttk.Label(
            header,
            text="Collect diagnostic logs for Microsoft Intune troubleshooting",
            font=('Segoe UI', 9)
        ).pack(anchor=tk.W)
        
        ttk.Separator(parent, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=5)
        
    def _create_info_section(self, parent):
        """Create info section"""
        info_frame = ttk.LabelFrame(parent, text="What This Tool Does", padding="10")
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        info_text = """This tool collects Intune One Data Collector (ODC) logs which are used by Microsoft Support to diagnose Intune issues.

Choose collection method:
• Native Python - Uses built-in Python collection (recommended, faster)
• Microsoft Tool - Downloads and runs official Microsoft PowerShell script

The process takes ~10 minutes and creates a compressed ZIP file with all data.

Note: This tool must be run as Administrator."""
        
        ttk.Label(
            info_frame,
            text=info_text,
            font=('Segoe UI', 9),
            justify=tk.LEFT,
            wraplength=650
        ).pack(anchor=tk.W)
        
    def _create_mode_section(self, parent):
        """Create collection mode selection"""
        mode_frame = ttk.LabelFrame(parent, text="Collection Mode", padding="10")
        mode_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.collection_mode = tk.StringVar(value='native')
        
        # Native Python option
        native_frame = ttk.Frame(mode_frame)
        native_frame.pack(fill=tk.X, pady=2)
        
        ttk.Radiobutton(
            native_frame,
            text="Native Python",
            variable=self.collection_mode,
            value='native'
        ).pack(side=tk.LEFT)
        
        ttk.Label(
            native_frame,
            text="(Recommended - Faster, works offline with cached XML)",
            font=('Segoe UI', 8),
            foreground='green'
        ).pack(side=tk.LEFT, padx=(5, 0))
        
        # Microsoft Tool option
        ms_frame = ttk.Frame(mode_frame)
        ms_frame.pack(fill=tk.X, pady=2)
        
        ttk.Radiobutton(
            ms_frame,
            text="Microsoft Tool",
            variable=self.collection_mode,
            value='microsoft'
        ).pack(side=tk.LEFT)
        
        ttk.Label(
            ms_frame,
            text="(Uses official Microsoft PowerShell script)",
            font=('Segoe UI', 8),
            foreground='gray'
        ).pack(side=tk.LEFT, padx=(5, 0))
        
        # Cache XML checkbox
        self.cache_xml_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            mode_frame,
            text="Cache Intune.XML locally (faster subsequent runs)",
            variable=self.cache_xml_var
        ).pack(anchor=tk.W, pady=(5, 0))
        
    def _create_actions_section(self, parent):
        """Create action buttons"""
        action_frame = ttk.Frame(parent)
        action_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.collect_btn = tk.Button(
            action_frame,
            text="Start",
            command=self.start_collection,
            bg='#007bff',
            fg='white',
            disabledforeground='white',
            font=('Segoe UI', 11, 'bold'),
            width=10,
            cursor='hand2'
        )
        self.collect_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.open_btn = tk.Button(
            action_frame,
            text="Open Folder",
            command=self.open_log_folder,
            bg='#6c757d',
            fg='white',
            disabledforeground='white',
            font=('Segoe UI', 10),
            width=12,
            cursor='hand2',
            state=tk.DISABLED
        )
        self.open_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.cancel_btn = tk.Button(
            action_frame,
            text="Cancel",
            command=self.cancel_collection,
            bg='#dc3545',
            fg='white',
            disabledforeground='white',
            font=('Segoe UI', 10),
            width=10,
            cursor='hand2',
            state=tk.DISABLED
        )
        self.cancel_btn.pack(side=tk.LEFT)
        
    def _create_progress_section(self, parent):
        """Create progress section"""
        progress_frame = ttk.LabelFrame(parent, text="Progress", padding="10")
        progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.progress_var = tk.DoubleVar(value=0)
        self.progress = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate'
        )
        self.progress.pack(fill=tk.X, pady=(0, 5))
        
        self.time_var = tk.StringVar(value="Estimated time: ~10 minutes")
        ttk.Label(progress_frame, textvariable=self.time_var, font=('Segoe UI', 9)).pack(anchor=tk.W)
        
        self.status_var = tk.StringVar(value="Ready to start")
        self.status_label = ttk.Label(
            progress_frame,
            textvariable=self.status_var,
            font=('Segoe UI', 9, 'bold')
        )
        self.status_label.pack(anchor=tk.W, pady=(5, 0))
        
    def _create_output_section(self, parent):
        """Create output log section"""
        output_frame = ttk.LabelFrame(parent, text="Output", padding="10")
        output_frame.pack(fill=tk.BOTH, expand=True)
        
        self.output_text = scrolledtext.ScrolledText(
            output_frame,
            wrap=tk.WORD,
            font=('Consolas', 9),
            height=12,
            bg='#f5f5f5'
        )
        self.output_text.pack(fill=tk.BOTH, expand=True)
        
        self._log("Ready to collect Intune ODC logs.")
        self._log("Click 'Start' to begin.")
        self._log("")
        self._log("Note: This tool must be run as Administrator.")
        
    def _log(self, message):
        """Add message to output"""
        self.output_text.config(state=tk.NORMAL)
        timestamp = time.strftime("%H:%M:%S")
        self.output_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.output_text.see(tk.END)
        self.output_text.config(state=tk.DISABLED)
        
    def _update_status(self, message):
        """Update status label"""
        self.status_var.set(message)
        self._log(message)
        
    def is_admin(self):
        """Check if running as administrator"""
        try:
            import ctypes
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    def download_xml(self):
        """Download Intune.XML from Microsoft with caching"""
        xml_path = os.path.join(self.log_dir, "Intune.xml")
        cache_path = os.path.join(self.log_dir, "Intune.xml.cached")
        url = "https://raw.githubusercontent.com/markstan/IntuneOneDataCollector/master/Intune.xml"
        
        # Check if we should use cached version
        if self.cache_xml_var.get() and os.path.exists(cache_path):
            # Check if cache is less than 7 days old
            cache_age = time.time() - os.path.getmtime(cache_path)
            if cache_age < 7 * 24 * 3600:  # 7 days
                self._log(f"Using cached Intune.XML (age: {cache_age/3600:.1f} hours)")
                shutil.copy(cache_path, xml_path)
                return xml_path
            else:
                self._log("Cache expired, downloading fresh copy...")
        
        self._update_status("Downloading Intune.XML...")
        
        try:
            urllib.request.urlretrieve(url, xml_path)
            self._log(f"Downloaded: {xml_path}")
            
            # Cache the file if caching is enabled
            if self.cache_xml_var.get():
                shutil.copy(xml_path, cache_path)
                self._log("Cached Intune.XML for future use")
            
            return xml_path
        except Exception as e:
            self._log(f"Error downloading XML: {e}")
            # Try to use cached version as fallback even if expired
            if os.path.exists(cache_path):
                self._log("Using expired cached XML as fallback...")
                shutil.copy(cache_path, xml_path)
                return xml_path
            raise

    def parse_xml(self, xml_path):
        """Parse the Intune XML file"""
        self._update_status("Parsing Intune.XML...")
        
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            # Debug: log the root element
            self._log(f"XML Root: {root.tag}")
            
            # Extract namespace if present
            ns = {}
            if '}' in root.tag:
                ns_uri = root.tag.split('}')[0].strip('{')
                ns = {'ns': ns_uri}
                self._log(f"Detected namespace: {ns_uri}")
            
            # Find all packages
            packages = []
            
            # Try with namespace
            if ns:
                for package in root.findall('.//ns:Package', ns):
                    pkg_id = package.get('ID', 'Unknown')
                    packages.append((pkg_id, package))
                    self._log(f"  Found package: {pkg_id}")
            
            # Try without namespace
            if not packages:
                for package in root.findall('.//Package'):
                    pkg_id = package.get('ID', 'Unknown')
                    packages.append((pkg_id, package))
                    self._log(f"  Found package: {pkg_id}")
            
            # Search all children recursively
            if not packages:
                self._log("Trying recursive search...")
                for elem in root.iter():
                    if 'Package' in elem.tag and elem.tag.endswith('Package'):
                        pkg_id = elem.get('ID', 'Unknown')
                        packages.append((pkg_id, elem))
                        self._log(f"  Found package: {pkg_id}")
                
            self._log(f"Total packages found: {len(packages)}")
            return packages
        except Exception as e:
            self._log(f"Error parsing XML: {e}")
            import traceback
            self._log(traceback.format_exc())
            raise

    def run_microsoft_tool(self):
        """Download and run Microsoft's official Intune ODC PowerShell script"""
        self._update_status("Downloading Microsoft Intune ODC script...")
        
        ps1_url = "https://aka.ms/intuneps1"
        xml_url = "https://aka.ms/intunexml"
        
        ps1_path = os.path.join(self.log_dir, "IntuneODCStandAlone.ps1")
        xml_path = os.path.join(self.log_dir, "Intune.xml")
        
        try:
            # Download the PowerShell script
            self._log(f"Downloading from {ps1_url}")
            urllib.request.urlretrieve(ps1_url, ps1_path)
            self._log(f"Downloaded: {ps1_path}")
            self.progress_var.set(20)
            
            # Download the XML
            self._log(f"Downloading from {xml_url}")
            urllib.request.urlretrieve(xml_url, xml_path)
            self._log(f"Downloaded: {xml_path}")
            self.progress_var.set(30)
            
            # Run the PowerShell script
            self._update_status("Running Microsoft Intune ODC collection script...")
            self._log("This may take 10-15 minutes...")
            
            result = subprocess.run(
                ['powershell', '-ExecutionPolicy', 'Bypass', '-File', ps1_path],
                capture_output=True,
                text=True,
                timeout=900,  # 15 minute timeout
                cwd=self.log_dir
            )
            
            self._log("Microsoft script output:")
            if result.stdout:
                for line in result.stdout.split('\n'):
                    if line.strip():
                        self._log(f"  {line.strip()}")
            
            if result.stderr:
                for line in result.stderr.split('\n'):
                    if line.strip():
                        self._log(f"  ! {line.strip()}")
            
            self.progress_var.set(90)
            
            # Find the created ZIP file
            zip_files = [f for f in os.listdir(self.log_dir) if f.endswith('.zip')]
            if zip_files:
                self._log(f"Created: {zip_files[0]}")
            
            self.progress_var.set(100)
            self._update_status("Microsoft tool collection complete!")
            
        except subprocess.TimeoutExpired:
            self._log("Microsoft script timed out (this is normal, it takes a while)")
            # Check if ZIP was created anyway
            zip_files = [f for f in os.listdir(self.log_dir) if f.endswith('.zip')]
            if zip_files:
                self._log(f"ZIP file was created: {zip_files[0]}")
                self.progress_var.set(100)
            else:
                raise Exception("Collection timed out and no ZIP file was created")
        except Exception as e:
            self._log(f"Error running Microsoft tool: {e}")
            raise

    def start_collection(self):
        """Start the log collection process"""
        if not self.is_admin():
            messagebox.showerror(
                "Administrator Required",
                "This tool must be run as Administrator.\n\n"
                "Please right-click and select 'Run as administrator'."
            )
            return
            
        result = messagebox.askyesno(
            "Start Collection",
            "This will collect Intune diagnostic logs.\n\n"
            "The process takes approximately 10 minutes.\n"
            f"Output will be saved to: {self.log_dir}\n\n"
            "Do you want to continue?"
        )
        if not result:
            return
            
        # Update UI
        self.is_running = True
        self.collect_btn.config(state=tk.DISABLED)
        self.cancel_btn.config(state=tk.NORMAL)
        self.open_btn.config(state=tk.DISABLED)
        self.progress_var.set(0)
        self.output_text.config(state=tk.NORMAL)
        self.output_text.delete(1.0, tk.END)
        
        # Start collection in thread
        thread = threading.Thread(target=self._collection_thread)
        thread.daemon = True
        thread.start()

    def _collection_thread(self):
        """Run collection in background thread"""
        try:
            # Create directories
            self._update_status("Creating directories...")
            os.makedirs(self.log_dir, exist_ok=True)
            
            # Check collection mode
            mode = self.collection_mode.get()
            self._log(f"Collection mode: {mode}")
            
            if mode == 'microsoft':
                # Use Microsoft's official tool
                self.run_microsoft_tool()
            else:
                # Use native Python implementation
                self._run_native_collection()
                
            # Enable open button for both modes
            self.open_btn.config(state=tk.NORMAL)
            
            messagebox.showinfo(
                "Collection Complete",
                f"Intune ODC logs have been collected!\n\n"
                f"Location: {self.log_dir}"
            )
            
        except Exception as e:
            self._update_status(f"Error: {str(e)}")
            messagebox.showerror("Error", str(e))
        finally:
            self.is_running = False
            self.collect_btn.config(state=tk.NORMAL)
            self.cancel_btn.config(state=tk.DISABLED)

    def _run_native_collection(self):
        """Run native Python collection"""
        self.result_dir = os.path.join(os.environ.get('TEMP', 'C:\\Temp'), 'CollectedData')
        if os.path.exists(self.result_dir):
            shutil.rmtree(self.result_dir)
        os.makedirs(self.result_dir, exist_ok=True)
        self.progress_var.set(10)
        
        # Download XML
        xml_path = self.download_xml()
        self.progress_var.set(25)
        
        # Parse XML
        packages = self.parse_xml(xml_path)
        self.progress_var.set(30)
        
        # Determine namespace from first package if available
        ns_uri = None
        if packages:
            first_pkg = packages[0][1]
            if '}' in first_pkg.tag:
                ns_uri = first_pkg.tag.split('}')[0].strip('{')
        ns_map = {'ns': ns_uri} if ns_uri else {}
        
        # Process each package
        total_packages = len(packages)
        for i, (pkg_id, package) in enumerate(packages):
            if not self.is_running:
                return
                
            self._update_status(f"Processing package: {pkg_id}")
            
            # Helper function to find child element with namespace handling
            def find_child(parent, tag_name):
                # Try with namespace
                if ns_map:
                    child = parent.find(f'ns:{tag_name}', ns_map)
                    if child is not None:
                        return child
                # Try without namespace
                child = parent.find(tag_name)
                if child is not None:
                    return child
                # Try by checking tag ending
                for child in parent:
                    if child.tag.endswith(tag_name):
                        return child
                return None
            
            # Collect files
            files_elem = find_child(package, 'Files')
            if files_elem is not None:
                count = self._collect_files(pkg_id, files_elem, ns_map)
                if count:
                    self._log(f"  Collected {count} files")
                
            # Collect registry
            reg_elem = find_child(package, 'Registries')
            if reg_elem is not None:
                count = self._collect_registry(pkg_id, reg_elem, ns_map)
                if count:
                    self._log(f"  Collected {count} registry keys")
                
            # Collect event logs
            evt_elem = find_child(package, 'EventLogs')
            if evt_elem is not None:
                count = self._collect_eventlogs(pkg_id, evt_elem, ns_map)
                if count:
                    self._log(f"  Collected {count} event logs")
                
            # Collect commands
            cmd_elem = find_child(package, 'Commands')
            if cmd_elem is not None:
                count = self._collect_commands(pkg_id, cmd_elem, ns_map)
                if count:
                    self._log(f"  Collected {count} command outputs")
                
            progress = 30 + (i + 1) / total_packages * 60
            self.progress_var.set(int(progress))
        
        # Create ZIP
        zip_path = self._create_zip()
        self.progress_var.set(100)
        
        # Cleanup
        if os.path.exists(self.result_dir):
            shutil.rmtree(self.result_dir)
        
        self._update_status("Collection complete!")

    def _collect_files(self, package_id, files_element, ns_map=None):
        """Collect files specified in XML"""
        if files_element is None:
            return 0
            
        files_collected = 0
        
        # Find all File elements with namespace handling
        if ns_map:
            file_elems = list(files_element.iter(f'{{{ns_map["ns"]}}}File'))
            if not file_elems:
                file_elems = files_element.findall(f'ns:File', ns_map)
        else:
            file_elems = list(files_element.iter('File'))
            if not file_elems:
                file_elems = files_element.findall('File')
        
        self._log(f"  Processing {len(file_elems)} file entries...")
        
        for file_elem in file_elems:
            if not self.is_running:
                return files_collected
                
            file_path = file_elem.text
            if not file_path:
                continue
                
            # Expand environment variables
            file_path = os.path.expandvars(file_path)
            file_path = file_path.replace('"', '')
            
            team = file_elem.get('Team', 'General')
            
            if os.path.exists(file_path):
                try:
                    # Handle wildcards
                    if '*' in file_path:
                        import glob
                        matched_files = glob.glob(file_path)
                    else:
                        matched_files = [file_path]
                        
                    for matched_file in matched_files:
                        if os.path.isfile(matched_file):
                            # Create destination directory
                            dest_dir = os.path.join(self.result_dir, package_id, "Files", team)
                            os.makedirs(dest_dir, exist_ok=True)
                            
                            # Copy file
                            file_name = os.path.basename(matched_file)
                            dest_name = f"{os.environ['COMPUTERNAME']}_{file_name}"
                            dest_path = os.path.join(dest_dir, dest_name)
                            
                            shutil.copy2(matched_file, dest_path)
                            files_collected += 1
                            self._log(f"  Collected file: {file_name}")
                            
                except Exception as e:
                    self._log(f"  Error collecting file {file_path}: {e}")
                    
        return files_collected

    def _collect_registry(self, package_id, reg_element, ns_map=None):
        """Collect registry keys specified in XML"""
        if reg_element is None:
            return 0
            
        reg_collected = 0
        
        # Find all Registry elements with namespace handling
        if ns_map:
            reg_elems = list(reg_element.iter(f'{{{ns_map["ns"]}}}Registry'))
            if not reg_elems:
                reg_elems = reg_element.findall(f'ns:Registry', ns_map)
        else:
            reg_elems = list(reg_element.iter('Registry'))
            if not reg_elems:
                reg_elems = reg_element.findall('Registry')
        
        self._log(f"  Processing {len(reg_elems)} registry entries...")
        
        for reg_elem in reg_elems:
            if not self.is_running:
                return reg_collected
                
            reg_path = reg_elem.text
            if not reg_path:
                continue
                
            # Remove trailing wildcard
            reg_path = reg_path.replace('\*', '')
            
            team = reg_elem.get('Team', 'General')
            output_file = reg_elem.get('OutputFileName', reg_path.replace('\\', '_'))
            
            try:
                # Create destination directory
                dest_dir = os.path.join(self.result_dir, package_id, "RegistryKeys", team)
                os.makedirs(dest_dir, exist_ok=True)
                
                # Export registry using reg.exe
                output_name = f"{os.environ['COMPUTERNAME']}_{output_file}.txt"
                dest_path = os.path.join(dest_dir, output_name)
                
                result = subprocess.run(
                    ['reg', 'export', reg_path, dest_path, '/y', '/reg:64'],
                    capture_output=True,
                    text=True
                )
                
                if result.returncode == 0:
                    reg_collected += 1
                    self._log(f"  Collected registry: {reg_path}")
                else:
                    self._log(f"  Error exporting registry: {reg_path}")
                    
            except Exception as e:
                self._log(f"  Error collecting registry {reg_path}: {e}")
                
        return reg_collected

    def _collect_eventlogs(self, package_id, evt_element, ns_map=None):
        """Collect event logs specified in XML"""
        if evt_element is None:
            return 0
            
        evt_collected = 0
        
        # Find all EventLog elements with namespace handling
        if ns_map:
            evt_elems = list(evt_element.iter(f'{{{ns_map["ns"]}}}EventLog'))
            if not evt_elems:
                evt_elems = evt_element.findall(f'ns:EventLog', ns_map)
        else:
            evt_elems = list(evt_element.iter('EventLog'))
            if not evt_elems:
                evt_elems = evt_element.findall('EventLog')
        
        self._log(f"  Processing {len(evt_elems)} event log entries...")
        
        for evt_elem in evt_elems:
            if not self.is_running:
                return evt_collected
                
            log_path = evt_elem.text
            if not log_path:
                continue
                
            log_path = os.path.expandvars(log_path)
            team = evt_elem.get('Team', 'General')
            
            if os.path.exists(log_path):
                try:
                    # Handle wildcards
                    if '*' in log_path:
                        import glob
                        matched_logs = glob.glob(log_path)
                    else:
                        matched_logs = [log_path]
                        
                    for matched_log in matched_logs:
                        if os.path.isfile(matched_log):
                            # Create destination directory
                            dest_dir = os.path.join(self.result_dir, package_id, "EventLogs", team)
                            os.makedirs(dest_dir, exist_ok=True)
                            
                            # Copy log
                            log_name = os.path.basename(matched_log)
                            dest_name = f"{os.environ['COMPUTERNAME']}_{log_name}"
                            dest_path = os.path.join(dest_dir, dest_name)
                            
                            shutil.copy2(matched_log, dest_path)
                            evt_collected += 1
                            self._log(f"  Collected event log: {log_name}")
                            
                except Exception as e:
                    self._log(f"  Error collecting event log {log_path}: {e}")
                    
        return evt_collected

    def _collect_commands(self, package_id, cmd_element, ns_map=None):
        """Run and collect command outputs specified in XML"""
        if cmd_element is None:
            return 0
            
        cmd_collected = 0
        
        # Find all Command elements with namespace handling
        if ns_map:
            cmd_elems = list(cmd_element.iter(f'{{{ns_map["ns"]}}}Command'))
            if not cmd_elems:
                cmd_elems = cmd_element.findall(f'ns:Command', ns_map)
        else:
            cmd_elems = list(cmd_element.iter('Command'))
            if not cmd_elems:
                cmd_elems = cmd_element.findall('Command')
        
        self._log(f"  Processing {len(cmd_elems)} command entries...")
        
        for cmd_elem in cmd_elems:
            if not self.is_running:
                return cmd_collected
                
            cmd_type = cmd_elem.get('Type', 'PS')
            cmd_text = cmd_elem.text
            
            if not cmd_text:
                continue
                
            team = cmd_elem.get('Team', 'General')
            output_file = cmd_elem.get('OutputFileName', 'output')
            
            if output_file == "NA":
                continue
                
            try:
                # Create destination directory
                dest_dir = os.path.join(self.result_dir, package_id, "Commands", team)
                os.makedirs(dest_dir, exist_ok=True)
                
                # Run command and capture output
                output_name = f"{os.environ['COMPUTERNAME']}_{output_file}.txt"
                dest_path = os.path.join(dest_dir, output_name)
                
                if cmd_type.upper() == "PS":
                    # Enhanced RunCommand function with better output handling
                    RUNCOMMAND_FUNCTION = '''
function RunCommand($cmdToRun) {
    Write-Host "=== Executing: $cmdToRun ==="
    try {
        if ($cmdToRun -match "^\\s*[a-zA-Z0-9_-]+\\.exe" -or $cmdToRun -match "^\\s*[a-zA-Z0-9_-]+\\.cmd" -or $cmdToRun -match "^\\s*[a-zA-Z0-9_-]+\\.bat") {
            $output = cmd /c $cmdToRun 2>&1
        } else {
            $output = Invoke-Expression $cmdToRun 2>&1
        }
        $outputString = $output | Out-String
        Write-Output $outputString
    }
    catch {
        Write-Error "Error executing command: $_"
    }
}

'''
                    # Always use temp script for better reliability
                    temp_script = os.path.join(os.environ.get('TEMP', 'C:\\Temp'), f"odc_cmd_{cmd_collected}.ps1")
                    with open(temp_script, 'w', encoding='utf-8') as f:
                        f.write(RUNCOMMAND_FUNCTION)
                        f.write('\n# Execute command(s)\n')
                        f.write(cmd_text)
                        f.write('\n')
                    
                    # Execute the script file
                    result = subprocess.run(
                        ['powershell', '-ExecutionPolicy', 'Bypass', '-File', temp_script],
                        capture_output=True,
                        text=True,
                        timeout=120
                    )
                    
                    # Clean up temp file
                    try:
                        os.remove(temp_script)
                    except:
                        pass
                    
                    output = result.stdout + result.stderr
                    
                elif cmd_type.upper() == "CMD":
                    # For CMD type, execute via cmd.exe
                    temp_script = os.path.join(os.environ.get('TEMP', 'C:\\Temp'), f"odc_cmd_{cmd_collected}.cmd")
                    with open(temp_script, 'w', encoding='utf-8') as f:
                        f.write('@echo off\n')
                        f.write(cmd_text)
                    
                    # Execute the batch file
                    result = subprocess.run(
                        ['cmd', '/c', temp_script],
                        capture_output=True,
                        text=True,
                        timeout=120
                    )
                    
                    # Clean up temp file
                    try:
                        os.remove(temp_script)
                    except:
                        pass
                    
                    output = result.stdout + result.stderr
                
                else:
                    continue
                
                # Write output to file
                with open(dest_path, 'w') as f:
                    f.write(output)
                    
                cmd_collected += 1
                self._log(f"  Collected command output: {cmd_text[:50]}...")
                
            except Exception as e:
                self._log(f"  Error running command: {e}")
                
        return cmd_collected

    def _create_zip(self):
        """Create ZIP file from collected data"""
        self._update_status("Creating ZIP file...")
        
        timestamp = datetime.utcnow().strftime("%m_%d_%Y_%H_%M_UTC")
        zip_name = f"{os.environ['COMPUTERNAME']}_CollectedData_{timestamp}.zip"
        zip_path = os.path.join(self.log_dir, zip_name)
        
        try:
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(self.result_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, self.result_dir)
                        zipf.write(file_path, arcname)
                        
            self._log(f"Created ZIP: {zip_name}")
            return zip_path
        except Exception as e:
            self._log(f"Error creating ZIP: {e}")
            raise

    def cancel_collection(self):
        """Cancel the collection process"""
        if self.is_running:
            result = messagebox.askyesno(
                "Cancel Collection",
                "Are you sure you want to cancel?"
            )
            if result:
                self.is_running = False
                self._update_status("Collection cancelled")
                
    def open_log_folder(self):
        """Open the log folder in Explorer"""
        if os.path.exists(self.log_dir):
            os.startfile(self.log_dir)
        else:
            messagebox.showwarning("Not Found", "Folder not found: " + self.log_dir)
            
    def run(self):
        """Run the application"""
        self.root.mainloop()


def main():
    app = ODCLogCollector()
    app.run()


if __name__ == '__main__':
    main()
