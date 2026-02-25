#!/usr/bin/env python3
"""
Intune ODC Log Collector

GUI tool to collect Intune One Data Collector (ODC) logs for troubleshooting.
Runs the official Microsoft diagnostic collection script.
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import subprocess
import threading
import os
import sys
from pathlib import Path
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
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the user interface"""
        # Main container
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        self._create_header(main_frame)
        
        # Info section
        self._create_info_section(main_frame)
        
        # Actions
        self._create_actions_section(main_frame)
        
        # Progress
        self._create_progress_section(main_frame)
        
        # Output
        self._create_output_section(main_frame)
        
    def _create_header(self, parent):
        """Create header"""
        header = ttk.Frame(parent)
        header.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(
            header,
            text="📁 Intune ODC Log Collector",
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

The process will:
1. Create directory: C:\IntuneODCLogs
2. Download Intune.XML configuration from Microsoft
3. Collect Intune configuration and logs using embedded PowerShell script (takes ~10 minutes)
4. Create a compressed ZIP file with all data

⚠️  Note: This tool must be run as Administrator."""
        
        ttk.Label(
            info_frame,
            text=info_text,
            font=('Segoe UI', 9),
            justify=tk.LEFT,
            wraplength=650
        ).pack(anchor=tk.W)
        
    def _create_actions_section(self, parent):
        """Create action buttons"""
        action_frame = ttk.Frame(parent)
        action_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Collect button
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
        
        # Open folder button
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
        
        # Cancel button
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
        
        # Progress bar
        self.progress_var = tk.DoubleVar(value=0)
        self.progress = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate'
        )
        self.progress.pack(fill=tk.X, pady=(0, 5))
        
        # Time estimate label
        self.time_var = tk.StringVar(value="Estimated time: ~10 minutes")
        ttk.Label(
            progress_frame,
            textvariable=self.time_var,
            font=('Segoe UI', 9)
        ).pack(anchor=tk.W)
        
        # Status label
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
        
        # Initial message
        self._log("Ready to collect Intune ODC logs.")
        self._log("Click 'Start Collection' to begin.")
        self._log("")
        self._log("⚠️  This tool must be run as Administrator.")
        
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
        
    def start_collection(self):
        """Start the log collection process"""
        # Check if running as admin
        if not self.is_admin():
            messagebox.showerror(
                "Administrator Required",
                "This tool must be run as Administrator.\n\n"
                "Please right-click and select 'Run as administrator'."
            )
            return
            
        # Confirm start
        result = messagebox.askyesno(
            "Start Collection",
            "This will collect Intune diagnostic logs.\n\n"
            "The process takes approximately 10 minutes.\n"
            "Output will be saved to: C:\\IntuneODCLogs\n\n"
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
        
        # Start collection in thread
        thread = threading.Thread(target=self._collection_thread)
        thread.daemon = True
        thread.start()
        
    def _collection_thread(self):
        """Run collection in background thread"""
        try:
            log_dir = r"C:\IntuneODCLogs"
            
            # Get path to embedded PowerShell script
            if getattr(sys, 'frozen', False):
                # Running in PyInstaller bundle
                base_dir = Path(sys.executable).parent
            else:
                # Running in normal Python
                base_dir = Path(__file__).parent
            
            ps_script_source = base_dir / "IntuneODCStandAlone.ps1"
            ps_script_dest = Path(log_dir) / "IntuneODCStandAlone.ps1"
            
            # Step 1: Create directory
            self._update_status("Creating log directory...")
            self._run_command(f'cmd /c "md {log_dir} 2>nul || echo Directory exists"')
            self.progress_var.set(10)
            
            # Step 2: Download Intune.xml
            self._update_status("Downloading Intune.xml...")
            self._run_command(
                f'powershell -Command "Set-Location {log_dir}; '
                f'Invoke-WebRequest -Uri https://aka.ms/intunexml -OutFile Intune.xml -UseBasicParsing"'
            )
            self.progress_var.set(25)
            
            # Step 3: Copy embedded PowerShell script
            self._update_status("Extracting embedded PowerShell script...")
            if ps_script_source.exists():
                import shutil
                shutil.copy(str(ps_script_source), str(ps_script_dest))
                self._log(f"Copied embedded script to {ps_script_dest}")
            else:
                self._log(f"Warning: Embedded script not found at {ps_script_source}")
                self._log("Falling back to download...")
                self._run_command(
                    f'powershell -Command "Set-Location {log_dir}; '
                    f'Invoke-WebRequest -Uri https://aka.ms/intuneps1 -OutFile IntuneODCStandAlone.ps1 -UseBasicParsing"'
                )
            self.progress_var.set(40)
            
            # Step 4: Run the collection script
            self._update_status("Collecting Intune logs (this takes ~10 minutes)...")
            self._update_status("Please wait, gathering diagnostic information...")
            
            # Run the main collection
            result = self._run_command(
                f'powershell -ExecutionPolicy Bypass -File "{log_dir}\\IntuneODCStandAlone.ps1"',
                timeout=600  # 10 minute timeout
            )
            
            self.progress_var.set(90)
            
            # Step 5: Check for output
            self._update_status("Checking for output files...")
            self._check_output_files(log_dir)
            
            self.progress_var.set(100)
            self._update_status("✅ Collection complete!")
            
            # Enable open folder button
            self.root.after(0, lambda: self.open_btn.config(state=tk.NORMAL))
            
            # Show completion message
            self.root.after(0, lambda: messagebox.showinfo(
                "Collection Complete",
                "Intune ODC logs have been collected successfully!\n\n"
                f"Location: {log_dir}\n\n"
                "You can now open the log folder to find the ZIP file."
            ))
            
        except Exception as e:
            self._update_status(f"❌ Error: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.is_running = False
            self.root.after(0, self._collection_finished)
            
    def _run_command(self, command, timeout=None):
        """Run a shell command and return output"""
        try:
            self._log(f"Running: {command[:80]}...")
            
            result = subprocess.run(
                command,
                shell=True,
                capture_output=True,
                text=True,
                timeout=timeout
            )
            
            if result.stdout:
                for line in result.stdout.strip().split('\n'):
                    if line.strip():
                        self._log(f"  {line.strip()}")
                        
            if result.stderr:
                for line in result.stderr.strip().split('\n'):
                    if line.strip():
                        self._log(f"  ! {line.strip()}")
                        
            return result
            
        except subprocess.TimeoutExpired:
            self._log("  Command timed out (this is normal for long operations)")
            return None
        except Exception as e:
            self._log(f"  Error: {str(e)}")
            raise
            
    def _check_output_files(self, log_dir):
        """Check for and list output files"""
        try:
            if os.path.exists(log_dir):
                files = os.listdir(log_dir)
                self._log(f"Files in {log_dir}:")
                for f in files:
                    file_path = os.path.join(log_dir, f)
                    size = os.path.getsize(file_path)
                    size_mb = size / (1024 * 1024)
                    self._log(f"  - {f} ({size_mb:.2f} MB)")
        except Exception as e:
            self._log(f"Could not list files: {e}")
            
    def _collection_finished(self):
        """Called when collection finishes"""
        self.collect_btn.config(state=tk.NORMAL)
        self.cancel_btn.config(state=tk.DISABLED)
        self.time_var.set("Collection finished")
        
    def cancel_collection(self):
        """Cancel the collection process"""
        if self.is_running:
            result = messagebox.askyesno(
                "Cancel Collection",
                "Are you sure you want to cancel?\n\n"
                "The collection process will be terminated."
            )
            if result:
                self.is_running = False
                self._update_status("⚠️  Collection cancelled by user")
                self._collection_finished()
                
    def open_log_folder(self):
        """Open the log folder in Explorer"""
        log_dir = r"C:\IntuneODCLogs"
        if os.path.exists(log_dir):
            os.startfile(log_dir)
        else:
            messagebox.showwarning(
                "Folder Not Found",
                f"Log folder not found: {log_dir}\n\n"
                "The collection may not have completed successfully."
            )
            
    def is_admin(self):
        """Check if running as administrator"""
        try:
            import ctypes
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False
            
    def run(self):
        """Run the application"""
        self.root.mainloop()


def main():
    """Main entry point"""
    app = ODCLogCollector()
    app.run()


if __name__ == '__main__':
    main()
