"""
GUI module for SQL Data Importer application.
"""
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import threading
import os
from typing import Optional
import logging

from config import DatabaseConfig, AppConfig, load_config
from import_service import ImportService

logger = logging.getLogger(__name__)


class ToolTip:
    """Creates tooltips for widgets."""
    
    def __init__(self, widget, text: str = 'widget info', wait_time: int = 500):
        self.wait_time = wait_time
        self.wrap_length = 300
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)
        self.id = None
        self.tw = None
    
    def enter(self, event=None):
        self.schedule()
    
    def leave(self, event=None):
        self.unschedule()
        self.hidetip()
    
    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(self.wait_time, self.showtip)
    
    def unschedule(self):
        id_val = self.id
        self.id = None
        if id_val:
            self.widget.after_cancel(id_val)
    
    def showtip(self, event=None):
        x = y = 0
        try:
            x, y, cx, cy = self.widget.bbox("insert")
        except:
            # Fallback if bbox fails
            x = y = 0
        
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        
        self.tw = tk.Toplevel(self.widget)
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry(f"+{x}+{y}")
        
        label = tk.Label(
            self.tw,
            text=self.text,
            justify='left',
            background="#ffffe0",
            relief='solid',
            borderwidth=1,
            font=("Tahoma", 8, "normal")
        )
        label.pack(ipadx=1)
    
    def hidetip(self):
        tw = self.tw
        self.tw = None
        if tw:
            tw.destroy()


class DBImporterApp:
    """Main application GUI."""
    
    def __init__(self, root: tk.Tk, config_file: Optional[str] = None):
        self.root = root
        self.root.title("SQL Data Importer")
        self.root.geometry("750x600")
        
        # Try to load configuration
        try:
            self.db_config, self.app_config = load_config(config_file)
            self.config_loaded = True
        except Exception as e:
            logger.error(f"Failed to load configuration: {e}")
            self.config_loaded = False
            self.db_config = None
            self.app_config = None
            messagebox.showerror(
                "Configuration Error",
                f"Failed to load configuration:\n{e}\n\n"
                "Please set up database credentials in environment variables or config file."
            )
            # Continue with limited functionality
        
        # File paths
        self.full_input_path = ""
        self.full_config_path = ""
        
        # Display variables
        self.input_filename_var = tk.StringVar()
        self.config_filename_var = tk.StringVar()
        self.progress_var = tk.StringVar(value="")
        
        # Import service
        self.import_service = None
        if self.config_loaded:
            self.import_service = ImportService(
                self.db_config,
                self.app_config,
                self.log
            )
        
        # Setup UI
        self._setup_ui()
        
        # Start connection check
        if self.config_loaded:
            self.root.after(100, self.auto_check_connection)
    
    def _setup_ui(self):
        """Setup the user interface."""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 1. Connection Status Bar
        self._create_status_bar(main_frame)
        
        # 2. File Selection Frame
        self._create_file_selection_frame(main_frame)
        
        # 3. Progress Frame
        self._create_progress_frame(main_frame)
        
        # 4. Action Buttons
        self._create_action_buttons(main_frame)
        
        # 5. Log Window
        self._create_log_window(main_frame)
    
    def _create_status_bar(self, parent):
        """Create connection status bar."""
        status_frame = ttk.Frame(parent)
        status_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(
            status_frame,
            text="DB Connection:",
            font=('Segoe UI', 9)
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        self.lbl_status = ttk.Label(
            status_frame,
            text="⏳ Checking..." if self.config_loaded else "❌ No Config",
            font=('Segoe UI', 9, 'bold'),
            foreground="blue" if self.config_loaded else "red"
        )
        self.lbl_status.pack(side=tk.LEFT)
        
        if self.config_loaded:
            # Show server info
            server_info = f"({self.db_config.server}/{self.db_config.database})"
            ttk.Label(
                status_frame,
                text=server_info,
                font=('Segoe UI', 8),
                foreground="gray"
            ).pack(side=tk.LEFT, padx=(5, 0))
    
    def _create_file_selection_frame(self, parent):
        """Create file selection frame."""
        file_frame = ttk.LabelFrame(parent, text=" Configuration ", padding="10")
        file_frame.pack(fill=tk.X, pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        # Input file row
        ttk.Label(file_frame, text="Input Data:").grid(
            row=0, column=0, sticky="w", pady=2
        )
        
        self.ent_input = ttk.Entry(
            file_frame,
            textvariable=self.input_filename_var,
            state="readonly"
        )
        self.ent_input.grid(row=0, column=1, padx=5, sticky="ew", pady=2)
        self.tooltip_input = ToolTip(
            self.ent_input,
            "No file selected",
            self.app_config.tooltip_wait_time if self.app_config else 500
        )
        
        ttk.Button(
            file_frame,
            text="Browse",
            width=10,
            command=self.browse_input
        ).grid(row=0, column=2, padx=5)
        
        # Config file row
        ttk.Label(file_frame, text="Mapping Config:").grid(
            row=1, column=0, sticky="w", pady=2
        )
        
        self.ent_config = ttk.Entry(
            file_frame,
            textvariable=self.config_filename_var,
            state="readonly"
        )
        self.ent_config.grid(row=1, column=1, padx=5, sticky="ew", pady=2)
        self.tooltip_config = ToolTip(
            self.ent_config,
            "No file selected",
            self.app_config.tooltip_wait_time if self.app_config else 500
        )
        
        ttk.Button(
            file_frame,
            text="Browse",
            width=10,
            command=self.browse_config
        ).grid(row=1, column=2, padx=5)
    
    def _create_progress_frame(self, parent):
        """Create progress indicator frame."""
        progress_frame = ttk.Frame(parent)
        progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.progress_label = ttk.Label(
            progress_frame,
            textvariable=self.progress_var,
            font=('Segoe UI', 9, 'italic'),
            foreground="blue"
        )
        self.progress_label.pack(side=tk.LEFT)
        
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            mode='indeterminate',
            length=200
        )
        # Initially hidden
    
    def _create_action_buttons(self, parent):
        """Create action buttons."""
        action_frame = ttk.Frame(parent)
        action_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.btn_dry_run = ttk.Button(
            action_frame,
            text="🔍 RUN SIMULATION (Dry Run)",
            command=lambda: self.start_process(dry_run=True)
        )
        self.btn_dry_run.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        
        self.btn_import = ttk.Button(
            action_frame,
            text="▶️ RUN REAL IMPORT (Update DB)",
            command=self.confirm_import
        )
        self.btn_import.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)
        
        # Disable if config not loaded
        if not self.config_loaded:
            self.btn_dry_run.config(state="disabled")
            self.btn_import.config(state="disabled")
    
    def _create_log_window(self, parent):
        """Create log window."""
        log_frame = ttk.LabelFrame(parent, text=" Execution Log ", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            state='disabled',
            height=12,
            font=("Consolas", 9),
            wrap=tk.WORD
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Configure tags for colored output
        self.log_text.tag_config("error", foreground="red")
        self.log_text.tag_config("success", foreground="#008000")
        self.log_text.tag_config("header", foreground="#00008B", font=("Consolas", 9, "bold"))
    
    def browse_input(self):
        """Browse for input Excel file."""
        filename = filedialog.askopenfilename(
            title="Select Input Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        
        if filename:
            self.full_input_path = filename
            self.input_filename_var.set(os.path.basename(filename))
            self.tooltip_input.text = filename
            
            # Auto-detect config file
            self._auto_detect_config(filename)
    
    def browse_config(self):
        """Browse for config Excel file."""
        filename = filedialog.askopenfilename(
            title="Select Configuration Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        
        if filename:
            self.full_config_path = filename
            self.config_filename_var.set(os.path.basename(filename))
            self.tooltip_config.text = filename
    
    def _auto_detect_config(self, input_path: str):
        """Auto-detect config file in same directory."""
        if self.full_config_path:
            return  # Already set
        
        folder = os.path.dirname(input_path)
        potential_config = os.path.join(folder, "Script_Config.xlsx")
        
        if os.path.exists(potential_config):
            self.full_config_path = potential_config
            self.config_filename_var.set("Script_Config.xlsx")
            self.tooltip_config.text = potential_config
            self.log("✓ Auto-detected configuration file", "success")
    
    def log(self, message: str, tag: Optional[str] = None):
        """Add message to log window."""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n", tag)
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
    
    def toggle_buttons(self, state: str):
        """Enable or disable action buttons."""
        self.btn_dry_run.config(state=state)
        self.btn_import.config(state=state)
    
    def show_progress(self, message: str):
        """Show progress indicator."""
        self.progress_var.set(message)
        if not self.progress_bar.winfo_viewable():
            self.progress_bar.pack(side=tk.LEFT, padx=(10, 0))
            self.progress_bar.start(10)
    
    def hide_progress(self):
        """Hide progress indicator."""
        self.progress_var.set("")
        self.progress_bar.stop()
        self.progress_bar.pack_forget()
    
    def auto_check_connection(self):
        """Check database connection in background."""
        def perform_check():
            if not self.import_service:
                return
            
            success, error_msg = self.import_service.test_connection()
            
            if success:
                self.root.after(
                    0,
                    lambda: self.lbl_status.config(text="✅ Online", foreground="green")
                )
            else:
                self.root.after(
                    0,
                    lambda: self.lbl_status.config(text="❌ Offline", foreground="red")
                )
                if error_msg:
                    self.root.after(0, lambda: self.log(f"DB Error: {error_msg}", "error"))
        
        threading.Thread(target=perform_check, daemon=True).start()
    
    def confirm_import(self):
        """Show confirmation dialog before real import."""
        response = messagebox.askyesno(
            "Confirm Database Update",
            "⚠️  WARNING  ⚠️\n\n"
            "You are about to write changes to the database.\n"
            "This action CANNOT be undone.\n\n"
            "Are you sure you want to proceed?",
            icon='warning'
        )
        
        if response:
            self.start_process(dry_run=False)
    
    def start_process(self, dry_run: bool):
        """Start the import process."""
        # Validation
        if not self.full_input_path or not os.path.exists(self.full_input_path):
            messagebox.showerror(
                "File Error",
                "Please select a valid Input Excel file."
            )
            return
        
        if not self.full_config_path or not os.path.exists(self.full_config_path):
            messagebox.showerror(
                "File Error",
                "Please select a valid Configuration Excel file."
            )
            return
        
        if not self.import_service:
            messagebox.showerror(
                "Configuration Error",
                "Import service not initialized. Please check database configuration."
            )
            return
        
        # Clear log
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        
        # Disable buttons
        self.toggle_buttons("disabled")
        
        # Show progress
        self.show_progress("Starting import process...")
        
        # Run in background thread
        def run_import():
            try:
                result = self.import_service.run_import(
                    self.full_input_path,
                    self.full_config_path,
                    dry_run,
                    progress_callback=lambda msg: self.root.after(0, lambda: self.show_progress(msg))
                )
                
                # Show completion message
                if result.success:
                    mode = "Simulation" if dry_run else "Import"
                    msg = f"{mode} completed successfully!\n\n"
                    msg += f"Total updates detected: {result.total_detected_updates}\n"
                    
                    if not dry_run:
                        msg += f"Total records updated: {result.total_committed_updates}"
                    
                    self.root.after(0, lambda: messagebox.showinfo("Success", msg))
                else:
                    self.root.after(
                        0,
                        lambda: messagebox.showerror(
                            "Error",
                            f"Import failed:\n{result.error_message}"
                        )
                    )
            
            except Exception as e:
                logger.exception("Error during import")
                self.root.after(
                    0,
                    lambda: messagebox.showerror("Error", f"Unexpected error:\n{str(e)}")
                )
            
            finally:
                # Re-enable buttons and hide progress
                self.root.after(0, lambda: self.toggle_buttons("normal"))
                self.root.after(0, self.hide_progress)
        
        threading.Thread(target=run_import, daemon=True).start()


def run_gui(config_file: Optional[str] = None):
    """Run the GUI application."""
    root = tk.Tk()
    app = DBImporterApp(root, config_file)
    root.mainloop()
