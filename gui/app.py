"""
Main application class for File Replacer & Archiver.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from pathlib import Path
import shutil
import json
import os
import datetime
import threading
import hashlib
import tempfile


from gui.dialogs import OutlookAttachmentDialog, ProgressDialog
from gui.styles import ModernStyle
from logic.file_ops import FileOperations
from logic.excel_ops import ExcelOperations
from logic.deadline_tracker import DeadlineTracker

from logic.config import ConfigManager
from utils.outlook import OUTLOOK_AVAILABLE
from utils.logging import LoggingMixin
from gui.scrollable_frame import VerticalScrolledFrame

CONFIG_FILE = 'file_replacer_config.json'
DEFAULT_TARGET = str(Path.home() / 'Desktop')
DEFAULT_ARCHIVE = str(Path.home() / 'Desktop' / 'Archive')


class FileReplacerApp(LoggingMixin):
    def __init__(self, root):
        # Initialize LoggingMixin
        super().__init__()
        
        self.root = root
        self.root.title("File Replacer & Archiver - Professional Edition")
        self.root.minsize(800, 600)
        self.root.geometry("900x700")

        # Initialize variables
        self.dark_mode = False
        self.operation_history = []
        self.current_operation = None
        self.file_ops = FileOperations(self)
        self.excel_ops = ExcelOperations(self)
        self.deadline_tracker = DeadlineTracker(self)
        self.config_manager = ConfigManager(self)
        
        # Initialize option variables (backend functionality always enabled)
        self.verify_backup_var = tk.BooleanVar(value=True)
        self.create_log_var = tk.BooleanVar(value=True)
        self.show_preview_var = tk.BooleanVar(value=True)
        self.theme_var = tk.StringVar(value="light")

        # Set up the application
        self.setup_styles()
        self.setup_ui()
        self.config_manager.load_config()
        self.setup_keyboard_shortcuts()
        # Add this line to force UI update
        self.root.update_idletasks()
        self.apply_theme()
        self.set_window_icon()
        
        # Check for half-year reminders
        self.deadline_tracker.check_and_prompt_halfyear_reminder()

    def set_window_icon(self):
        """Set application icon if available"""
        try:
            self.root.iconbitmap("app_icon.ico")
        except:
            pass

    def setup_styles(self):
        """Configure ttk styles for modern appearance"""
        self.style = ttk.Style()
        # Base styles
        self.style.configure('.', font=ModernStyle.NORMAL_FONT)
        self.style.configure('TFrame', background=ModernStyle.LIGHT_BG)
        self.style.configure('TLabel', padding=2)
        self.style.configure('TButton', padding=6, anchor='center')
        # Button styles
        self.style.configure('Accent.TButton',
                            font=ModernStyle.HEADER_FONT,
                            foreground='black',  # Changed to black
                            background=ModernStyle.LIGHT_ACCENT)
        self.style.configure('Success.TButton',
                            foreground='black',  # Changed to black
                            background=ModernStyle.LIGHT_SUCCESS)
        self.style.configure('Warning.TButton',
                            foreground='black',  # Changed to black
                            background=ModernStyle.LIGHT_WARNING)
        # Frame styles
        self.style.configure('Card.TFrame',
                            relief='solid',
                            borderwidth=1,
                            padding=10)
        # Label styles
        self.style.configure('Header.TLabel',
                            font=ModernStyle.HEADER_FONT)
        self.style.configure('Title.TLabel',
                            font=ModernStyle.TITLE_FONT)
        # Make sure buttons maintain minimum width
        self.style.layout('TButton',
            [('Button.button', {'children':
                [('Button.focus', {'children':
                    [('Button.padding', {'children':
                        [('Button.label', {'sticky': 'nswe'})],
                    'sticky': 'nswe'})],
                'sticky': 'nswe'})],
            'sticky': 'nswe'})])

    def setup_ui(self):
        # Create main container with padding
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        # Create a scrollable frame for better layout control
        self.scrollable_frame = VerticalScrolledFrame(self.main_frame)
        self.scrollable_frame.pack(fill=tk.BOTH, expand=True)
        self.content_frame = self.scrollable_frame.interior
        # Application title
        self.create_title_section()
        # Add separator
        ttk.Separator(self.content_frame).pack(fill=tk.X, pady=10)
        # File selection section
        self.create_file_selection_section()
        # Output console
        self.create_output_console()
        # Place action buttons directly in main_frame, above status bar
        self.create_action_buttons()
        # Status bar (pack directly into main_frame, not scrollable)
        self.create_status_bar()
        # Bind resize handler for responsive button layout
        self.root.bind('<Configure>', self.on_resize_action_buttons)
        # Force initial layout calculation
        self.root.update_idletasks()
        self.root.after(100, self.finalize_layout)

    def finalize_layout(self):
        """Final adjustments after initial layout"""
        # Ensure minimum window size
        self.root.update_idletasks()
        self.root.minsize(self.root.winfo_width(), self.root.winfo_height())
        # Center window on screen
        self.center_window()
        # Make sure all widgets are visible
        self.root.update()

    def center_window(self):
        """Center the window on screen"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def create_title_section(self):
        """Create professional title section"""
        title_frame = ttk.Frame(self.content_frame)
        title_frame.pack(fill=tk.X, pady=(0, 15))

        title_label = ttk.Label(title_frame, text="File Replacer & Archiver",
                                style="Title.TLabel")
        title_label.pack(side=tk.LEFT)

        version_label = ttk.Label(title_frame, text="v2.0 Professional Edition",
                                  font=ModernStyle.NORMAL_FONT)
        version_label.pack(side=tk.LEFT, padx=(10, 0))

        # Removed the separator under the title

    def create_file_selection_section(self):
        """Enhanced file selection with validation indicators"""
        section_frame = ttk.LabelFrame(self.content_frame, text="File Operations",
                                       padding="15", style="Card.TFrame")
        section_frame.pack(fill=tk.X, pady=(0, 15))

        # Configure grid weights for better responsiveness
        section_frame.columnconfigure(1, weight=1)

        # Attachment file with validation - ENHANCED WITH OUTLOOK BUTTON
        self.create_file_input_row_with_outlook(section_frame, 0, "Attachment Files:",
                                                "attachment", self.browse_file)

        # Excel file selection with validation
        self.create_excel_input_row(section_frame, 1, "Excel File:",
                                    "excel", self.browse_excel_file)

        # Target directory with validation
        self.create_file_input_row(section_frame, 2, "Target Directory:",
                                   "target", self.browse_directory)

        # Archive directory with validation
        self.create_file_input_row(section_frame, 3, "Archive Directory:",
                                   "archive", self.browse_directory)



    def create_file_input_row_with_outlook(self, parent, row, label_text, field_name, browse_command):
        """Create a file input row with Outlook button for attachment selection"""
        # Label
        label = ttk.Label(parent, text=label_text, font=ModernStyle.NORMAL_FONT)
        label.grid(row=row, column=0, sticky=tk.W, pady=5, padx=(0, 10))

        # Entry with frame for styling
        entry_frame = ttk.Frame(parent)
        entry_frame.grid(row=row, column=1, sticky=tk.EW, pady=5, padx=(0, 10))
        entry_frame.columnconfigure(0, weight=1)

        entry = ttk.Entry(entry_frame, font=ModernStyle.NORMAL_FONT)
        entry.grid(row=0, column=0, sticky=tk.EW)

        # Validation indicator
        indicator = ttk.Label(entry_frame, text="", width=3)
        indicator.grid(row=0, column=1, padx=(5, 0))

        # Button frame for multiple buttons
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=row, column=2, padx=5, pady=5)

        # Browse button
        browse_btn = ttk.Button(button_frame, text="Browse Files...",
                                command=lambda: browse_command(entry))
        browse_btn.pack(side=tk.LEFT, padx=(0, 5))

        # Outlook button
        outlook_btn = ttk.Button(button_frame, text="📧 Outlook",
                                 command=lambda: self.browse_outlook_attachment(entry))
        outlook_btn.pack(side=tk.LEFT)

        # Disable Outlook button if not available
        if not OUTLOOK_AVAILABLE:
            outlook_btn.config(state=tk.DISABLED)

        # Store references
        setattr(self, f"{field_name}_entry", entry)
        setattr(self, f"{field_name}_indicator", indicator)

        # Bind validation
        entry.bind('<KeyRelease>', lambda e: self.validate_input(field_name))
        entry.bind('<FocusOut>', lambda e: self.validate_input(field_name))

    def create_file_input_row(self, parent, row, label_text, field_name, browse_command):
        """Create a file input row with validation indicator"""
        # Label
        label = ttk.Label(parent, text=label_text, font=ModernStyle.NORMAL_FONT)
        label.grid(row=row, column=0, sticky=tk.W, pady=5, padx=(0, 10))

        # Entry with frame for styling
        entry_frame = ttk.Frame(parent)
        entry_frame.grid(row=row, column=1, sticky=tk.EW, pady=5, padx=(0, 10))
        entry_frame.columnconfigure(0, weight=1)

        entry = ttk.Entry(entry_frame, font=ModernStyle.NORMAL_FONT)
        entry.grid(row=0, column=0, sticky=tk.EW)

        # Validation indicator
        indicator = ttk.Label(entry_frame, text="", width=3)
        indicator.grid(row=0, column=1, padx=(5, 0))

        # Browse button
        browse_btn = ttk.Button(parent, text="Browse...",
                                command=lambda: browse_command(entry))
        browse_btn.grid(row=row, column=2, padx=5, pady=5)

        # Store references
        setattr(self, f"{field_name}_entry", entry)
        setattr(self, f"{field_name}_indicator", indicator)

        # Bind validation
        entry.bind('<KeyRelease>', lambda e: self.validate_input(field_name))
        entry.bind('<FocusOut>', lambda e: self.validate_input(field_name))

    def browse_outlook_attachment(self, entry_widget):
        """Browse for attachment from Outlook"""
        if not OUTLOOK_AVAILABLE:
            messagebox.showerror("Outlook Not Available",
                                 "Outlook integration requires the pywin32 package.\n"
                                 "Install it with: pip install pywin32")
            return

        try:
            # Show Outlook attachment dialog
            dialog = OutlookAttachmentDialog(self.root)
            self.root.wait_window(dialog.dialog)

            # Check if attachment was selected
            if hasattr(dialog, 'selected_attachment') and dialog.selected_attachment:
                entry_widget.delete(0, tk.END)
                entry_widget.insert(0, dialog.selected_attachment)
                self.validate_input("attachment")

                # Show success message
                if ";" in dialog.selected_attachment:
                    # Multiple files selected
                    file_count = len(dialog.selected_attachment.split(";"))
                    self.update_status(f"Selected {file_count} Outlook attachments")
                    self.log_message(f"📧 Selected {file_count} attachments from Outlook")
                else:
                    # Single file selected
                    filename = Path(dialog.selected_attachment).name
                    self.update_status(f"Selected Outlook attachment: {filename}")
                    self.log_message(f"📧 Selected attachment from Outlook: {filename}")

        except Exception as e:
            error_msg = str(e)
            self.log_message(f"❌ Error accessing Outlook: {error_msg}")
            messagebox.showerror("Outlook Error", f"Failed to access Outlook:\n{error_msg}")

    # def create_options_section(self):
    #     """Create options section with advanced settings - REMOVED FROM UI"""
    #     # Options functionality is now always enabled in backend
    #     # Variables are initialized in __init__ with default values:
    #     # - verify_backup_var = True
    #     # - create_log_var = True  
    #     # - show_preview_var = True
    #     # - theme_var = "light" (managed by toggle_theme method)
    #     pass

    def create_output_console(self):
        """Create enhanced output console with tabs"""
        console_frame = ttk.LabelFrame(self.content_frame, text="Output & Logs", padding=10)
        console_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        # Configure grid weights
        console_frame.columnconfigure(0, weight=1)
        console_frame.rowconfigure(0, weight=1)
        # Create notebook for different views
        self.console_notebook = ttk.Notebook(console_frame)
        self.console_notebook.grid(row=0, column=0, sticky='nsew')
        # Console tab
        console_tab = ttk.Frame(self.console_notebook)
        self.console_notebook.add(console_tab, text="Console")
        # Console text widget with scrollbar
        self.console = scrolledtext.ScrolledText(
            console_tab,
            wrap=tk.WORD,
            font=ModernStyle.CONSOLE_FONT,
            padx=5,
            pady=5
        )
        self.console.pack(fill=tk.BOTH, expand=True)
        # History tab
        history_tab = ttk.Frame(self.console_notebook)
        self.console_notebook.add(history_tab, text="History")

        # History treeview
        history_columns = ("Timestamp", "Operation", "Status", "Files")
        self.history_tree = ttk.Treeview(history_tab, columns=history_columns, show="headings")

        # Configure columns
        self.history_tree.heading("Timestamp", text="Timestamp")
        self.history_tree.heading("Operation", text="Operation")
        self.history_tree.heading("Status", text="Status")
        self.history_tree.heading("Files", text="Files")

        self.history_tree.column("Timestamp", width=150)
        self.history_tree.column("Operation", width=120)
        self.history_tree.column("Status", width=100)
        self.history_tree.column("Files", width=300)

        # Scrollbars
        history_scrollbar_y = ttk.Scrollbar(history_tab, orient=tk.VERTICAL, command=self.history_tree.yview)
        history_scrollbar_x = ttk.Scrollbar(history_tab, orient=tk.HORIZONTAL, command=self.history_tree.xview)
        self.history_tree.configure(yscrollcommand=history_scrollbar_y.set, xscrollcommand=history_scrollbar_x.set)

        # Pack treeview and scrollbars
        self.history_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        history_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        history_scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

        # Bind selection event
        self.history_tree.bind("<<TreeviewSelect>>", self.on_history_select)

    def create_action_buttons(self):
        """Create action buttons with robust layout"""
        self.button_frame = ttk.Frame(self.main_frame, padding=(0, 10, 0, 0))
        # Pack at the bottom, above the status bar
        self.button_frame.pack(fill=tk.X, side=tk.BOTTOM)
        button_container = ttk.Frame(self.button_frame)
        button_container.pack(expand=True, fill=tk.X)
        self.action_buttons = []  # Store references for later use
        buttons = [
            ("Process Files", "Accent.TButton", self.process_files),
            ("Preview Changes", "Success.TButton", self.preview_changes),
            ("Clear Logs", "Warning.TButton", self.clear_logs),
            ("Verbose Logs: OFF", "Warning.TButton", self.toggle_verbose_logging),
            ("Deadline Status", "Success.TButton", self.show_deadline_status),
            ("Reset Deadlines", "Warning.TButton", self.reset_deadline_status),
            ("Help", None, self.show_help)
        ]
        for text, style, command in buttons:
            btn = ttk.Button(
                button_container,
                text=text,
                style=style,
                command=command,
                padding=(10, 5)
            )
            btn.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
            self.action_buttons.append(btn)
        # self.button_frame.pack_propagate(False)  # Allow frame to expand to fit buttons

    def layout_action_buttons(self, vertical=False):
        # Remove all buttons from frame
        for btn in self.action_buttons:
            btn.pack_forget()
        if vertical:
            for btn in self.action_buttons:
                btn.pack(fill=tk.X, pady=2, padx=0)
        else:
            for i, btn in enumerate(self.action_buttons):
                side = tk.LEFT if i < len(self.action_buttons) - 1 else tk.RIGHT
                btn.pack(side=side, padx=(0, 10) if side == tk.LEFT else 0)

    def on_resize_action_buttons(self, event):
        # Dynamically stack buttons vertically if window is too narrow
        min_width_for_horizontal = 600
        if event.widget == self.root:
            if self.root.winfo_width() < min_width_for_horizontal:
                self.layout_action_buttons(vertical=True)
            else:
                self.layout_action_buttons(vertical=False)

    def create_status_bar(self):
        """Create enhanced status bar"""
        self.status_bar = ttk.Frame(self.main_frame, height=25)
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM, anchor='s')

        # Status message
        self.status_message = ttk.Label(self.status_bar, text="Ready", font=ModernStyle.NORMAL_FONT)
        self.status_message.pack(side=tk.LEFT, padx=10)

        # Version info
        ttk.Label(self.status_bar, text="v2.0 Professional Edition", font=ModernStyle.NORMAL_FONT).pack(side=tk.RIGHT,
                                                                                                        padx=10)

    def setup_keyboard_shortcuts(self):
        """Setup keyboard shortcuts"""
        self.root.bind('<F1>', lambda e: self.show_help())
        self.root.bind('<Control-p>', lambda e: self.preview_changes())
        self.root.bind('<Control-r>', lambda e: self.process_files())
        self.root.bind('<Control-l>', lambda e: self.clear_logs())
        self.root.bind('<Control-d>', lambda e: self.toggle_theme())

    def validate_input(self, field_name):
        """Validate input fields with visual feedback"""
        entry = getattr(self, f"{field_name}_entry")
        indicator = getattr(self, f"{field_name}_indicator")
        path = entry.get().strip()

        if not path:
            indicator.config(text="⚠️", foreground="orange")
            return False

        if field_name == "attachment":
            # Handle multiple files separated by semicolon
            if ";" in path:
                file_paths = path.split(";")
                valid = all(os.path.isfile(p.strip()) for p in file_paths)
            else:
                valid = os.path.isfile(path)
        else:
            valid = os.path.isdir(path)

        if valid:
            indicator.config(text="✓", foreground="green")
        else:
            indicator.config(text="✗", foreground="red")

        return valid

    def all_inputs_valid(self):
        """Check if all inputs are valid"""
        return (self.validate_input("attachment") and
                self.validate_excel_input("excel") and
                self.validate_input("target") and
                self.validate_input("archive"))

    def browse_file(self, entry_widget):
        """Browse for files (single or multiple)"""
        file_paths = filedialog.askopenfilenames(title="Select Attachment Files")
        if file_paths:
            # Join multiple file paths with semicolon separator
            file_paths_str = ";".join(file_paths)
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, file_paths_str)
            self.validate_input("attachment")

    def browse_directory(self, entry_widget):
        """Browse for a directory"""
        dir_path = filedialog.askdirectory(title="Select Directory")
        if dir_path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, dir_path)
            if entry_widget == self.target_entry:
                self.validate_input("target")
                self.config_manager.save_config()  # Auto-save persistent directory
                self.log_message(f"📁 Target directory updated and saved: {Path(dir_path).name}")
            elif entry_widget == self.archive_entry:
                self.validate_input("archive")
                self.config_manager.save_config()  # Auto-save persistent directory
                self.log_message(f"📁 Archive directory updated and saved: {Path(dir_path).name}")

    def browse_excel_file(self, entry_widget):
        """Browse for an Excel file"""
        return self.excel_ops.browse_excel_file(entry_widget)

    def create_excel_input_row(self, parent, row, label_text, field_name, browse_command):
        """Create an Excel file input row with validation indicator and preview button"""
        # Label
        label = ttk.Label(parent, text=label_text, font=ModernStyle.NORMAL_FONT)
        label.grid(row=row, column=0, sticky=tk.W, pady=5, padx=(0, 10))

        # Entry with frame for styling
        entry_frame = ttk.Frame(parent)
        entry_frame.grid(row=row, column=1, sticky=tk.EW, pady=5, padx=(0, 10))
        entry_frame.columnconfigure(0, weight=1)

        entry = ttk.Entry(entry_frame, font=ModernStyle.NORMAL_FONT)
        entry.grid(row=0, column=0, sticky=tk.EW)

        # Validation indicator
        indicator = ttk.Label(entry_frame, text="", width=3)
        indicator.grid(row=0, column=1, padx=(5, 0))

        # Button frame for multiple buttons
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=row, column=2, padx=5, pady=5)

        # Browse button
        browse_btn = ttk.Button(button_frame, text="Browse...",
                                command=lambda: browse_command(entry))
        browse_btn.pack(side=tk.LEFT)

        # Store references
        setattr(self, f"{field_name}_entry", entry)
        setattr(self, f"{field_name}_indicator", indicator)

        # Bind validation
        entry.bind('<KeyRelease>', lambda e: self.validate_excel_input(field_name))
        entry.bind('<FocusOut>', lambda e: self.validate_excel_input(field_name))

    def validate_excel_input(self, field_name):
        """Validate Excel input fields with visual feedback"""
        entry = getattr(self, f"{field_name}_entry")
        indicator = getattr(self, f"{field_name}_indicator")
        path = entry.get().strip()

        if not path:
            indicator.config(text="⚠️", foreground="orange")
            return False

        # Validate Excel file
        is_valid, message = self.excel_ops.validate_excel_file(path)
        
        if is_valid:
            indicator.config(text="✓", foreground="green")
        else:
            indicator.config(text="✗", foreground="red")

        return is_valid



    def toggle_theme(self):
        """Toggle between light and dark theme"""
        self.dark_mode = not self.dark_mode
        self.apply_theme()

    def apply_theme(self):
        """Apply the current theme to all widgets"""
        if self.dark_mode:
            bg = ModernStyle.DARK_BG
            fg = ModernStyle.DARK_FG
            accent = ModernStyle.DARK_ACCENT
            border = ModernStyle.DARK_BORDER
            self.theme_var.set("dark")
        else:
            bg = ModernStyle.LIGHT_BG
            fg = ModernStyle.LIGHT_FG
            accent = ModernStyle.LIGHT_ACCENT
            border = ModernStyle.LIGHT_BORDER
            self.theme_var.set("light")

        # Apply colors to widgets
        self.root.configure(background=bg)
        self.main_frame.configure(style="TFrame")

        # Update style configurations
        self.style.configure('.', background=bg, foreground=fg)
        self.style.configure('TFrame', background=bg)
        self.style.configure('TLabel', background=bg, foreground=fg)
        self.style.configure('TEntry', fieldbackground="white", foreground="black")
        self.style.configure('TButton', background=accent, foreground=fg)
        self.style.configure('Treeview', background="white", foreground="black",
                             fieldbackground="white")
        self.style.configure('Treeview.Heading', background=border, foreground=fg)
        self.style.map('Treeview', background=[('selected', accent)],
                       foreground=[('selected', 'white')])

        # Update console colors
        self.console.configure(background="white" if not self.dark_mode else "#2d2d2d",
                               foreground="black" if not self.dark_mode else "#e0e0e0",
                               insertbackground="black" if not self.dark_mode else "white")

    def update_status(self, message):
        """Update the status bar message"""
        self.status_message.config(text=message)

    def clear_logs(self):
        """
         Clears the console and history logs in the application.

         This method performs the following actions:
         - Deletes all text from the console widget.
         - Removes all entries from the history treeview.
         - Logs a message indicating that the logs have been cleared.
         - Updates the status bar to display "Ready".

         Parameters:
         - None

         Returns:
         - None
         """
        self.console.delete(1.0, tk.END)
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        self.log_message("Logs cleared")
        self.update_status("Ready")

    def show_help(self):
        """Show help information"""
        help_text = """File Replacer & Archiver - Help

        1. Select one or more attachment files you want to deploy (multiple files supported)
        2. Choose an Excel file for data processing (optional)
        3. Choose the target directory where files will be replaced
        4. Specify an archive directory for backups
        5. Preview changes before processing
        6. Click 'Process Files' to execute (files will be processed sequentially)

        Multiple File Processing:
        - Select multiple files by holding Ctrl or Shift when browsing
        - Outlook attachment selection also supports multiple files (hold Ctrl to select multiple)
        - Files are processed one after another in the order selected
        - Progress is shown for each file being processed
        - If one file fails, processing continues with the remaining files
        - Final summary shows success/failure count

        Duplicate File Handling:
        - Files with the same name (ignoring extensions) are grouped together
        - If multiple versions exist (e.g., contract.pdf, contract.docx, contract.xlsx):
          * PDF versions are prioritized if available
          * If no PDF exists, any other version is used
        - Only one file per unique name is processed
        - Skipped duplicates are logged for transparency

        Keyboard Shortcuts:
        - F1: Show this help
        - Ctrl+P: Preview changes
        - Ctrl+R: Process files
        - Ctrl+L: Clear logs
        - Ctrl+D: Toggle dark mode

        Logging Options:
        - Verbose Logs button: Toggle detailed logging mode
        - When OFF: Only essential logs are shown
        - When ON: All detailed logs with emojis are displayed

        Professional Features:
        - Multiple file selection and batch processing
        - Outlook integration for attachment selection
        - Excel file support with pandas and openpyxl
        - Excel file preview and validation
        - Document watermarking (adds "UNGÜLTIG" in red to Word, PDF, and Excel files with image-based watermarking for Excel using ungültig_transparent.png)
        - Comprehensive logging and history
        - File verification and backup integrity checks
        - Dark/light theme support
        - Half-year deadline tracking and reminders
        - Automatic deadline email generation and sending
        """
        messagebox.showinfo("Help", help_text)

    def show_deadline_status(self):
        """Show deadline tracking status"""
        self.deadline_tracker.show_tracking_status()
    
    def reset_deadline_status(self):
        """Reset deadline tracking status"""
        response = messagebox.askyesno(
            "Reset Deadline Status",
            "Are you sure you want to reset all deadline tracking status?\n\n"
            "This will clear all sent records and allow you to send deadline emails again.",
            icon='warning'
        )
        if response:
            self.deadline_tracker.reset_halfyear_status()
            messagebox.showinfo("Reset Complete", "Deadline tracking status has been reset.")

    def toggle_verbose_logging(self):
        """Toggle verbose logging mode"""
        self.verbose_logging = not self.verbose_logging
        self.set_verbose_logging(self.verbose_logging)
        
        # Update button text
        for btn in self.action_buttons:
            if btn.cget("text") == "Verbose Logs" or btn.cget("text").startswith("Verbose Logs:"):
                btn.config(text="Verbose Logs: " + ("ON" if self.verbose_logging else "OFF"))
                break

    def preview_changes(self):
        """Preview what files will be replaced"""
        if not self.all_inputs_valid():
            messagebox.showwarning("Invalid Input", "Please check all input fields")
            return

        try:
            attachment_paths_str = self.attachment_entry.get()
            target_dir = Path(self.target_entry.get())
            archive_dir = Path(self.archive_entry.get())

            # Handle multiple files
            if ";" in attachment_paths_str:
                attachment_paths = [Path(p.strip()) for p in attachment_paths_str.split(";")]
            else:
                attachment_paths = [Path(attachment_paths_str)]

            # Process duplicates for preview
            self.log_message(f"🔄 Preview: Processing {len(attachment_paths)} files for duplicates...")
            files_to_process = self.file_ops.process_duplicate_files(attachment_paths)

            # Build preview text
            preview_text = f"File Replacement Preview:\n\n"
            preview_text += f"Total files selected: {len(attachment_paths)}\n"
            preview_text += f"Unique files to process: {len(files_to_process)}\n\n"

            for i, attachment_path in enumerate(files_to_process, 1):
                preview_text += f"File {i}: {attachment_path.name}\n"
                preview_text += f"Size: {self.format_file_size(attachment_path.stat().st_size)}\n"
                preview_text += f"Modified: {datetime.datetime.fromtimestamp(attachment_path.stat().st_mtime)}\n"

                # Find matching files in target directory by first 10 characters AND same extension
                attachment_prefix = attachment_path.name[:10]
                attachment_ext = attachment_path.suffix.lower()
                matching_files = [
                    f for f in target_dir.iterdir()
                    if f.is_file() and f.name[:10] == attachment_prefix and f.suffix.lower() == attachment_ext
                ]

                if matching_files:
                    preview_text += f"Matching files to replace: {len(matching_files)}\n"
                    for f in matching_files[:3]:  # Show first 3 matches
                        preview_text += f"  - {f.name}\n"
                    if len(matching_files) > 3:
                        preview_text += f"  ... and {len(matching_files) - 3} more\n"
                else:
                    preview_text += "No matching files found\n"
                preview_text += "\n"

            preview_text += f"Backup will be saved to: {archive_dir}"

            self.log_message(f"Previewing replacement for {len(attachment_paths)} files")
            messagebox.showinfo("Replacement Preview", preview_text)

        except Exception as e:
            self.log_message(f"❌ Preview error: {str(e)}")
            messagebox.showerror("Preview Error", f"Failed to generate preview:\n{str(e)}")

    def format_file_size(self, size_bytes):
        """Format file size in human readable format"""
        if size_bytes == 0:
            return "0 B"
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.1f} TB"

    def process_files(self):
        """Process the file replacement operation"""
        if not self.all_inputs_valid():
            messagebox.showwarning("Invalid Input", "Please check all input fields")
            return

        # Create progress dialog
        progress = ProgressDialog(self.root, "Processing File Replacement")

        # Run in separate thread to keep UI responsive
        threading.Thread(target=self._process_files_thread, args=(progress,), daemon=True).start()

    def _process_files_thread(self, progress):
        """Thread worker for file processing"""
        try:
            # Get paths from UI
            attachment_paths_str = self.attachment_entry.get()
            target_dir = Path(self.target_entry.get())
            archive_dir = Path(self.archive_entry.get())

            # Handle multiple files
            if ";" in attachment_paths_str:
                attachment_paths = [Path(p.strip()) for p in attachment_paths_str.split(";")]
            else:
                attachment_paths = [Path(attachment_paths_str)]

            # Process duplicates and get unique files to process
            self.log_message(f"🔄 Processing {len(attachment_paths)} files for duplicates...")
            files_to_process = self.file_ops.process_duplicate_files(attachment_paths)
            
            total_files = len(files_to_process)
            processed_files = 0
            failed_files = []

            progress.update_progress(0, f"Starting processing of {total_files} unique files...")
            self.log_message(f"Starting file replacement process for {total_files} unique files")

            # Create archive directory if it doesn't exist
            archive_dir.mkdir(parents=True, exist_ok=True)

            for i, attachment_path in enumerate(files_to_process):
                try:
                    # Update progress for current file
                    progress_percent = int((i / total_files) * 100)
                    progress.update_progress(progress_percent, f"Processing file {i+1}/{total_files}: {attachment_path.name}")
                    
                    self.log_message(f"Processing file {i+1}/{total_files}: {attachment_path.name}")
                    
                    target_file = target_dir / attachment_path.name

                    # Check for exact filename match in target directory
                    if target_file.exists():
                        error_msg = f"A file with the exact name '{attachment_path.name}' already exists in the target directory. Overwriting is not allowed."
                        self.log_message(f"❌ Error: {error_msg}")
                        failed_files.append((attachment_path.name, error_msg))
                        continue

                    # Find matching files in target directory by first 10 characters AND same extension
                    attachment_prefix = attachment_path.name[:10]
                    attachment_ext = attachment_path.suffix.lower()
                    matching_files = [
                        f for f in target_dir.iterdir()
                        if f.is_file() and f.name[:10] == attachment_prefix and f.suffix.lower() == attachment_ext
                    ]

                    # Check if this is a V1.0 file (character beside "V" is "1") and no matching files found
                    is_v1_file = False
                    v_index = attachment_path.name.find('V')
                    if v_index != -1 and v_index + 1 < len(attachment_path.name):
                        char_beside_v = attachment_path.name[v_index + 1]
                        if char_beside_v == '1':
                            is_v1_file = True
                            self.log_message(f"🔍 Detected V1.0 file: {attachment_path.name}")

                    # Archive all matching files with watermark for Word, PDF, and Excel documents
                    if matching_files:
                        self.log_message(f"Adding watermarks to {len(matching_files)} documents...")
                        self.file_ops.archive_files(matching_files, archive_dir)
                        self.log_message(f"Archiving completed with watermarks for {attachment_path.name}")
                    else:
                        if is_v1_file:
                            self.log_message(f"ℹ️ V1.0 file detected: {attachment_path.name} - no previous versions to archive")
                        else:
                            self.log_message(f"No files to archive for {attachment_path.name}")

                    # After archiving, copy the new file
                    self.file_ops.copy_file(attachment_path, target_file)

                    # Verify replacement if option is enabled
                    if self.verify_backup_var.get():
                        if not self.file_ops.verify_file_copy(attachment_path, target_file):
                            raise Exception("Replacement verification failed - files differ")

                    # Excel tracking functionality
                    try:
                        # Extract doc_prefix from filename (first 10 characters)
                        doc_prefix = attachment_path.name[:10]
                        self.log_message(f"🔍 Updating Excel tracking with doc_prefix: '{doc_prefix}'")
                        
                        # Check if enhanced operation should be used (V1.0 file with no matching files)
                        use_enhanced_operation = is_v1_file and not matching_files
                        if use_enhanced_operation:
                            self.log_message(f"🚀 Using enhanced operation for V1.0 file: {attachment_path.name}")
                        
                        # Call Excel tracking update
                        success = self.excel_ops.update_excel_tracking(doc_prefix, attachment_path.name)
                        if success:
                            self.log_message(f"✅ Excel tracking updated successfully for {attachment_path.name}")
                        else:
                            self.log_message(f"⚠️ Excel tracking update failed or no match found for {attachment_path.name}")
                    except Exception as e:
                        self.log_message(f"❌ Error updating Excel tracking for {attachment_path.name}: {str(e)}")

                    processed_files += 1
                    self.log_message(f"✅ Successfully processed {attachment_path.name}")
                    self.record_operation("Replace", "Success", attachment_path.name)

                    # Remove the original attachment file if it is a local file (not from Outlook)
                    attachment_path_str = str(attachment_path)
                    temp_dir = tempfile.gettempdir()
                    if not attachment_path_str.startswith(temp_dir):
                        try:
                            os.remove(attachment_path_str)
                            self.log_message(f"Attachment file '{attachment_path_str}' removed from local computer after processing.")
                        except Exception as e:
                            self.log_message(f"⚠️ Could not remove attachment file: {e}")

                except Exception as e:
                    error_msg = str(e)
                    self.log_message(f"❌ Error processing {attachment_path.name}: {error_msg}")
                    failed_files.append((attachment_path.name, error_msg))
                    self.record_operation("Replace", "Failed", f"{attachment_path.name}: {error_msg}")

            # Final progress update
            progress.update_progress(100, "Processing completed")
            
            # Show final results
            if failed_files:
                failed_summary = "\n".join([f"- {name}: {error}" for name, error in failed_files])
                self.log_message(f"⚠️ Processing completed with {len(failed_files)} failures out of {total_files} unique files")
                self.root.after(0, lambda: messagebox.showwarning(
                    "Processing Completed with Errors",
                    f"Processed {processed_files}/{total_files} unique files successfully.\n\n"
                    f"Failed files:\n{failed_summary}"
                ))
            else:
                self.log_message(f"✅ All {total_files} unique files processed successfully")
                self.root.after(0, lambda: messagebox.showinfo(
                    "Success", 
                    f"All {total_files} unique files processed successfully!"
                ))

        except Exception as e:
            error_msg = str(e)
            self.log_message(f"❌ Critical error during processing: {error_msg}")
            self.root.after(0, lambda: messagebox.showerror("Error", f"Critical error during processing:\n{error_msg}"))

        finally:
            progress.close()

    def verify_file_copy(self, source, destination):
        """Verify that two files are identical by comparing hashes"""
        self.log_message(f"Verifying copy from {source} to {destination}")

        if not os.path.exists(destination):
            return False

        if os.path.getsize(source) != os.path.getsize(destination):
            return False

        # Compare file hashes
        source_hash = self.calculate_file_hash(source)
        dest_hash = self.calculate_file_hash(destination)

        if source_hash != dest_hash:
            self.log_message(f"Hash mismatch: {source_hash} != {dest_hash}")
            return False

        return True

    def calculate_file_hash(self, filepath):
        """Calculate SHA-256 hash of a file"""
        sha256 = hashlib.sha256()
        with open(filepath, 'rb') as f:
            for block in iter(lambda: f.read(4096), b''):
                sha256.update(block)
        return sha256.hexdigest()

    def record_operation(self, operation, status, details):
        """Record an operation in the history"""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.history_tree.insert("", tk.END, values=(timestamp, operation, status, details))
        self.operation_history.append((timestamp, operation, status, details))

    def on_history_select(self, event):
        """Handle history item selection"""
        selection = self.history_tree.selection()
        if selection:
            item = self.history_tree.item(selection[0])
            details = "\n".join(
                f"{k}: {v}" for k, v in zip(["Timestamp", "Operation", "Status", "Details"], item['values']))
            messagebox.showinfo("History Details", details)



    def on_close(self):
        """Handle application close"""
        self.config_manager.save_config()
        self.root.destroy()

