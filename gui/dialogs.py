import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import os
import tempfile
from pathlib import Path
import datetime
from tkcalendar import DateEntry

from .styles import ModernStyle
from .scrollable_frame import VerticalScrolledFrame
from utils.outlook import OUTLOOK_AVAILABLE
import win32com.client
import logging


class ExcelCellInputDialog:
    """Dialog for entering values in Excel columns E, F, and G"""
    
    def __init__(self, parent, found_row_data=None, new_row_data=None, document_info=None):
        self.parent = parent
        self.found_row_data = found_row_data or {}
        self.new_row_data = new_row_data or {}
        self.document_info = document_info or {}
        self.result = None
        
        # Create modal dialog
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Excel Cell Input - Columns E, F, G")
        self.dialog.geometry("700x650")  # Increased height to accommodate all elements
        self.dialog.resizable(True, True)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Set minimum size to ensure buttons are visible
        self.dialog.minsize(600, 500)
        
        # Center the dialog
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (350)
        y = (self.dialog.winfo_screenheight() // 2) - (325)
        self.dialog.geometry(f"700x650+{x}+{y}")
        
        self.setup_ui()
        
        # Ensure the dialog is properly configured
        self.dialog.update_idletasks()
    
    def setup_ui(self):
        """Setup the dialog UI"""
        # Create main container frame
        container_frame = ttk.Frame(self.dialog)
        container_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # Create scrollable frame
        scrollable_frame = VerticalScrolledFrame(container_frame)
        scrollable_frame.pack(fill=tk.BOTH, expand=True)
        
        # Use the interior of the scrollable frame as the main frame
        main_frame = scrollable_frame.interior
        
        # Title
        title_label = ttk.Label(main_frame, text="Enter Values for Columns Gültig ab, Gesperrt ab, Letzte Überprüfung",
                                font=ModernStyle.TITLE_FONT)
        title_label.pack(pady=(0, 10))
        
        # Document information section
        if self.document_info:
            self.setup_document_info_section(main_frame)
        
        # Create a frame for the notebook and buttons
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create notebook for found row and new row
        self.notebook = ttk.Notebook(content_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Found row tab (only show if found_row_data is provided)
        if self.found_row_data:
            self.found_frame = ttk.Frame(self.notebook)
            self.notebook.add(self.found_frame, text="Found Row")
            self.setup_found_row_ui()
        
        # New row tab
        self.new_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.new_frame, text="New Row")
        self.setup_new_row_ui()
        
        # Buttons - ensure they're at the bottom
        button_frame = ttk.Frame(content_frame)
        button_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(10, 0))
        
        # OK button with better styling and positioning
        try:
            ok_button = ttk.Button(
                button_frame, 
                text="OK",
                command=self.confirm_input,
                style="Accent.TButton",  # Use accent style for better visibility
                padding=(20, 8)  # Add padding for better button size
            )
        except tk.TclError:
            # Fallback if Accent.TButton style is not available
            ok_button = ttk.Button(
                button_frame, 
                text="OK",
                command=self.confirm_input,
                padding=(20, 8)  # Add padding for better button size
            )
        
        ok_button.pack(side=tk.RIGHT, padx=(10, 0))
        
        # Cancel button
        cancel_button = ttk.Button(
            button_frame, 
            text="Cancel",
            command=self.dialog.destroy,
            padding=(20, 8)  # Add padding for better button size
        )
        cancel_button.pack(side=tk.RIGHT)
        
        # Force update to ensure buttons are rendered
        self.dialog.update_idletasks()
    
    def setup_document_info_section(self, parent):
        """Setup document information display section"""
        # Document info frame
        doc_frame = ttk.LabelFrame(parent, text="📄 Document Information", padding="10")
        doc_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Create a grid for document info
        info_frame = ttk.Frame(doc_frame)
        info_frame.pack(fill=tk.X)
        
        # Document filename
        if 'filename' in self.document_info:
            filename = self.document_info['filename']
            ttk.Label(info_frame, text="Document:", font=ModernStyle.HEADER_FONT).grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
            filename_label = ttk.Label(info_frame, text=filename, font=ModernStyle.NORMAL_FONT, foreground=ModernStyle.LIGHT_ACCENT)
            filename_label.grid(row=0, column=1, sticky=tk.W)
            # Make filename bold for emphasis
            filename_label.config(font=("Segoe UI", 10, "bold"))
        
        # Document prefix
        if 'doc_prefix' in self.document_info:
            doc_prefix = self.document_info['doc_prefix']
            ttk.Label(info_frame, text="Prefix:", font=ModernStyle.HEADER_FONT).grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
            ttk.Label(info_frame, text=doc_prefix, font=ModernStyle.NORMAL_FONT).grid(row=1, column=1, sticky=tk.W, pady=(5, 0))
        
        # Sheet name
        if 'sheet_name' in self.document_info:
            sheet_name = self.document_info['sheet_name']
            ttk.Label(info_frame, text="Sheet:", font=ModernStyle.HEADER_FONT).grid(row=2, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
            ttk.Label(info_frame, text=sheet_name, font=ModernStyle.NORMAL_FONT).grid(row=2, column=1, sticky=tk.W, pady=(5, 0))
        
        # Row number
        if 'row_num' in self.document_info:
            row_num = self.document_info['row_num']
            ttk.Label(info_frame, text="Row:", font=ModernStyle.HEADER_FONT).grid(row=3, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
            ttk.Label(info_frame, text=str(row_num), font=ModernStyle.NORMAL_FONT).grid(row=3, column=1, sticky=tk.W, pady=(5, 0))

        
        # Configure grid columns to expand
        info_frame.columnconfigure(1, weight=1)
        
        # Add a subtle separator
        separator = ttk.Separator(doc_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=(10, 0))
        
    def setup_found_row_ui(self):
        """Setup UI for found row input"""
        # Found row frame
        found_frame = ttk.LabelFrame(self.found_frame, text="Found Row Values", padding="10")
        found_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Column E
        e_frame = ttk.Frame(found_frame)
        e_frame.pack(fill=tk.X, pady=5)
        ttk.Label(e_frame, text="Gültig ab:", width=15).pack(side=tk.LEFT)
        self.found_e_var = tk.StringVar(value=self.found_row_data.get('E', ''))
        self.found_e_entry = ttk.Entry(e_frame, textvariable=self.found_e_var, width=30)
        self.found_e_entry.pack(side=tk.LEFT, padx=(5, 0))
        ttk.Button(e_frame, text="📅", command=lambda: self.show_calendar(self.found_e_var)).pack(side=tk.LEFT, padx=(5, 0))
        
        # Column F
        f_frame = ttk.Frame(found_frame)
        f_frame.pack(fill=tk.X, pady=5)
        ttk.Label(f_frame, text="Gesperrt ab:", width=15).pack(side=tk.LEFT)
        self.found_f_var = tk.StringVar(value=self.found_row_data.get('F', ''))
        self.found_f_entry = ttk.Entry(f_frame, textvariable=self.found_f_var, width=30)
        self.found_f_entry.pack(side=tk.LEFT, padx=(5, 0))
        ttk.Button(f_frame, text="📅", command=lambda: self.show_calendar(self.found_f_var)).pack(side=tk.LEFT, padx=(5, 0))
        
        # Column G
        g_frame = ttk.Frame(found_frame)
        g_frame.pack(fill=tk.X, pady=5)
        ttk.Label(g_frame, text="Letzte Überprüfung:", width=15).pack(side=tk.LEFT)
        self.found_g_var = tk.StringVar(value=self.found_row_data.get('G', ''))
        self.found_g_entry = ttk.Entry(g_frame, textvariable=self.found_g_var, width=30)
        self.found_g_entry.pack(side=tk.LEFT, padx=(5, 0))
        ttk.Button(g_frame, text="📅", command=lambda: self.show_calendar(self.found_g_var)).pack(side=tk.LEFT, padx=(5, 0))
        
        # Quick buttons
        quick_frame = ttk.Frame(found_frame)
        quick_frame.pack(fill=tk.X, pady=10)
        ttk.Button(quick_frame, text="Set 'aktuell gültig'", 
                  command=lambda: self.set_aktuell_gueltig([self.found_e_var, self.found_f_var, self.found_g_var])).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(quick_frame, text="Set '-'", 
                  command=lambda: self.set_dash([self.found_e_var, self.found_f_var, self.found_g_var])).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(quick_frame, text="Clear All", 
                  command=lambda: self.clear_all([self.found_e_var, self.found_f_var, self.found_g_var])).pack(side=tk.LEFT)
        
    def setup_new_row_ui(self):
        """Setup UI for new row input"""
        # New row frame
        new_frame = ttk.LabelFrame(self.new_frame, text="New Row Values (Gesperrt ab='aktuell gültig', Letzte Überprüfung='-')", padding="10")
        new_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Column E (Date with same format as found row)
        e_frame = ttk.Frame(new_frame)
        e_frame.pack(fill=tk.X, pady=5)
        ttk.Label(e_frame, text="Gültig ab (Date):", width=15).pack(side=tk.LEFT)
        self.new_e_var = tk.StringVar(value=self.new_row_data.get('E', ''))
        self.new_e_entry = ttk.Entry(e_frame, textvariable=self.new_e_var, width=30)
        self.new_e_entry.pack(side=tk.LEFT, padx=(5, 0))
        ttk.Button(e_frame, text="📅", command=lambda: self.show_calendar(self.new_e_var)).pack(side=tk.LEFT, padx=(5, 0))
        
        # Column F (Fixed to "aktuell gültig")
        f_frame = ttk.Frame(new_frame)
        f_frame.pack(fill=tk.X, pady=5)
        ttk.Label(f_frame, text="Gesperrt ab:", width=15).pack(side=tk.LEFT)
        self.new_f_var = tk.StringVar(value="aktuell gültig")
        self.new_f_entry = ttk.Entry(f_frame, textvariable=self.new_f_var, width=30, state='readonly')
        self.new_f_entry.pack(side=tk.LEFT, padx=(5, 0))
        ttk.Label(f_frame, text="(Fixed)", font=("Arial", 8, "italic")).pack(side=tk.LEFT, padx=(5, 0))
        
        # Column G (Fixed to "-")
        g_frame = ttk.Frame(new_frame)
        g_frame.pack(fill=tk.X, pady=5)
        ttk.Label(g_frame, text="Letzte Überprüfung:", width=15).pack(side=tk.LEFT)
        self.new_g_var = tk.StringVar(value="-")
        self.new_g_entry = ttk.Entry(g_frame, textvariable=self.new_g_var, width=30, state='readonly')
        self.new_g_entry.pack(side=tk.LEFT, padx=(5, 0))
        ttk.Label(g_frame, text="(Fixed)", font=("Arial", 8, "italic")).pack(side=tk.LEFT, padx=(5, 0))

        # Quick buttons (only for column E since F and G are fixed)
        quick_frame = ttk.Frame(new_frame)
        quick_frame.pack(fill=tk.X, pady=10)
        ttk.Button(quick_frame, text="Set Today's Date", 
                  command=self.set_todays_date).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(quick_frame, text="Clear Date", 
                  command=lambda: self.new_e_var.set("")).pack(side=tk.LEFT)
        
    def show_calendar(self, var):
        """Show calendar picker for date selection"""
        try:
            # Create a simple date picker dialog
            calendar_dialog = tk.Toplevel(self.dialog)
            calendar_dialog.title("Select Date")
            calendar_dialog.geometry("300x250")
            calendar_dialog.transient(self.dialog)
            calendar_dialog.grab_set()
            
            # Center the dialog
            calendar_dialog.update_idletasks()
            x = (calendar_dialog.winfo_screenwidth() // 2) - (150)
            y = (calendar_dialog.winfo_screenheight() // 2) - (125)
            calendar_dialog.geometry(f"300x250+{x}+{y}")
            
            # Calendar widget
            cal = DateEntry(calendar_dialog, width=20, background='darkblue',
                           foreground='white', borderwidth=2, date_pattern='dd.mm.yyyy')
            cal.pack(pady=20)
            
            # Buttons
            button_frame = ttk.Frame(calendar_dialog)
            button_frame.pack(pady=10)
            
            def set_date():
                # Get the date and format it as DD.MM.YYYY without time component
                selected_date = cal.get_date()
                formatted_date = selected_date.strftime("%d.%m.%Y")
                var.set(formatted_date)
                calendar_dialog.destroy()
            
            ttk.Button(button_frame, text="OK", command=set_date).pack(side=tk.LEFT, padx=(0, 5))
            ttk.Button(button_frame, text="Cancel", command=calendar_dialog.destroy).pack(side=tk.LEFT)
            
        except Exception as e:
            # Fallback to manual entry if calendar fails
            messagebox.showwarning("Calendar Error", f"Calendar picker not available: {str(e)}\nPlease enter date manually in format DD.MM.YYYY")
    
    def set_aktuell_gueltig(self, vars_list):
        """Set 'aktuell gültig' for all specified variables"""
        for var in vars_list:
            var.set("aktuell gültig")
    
    def set_dash(self, vars_list):
        """Set '-' for all specified variables"""
        for var in vars_list:
            var.set("-")
    
    def clear_all(self, vars_list):
        """Clear all specified variables"""
        for var in vars_list:
            var.set("")
    
    def set_todays_date(self):
        """Set today's date in the new row column E"""
        from datetime import datetime
        today = datetime.now().strftime("%d.%m.%Y")
        self.new_e_var.set(today)
    
    def confirm_input(self):
        """Confirm the input and return the result"""
        self.result = {
            'new_row': {
                'E': self.new_e_var.get().strip(),
                'F': self.new_f_var.get().strip(),
                'G': self.new_g_var.get().strip()
            }
        }
        
        # Add found row data only if it exists
        if self.found_row_data:
            self.result['found_row'] = {
                'E': self.found_e_var.get().strip(),
                'F': self.found_f_var.get().strip(),
                'G': self.found_g_var.get().strip()
            }
        
        self.dialog.destroy()


class OutlookAttachmentDialog:
    """Dialog for selecting attachments from Outlook emails"""

    def __init__(self, parent):
        self.parent = parent
        self.selected_attachment = None
        self.outlook = None

        # Create modal dialog
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Select Outlook Attachment")
        self.dialog.geometry("800x600")
        self.dialog.resizable(True, True)
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # Center the dialog
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (400)
        y = (self.dialog.winfo_screenheight() // 2) - (300)
        self.dialog.geometry(f"800x600+{x}+{y}")

        self.setup_ui()
        self.load_outlook_emails()

    def setup_ui(self):
        """Setup the dialog UI"""
        main_frame = ttk.Frame(self.dialog, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title_label = ttk.Label(main_frame, text="Select Attachment from Outlook",
                                font=ModernStyle.TITLE_FONT)
        title_label.pack(pady=(0, 15))

        # Status label
        self.status_label = ttk.Label(main_frame, text="Loading emails...",
                                      font=ModernStyle.NORMAL_FONT)
        self.status_label.pack(pady=(0, 10))

        # Create notebook for different views
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        # Emails tab
        self.emails_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.emails_frame, text="Recent Emails")

        # Email list
        self.setup_email_list()

        # Attachments tab
        self.attachments_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.attachments_frame, text="Attachments")

        # Attachment list
        self.setup_attachment_list()

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)

        ttk.Button(button_frame, text="Cancel",
                   command=self.dialog.destroy).pack(side=tk.RIGHT, padx=(10, 0))
        self.select_button = ttk.Button(button_frame, text="Select Attachment(s)",
                                        command=self.select_attachment, state=tk.DISABLED)
        self.select_button.pack(side=tk.RIGHT)

        ttk.Button(button_frame, text="Refresh",
                   command=self.load_outlook_emails).pack(side=tk.LEFT)

    def setup_email_list(self):
        """Setup email list with treeview"""
        # Email list frame
        list_frame = ttk.Frame(self.emails_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Email treeview
        columns = ("Subject", "From", "Date", "Attachments")
        self.email_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=10)

        # Configure columns
        self.email_tree.heading("Subject", text="Subject")
        self.email_tree.heading("From", text="From")
        self.email_tree.heading("Date", text="Date")
        self.email_tree.heading("Attachments", text="Attachments")

        self.email_tree.column("Subject", width=300)
        self.email_tree.column("From", width=200)
        self.email_tree.column("Date", width=120)
        self.email_tree.column("Attachments", width=100)

        # Scrollbars
        email_scrollbar_y = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.email_tree.yview)
        email_scrollbar_x = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.email_tree.xview)
        self.email_tree.configure(yscrollcommand=email_scrollbar_y.set, xscrollcommand=email_scrollbar_x.set)

        # Pack treeview and scrollbars
        self.email_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        email_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        email_scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

        # Bind selection event
        self.email_tree.bind("<<TreeviewSelect>>", self.on_email_select)

        # Email preview
        preview_frame = ttk.LabelFrame(self.emails_frame, text="Email Preview", padding="10")
        preview_frame.pack(fill=tk.X, padx=5, pady=5)

        self.email_preview = scrolledtext.ScrolledText(preview_frame, height=8,
                                                       font=ModernStyle.CONSOLE_FONT)
        self.email_preview.pack(fill=tk.BOTH, expand=True)

    def setup_attachment_list(self):
        """Setup attachment list"""
        # Attachment list frame
        att_frame = ttk.Frame(self.attachments_frame)
        att_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Attachment treeview with multiple selection enabled
        att_columns = ("Name", "Size", "Type", "Email Subject")
        self.attachment_tree = ttk.Treeview(att_frame, columns=att_columns, show="headings", height=15, selectmode="extended")

        # Configure columns
        self.attachment_tree.heading("Name", text="Attachment Name")
        self.attachment_tree.heading("Size", text="Size")
        self.attachment_tree.heading("Type", text="Type")
        self.attachment_tree.heading("Email Subject", text="From Email")

        self.attachment_tree.column("Name", width=250)
        self.attachment_tree.column("Size", width=100)
        self.attachment_tree.column("Type", width=100)
        self.attachment_tree.column("Email Subject", width=300)

        # Scrollbars
        att_scrollbar_y = ttk.Scrollbar(att_frame, orient=tk.VERTICAL, command=self.attachment_tree.yview)
        att_scrollbar_x = ttk.Scrollbar(att_frame, orient=tk.HORIZONTAL, command=self.attachment_tree.xview)
        self.attachment_tree.configure(yscrollcommand=att_scrollbar_y.set, xscrollcommand=att_scrollbar_x.set)

        # Pack treeview and scrollbars
        self.attachment_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        att_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        att_scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

        # Bind selection event
        self.attachment_tree.bind("<<TreeviewSelect>>", self.on_attachment_select)

    def load_outlook_emails(self):
        """Load emails from Outlook"""
        if not OUTLOOK_AVAILABLE:
            self.status_label.config(text="❌ Outlook COM interface not available")
            messagebox.showerror("Outlook Not Available",
                                 "The Outlook integration requires the pywin32 package.\n"
                                 "Install it with: pip install pywin32")
            return

        try:
            self.status_label.config(text="🔄 Connecting to Outlook...")
            self.dialog.update()

            # Connect to Outlook
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = self.outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox

            # Get recent emails with attachments
            messages = inbox.Items
            messages.Sort("[ReceivedTime]", True)  # Sort by received time, newest first

            # Clear existing items
            for item in self.email_tree.get_children():
                self.email_tree.delete(item)
            for item in self.attachment_tree.get_children():
                self.attachment_tree.delete(item)

            self.status_label.config(text="📧 Loading emails with Excel/PDF/Word attachments...")
            self.dialog.update()

            # Process emails
            email_count = 0
            attachment_count = 0

            for i, message in enumerate(messages):
                if i >= 50:  # Limit to 50 most recent emails
                    break

                try:
                    if message.Attachments.Count > 0:
                        # Add email to tree
                        subject = message.Subject or "(No Subject)"
                        sender = message.SenderName or "(Unknown Sender)"
                        date = message.ReceivedTime.strftime("%Y-%m-%d %H:%M")
                        att_count = message.Attachments.Count

                        # Use iid for index
                        email_item = self.email_tree.insert("", tk.END, iid=str(i), values=(
                            subject[:50] + "..." if len(subject) > 50 else subject,
                            sender[:30] + "..." if len(sender) > 30 else sender,
                            date,
                            str(att_count)
                        ))

                        # Add attachments to attachment tree (only Excel, PDF, and Word files)
                        for att in message.Attachments:
                            file_extension = Path(att.FileName).suffix.lower()
                            # Only show Excel, PDF, and Word files
                            if file_extension in ['.xlsx', '.xls', '.pdf', '.docx', '.doc']:
                                att_size = getattr(att, 'Size', 0)
                                att_type = Path(att.FileName).suffix.upper() or "FILE"
                                # Use iid as "emailIndex:attIndex"
                                att_item = self.attachment_tree.insert("", tk.END, iid=f"{i}:{att.Index}", values=(
                                    att.FileName,
                                    self.format_file_size(att_size),
                                    att_type,
                                    subject[:40] + "..." if len(subject) > 40 else subject
                                ))

                                attachment_count += 1

                        email_count += 1

                        if email_count % 5 == 0:
                            self.status_label.config(
                                text=f"📧 Loaded {email_count} emails, {attachment_count} Excel/PDF/Word attachments...")
                            self.dialog.update()

                except Exception as e:
                    continue  # Skip problematic emails

            self.status_label.config(text=f"✅ Loaded {email_count} emails with {attachment_count} Excel/PDF/Word attachments")

        except Exception as e:
            error_msg = str(e)
            self.status_label.config(text=f"❌ Error loading Outlook data: {error_msg}")
            messagebox.showerror("Outlook Error", f"Failed to load Outlook data:\n{error_msg}")

    def format_file_size(self, size_bytes):
        """Format file size in human readable format"""
        if size_bytes == 0:
            return "0 B"
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.1f} TB"

    def on_email_select(self, event):
        """Handle email selection"""
        selection = self.email_tree.selection()
        if not selection:
            return

        try:
            # Get email index from iid
            email_index = int(selection[0])

            # Get message
            namespace = self.outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)
            messages = inbox.Items
            messages.Sort("[ReceivedTime]", True)

            message = messages[email_index]

            # Update preview
            self.email_preview.delete(1.0, tk.END)

            preview_text = f"📧 EMAIL DETAILS\n{'=' * 50}\n\n"
            preview_text += f"Subject: {message.Subject}\n"
            preview_text += f"From: {message.SenderName} <{message.SenderEmailAddress}>\n"
            preview_text += f"Date: {message.ReceivedTime}\n"
            preview_text += f"Attachments: {message.Attachments.Count}\n\n"

            if message.Attachments.Count > 0:
                preview_text += "📎 ATTACHMENTS:\n"
                for i, att in enumerate(message.Attachments):
                    att_size = getattr(att, 'Size', 0)
                    preview_text += f"  {i + 1}. {att.FileName} ({self.format_file_size(att_size)})\n"
                preview_text += "\n"

            # Add body preview (first 500 characters)
            body = message.Body or ""
            if len(body) > 500:
                body = body[:500] + "..."
            preview_text += f"MESSAGE PREVIEW:\n{'-' * 20}\n{body}"

            self.email_preview.insert(tk.END, preview_text)

        except Exception as e:
            self.email_preview.delete(1.0, tk.END)
            self.email_preview.insert(tk.END, f"Error loading email preview: {str(e)}")

    def on_attachment_select(self, event):
        """Handle attachment selection"""
        selection = self.attachment_tree.selection()
        if selection:
            self.select_button.config(state=tk.NORMAL)
        else:
            self.select_button.config(state=tk.DISABLED)

    def select_attachment(self):
        """Select and save the chosen attachments"""
        selection = self.attachment_tree.selection()
        if not selection:
            return

        try:
            selected_attachments = []
            
            for att_ref in selection:
                # Get attachment reference from iid
                email_index, att_index = map(int, att_ref.split(':'))

                # Get message and attachment
                namespace = self.outlook.GetNamespace("MAPI")
                inbox = namespace.GetDefaultFolder(6)
                messages = inbox.Items
                messages.Sort("[ReceivedTime]", True)

                message = messages[email_index]
                attachment = message.Attachments.Item(att_index)

                # Create temporary file
                temp_dir = tempfile.mkdtemp()
                temp_file = os.path.join(temp_dir, attachment.FileName)

                # Save attachment
                attachment.SaveAsFile(temp_file)
                
                selected_attachments.append(temp_file)

            # Store the paths (single file or multiple files)
            if len(selected_attachments) == 1:
                self.selected_attachment = selected_attachments[0]
            else:
                # Join multiple files with semicolon separator (same as browser selection)
                self.selected_attachment = ";".join(selected_attachments)

            if len(selected_attachments) == 1:
                messagebox.showinfo("Success", f"Attachment '{Path(selected_attachments[0]).name}' selected successfully!")
            else:
                messagebox.showinfo("Success", f"{len(selected_attachments)} attachments selected successfully!")
            
            self.dialog.destroy()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to select attachment(s):\n{str(e)}")

class ProgressDialog:
    """Professional progress dialog with cancellation support"""

    def __init__(self, parent, title="Processing..."):
        self.parent = parent
        self.cancelled = False

        # Create modal dialog
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("400x220")  # Increased for better button text visibility
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # Center the dialog
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (400 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (220 // 2)
        self.dialog.geometry(f"400x220+{x}+{y}")

        # Create UI elements
        self.setup_ui()

    def setup_ui(self):
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Status label
        self.status_label = ttk.Label(main_frame, text="Preparing operation...",
                                      font=ModernStyle.NORMAL_FONT)
        self.status_label.pack(pady=(0, 10))

        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var,
                                            maximum=100, length=300)
        self.progress_bar.pack(pady=(0, 10))

        # Progress percentage
        self.percent_label = ttk.Label(main_frame, text="0%", font=ModernStyle.NORMAL_FONT)
        self.percent_label.pack(pady=(0, 15))

        # Button frame for better positioning
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(5, 0))

        # Cancel button with explicit styling and better positioning
        try:
            self.cancel_button = ttk.Button(
                button_frame, 
                text="Cancel",
                command=self.cancel_operation,
                style="Accent.TButton",  # Use accent style for better visibility
                padding=(20, 8)  # Add padding for better button size
            )
        except tk.TclError:
            # Fallback if Accent.TButton style is not available
            self.cancel_button = ttk.Button(
                button_frame, 
                text="Cancel",
                command=self.cancel_operation,
                padding=(20, 8)  # Add padding for better button size
            )
        
        # Ensure button is visible and properly configured with adequate spacing
        self.cancel_button.pack(expand=True, fill=tk.X, padx=10, pady=5)
        
        # Force update to ensure button is rendered
        self.dialog.update_idletasks()

        # Bind close event
        self.dialog.protocol("WM_DELETE_WINDOW", self.cancel_operation)

    def update_progress(self, percentage, status=""):
        if not self.cancelled:
            self.progress_var.set(percentage)
            self.percent_label.config(text=f"{percentage:.1f}%")
            if status:
                self.status_label.config(text=status)
            self.dialog.update()

    def cancel_operation(self):
        self.cancelled = True
        self.dialog.destroy()

    def close(self):
        if not self.cancelled:
            self.dialog.destroy()

