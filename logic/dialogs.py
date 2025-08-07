import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import os
import tempfile
from pathlib import Path
import datetime
from gui.styles import ModernStyle
from utils.outlook import OUTLOOK_AVAILABLE
import win32com.client
import logging


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
        self.select_button = ttk.Button(button_frame, text="Select Attachment",
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

        # Attachment treeview
        att_columns = ("Name", "Date", "Type", "Email Subject")
        self.attachment_tree = ttk.Treeview(att_frame, columns=att_columns, show="headings", height=15)

        # Configure columns
        self.attachment_tree.heading("Name", text="Attachment Name")
        self.attachment_tree.heading("Date", text="Email Date")
        self.attachment_tree.heading("Type", text="Type")
        self.attachment_tree.heading("Email Subject", text="From Email")

        self.attachment_tree.column("Name", width=250)
        self.attachment_tree.column("Date", width=120)
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
        logging.info("Starting to load emails from Outlook...")
        if not OUTLOOK_AVAILABLE:
            self.status_label.config(text="❌ Outlook COM interface not available")
            messagebox.showerror("Outlook Not Available",
                                 "The Outlook integration requires the pywin32 package.\n"
                                 "Please install pywin32 and restart the application.")
            logging.error("Outlook COM interface not available.")
            return

        try:
            self.status_label.config(text="🔄 Connecting to Outlook...")
            logging.info("Connecting to Outlook COM...")
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = self.outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
            
            # Clear any existing filters and get ALL items
            messages = inbox.Items
            messages.Sort("[ReceivedTime]", True)  # Sort by received time, newest first
            
            # Try to get more emails by using a broader filter or no filter
            try:
                # First try to get all items without any restriction
                message_count = messages.Count
                logging.info(f"Initial fetch: {message_count} messages from Outlook.")
                
                # If we get very few messages, try to clear any view filters
                if message_count < 50:  # Suspiciously low number
                    logging.info("Low message count detected, attempting to clear filters...")
                    # Try to get items with a very broad date range
                    try:
                        # Get items from the last 10 years to ensure we get everything
                        from datetime import datetime, timedelta
                        ten_years_ago = datetime.now() - timedelta(days=3650)
                        date_filter = ten_years_ago.strftime("%m/%d/%Y")
                        messages = messages.Restrict(f"[ReceivedTime] >= '{date_filter}'")
                        message_count = messages.Count
                        logging.info(f"After date filter expansion: {message_count} messages")
                    except Exception as filter_error:
                        logging.warning(f"Date filter expansion failed: {filter_error}")
                        # Fall back to original messages
                        messages = inbox.Items
                        messages.Sort("[ReceivedTime]", True)
                        message_count = messages.Count
                        
            except Exception as e:
                logging.warning(f"Error during message count check: {e}")
                message_count = messages.Count
                
            logging.info(f"Final message count: {message_count} messages from Outlook.")

            # Clear existing items
            for item in self.email_tree.get_children():
                self.email_tree.delete(item)
            for item in self.attachment_tree.get_children():
                self.attachment_tree.delete(item)

            self.status_label.config(text="📧 Scanning all emails for filtered attachments...")
            self.dialog.update()
            
            # Log the date range we're working with
            if len(message_list) > 0:
                try:
                    first_date = message_list[0].ReceivedTime
                    last_date = message_list[-1].ReceivedTime
                    logging.info(f"Email date range: {first_date} to {last_date}")
                except Exception as date_error:
                    logging.warning(f"Could not determine date range: {date_error}")

            # Process emails - Safe iteration to handle COM object issues
            email_count = 0
            attachment_count = 0
            filtered_extensions = {'.pdf', '.doc', '.docx', '.xls', '.xlsx', '.xlsm', '.xlsb'}
            processed_count = 0

            # Create a safe list of messages to iterate over
            message_list = []
            for i in range(messages.Count):
                try:
                    message_list.append(messages.Item(i + 1))
                except Exception as e:
                    logging.warning(f"Skipping message {i+1} due to error: {e}")
                    continue

            logging.info(f"Successfully loaded {len(message_list)} messages for processing")
            
            # If we still have very few messages, try alternative approaches
            if len(message_list) < 50:
                logging.info("Still low message count, trying alternative methods...")
                try:
                    # Try getting items from different folders or with different filters
                    all_items = inbox.Items
                    all_items.Sort("[ReceivedTime]", True)
                    
                    # Try without any date restrictions
                    alternative_messages = []
                    for i in range(all_items.Count):
                        try:
                            msg = all_items.Item(i + 1)
                            alternative_messages.append(msg)
                        except Exception as e:
                            logging.warning(f"Skipping alternative message {i+1}: {e}")
                            continue
                    
                    if len(alternative_messages) > len(message_list):
                        logging.info(f"Alternative method found {len(alternative_messages)} messages vs {len(message_list)}")
                        message_list = alternative_messages
                    
                    # If still low, try accessing the inbox differently
                    if len(message_list) < 50:
                        logging.info("Trying to access inbox with different method...")
                        try:
                            # Try to get items without any sorting first
                            raw_items = inbox.Items
                            raw_messages = []
                            for i in range(raw_items.Count):
                                try:
                                    msg = raw_items.Item(i + 1)
                                    raw_messages.append(msg)
                                except Exception as e:
                                    logging.warning(f"Skipping raw message {i+1}: {e}")
                                    continue
                            
                            if len(raw_messages) > len(message_list):
                                logging.info(f"Raw method found {len(raw_messages)} messages vs {len(message_list)}")
                                message_list = raw_messages
                                
                        except Exception as raw_error:
                            logging.warning(f"Raw method failed: {raw_error}")
                        
                except Exception as alt_error:
                    logging.warning(f"Alternative method failed: {alt_error}")

            for i, message in enumerate(message_list):
                processed_count += 1
                if processed_count % 20 == 0:
                    self.status_label.config(
                        text=f"📧 Scanning... Processed {processed_count} emails, found {email_count} with filtered attachments")
                    self.dialog.update()

                try:
                    if message.Attachments.Count > 0:
                        # Check if this email has any attachments with the filtered extensions
                        has_filtered_attachments = False
                        filtered_attachments = []
                        
                        for att in message.Attachments:
                            file_extension = Path(att.FileName).suffix.lower()
                            if file_extension in filtered_extensions:
                                has_filtered_attachments = True
                                filtered_attachments.append(att)

                        if has_filtered_attachments:
                            # Add email to tree
                            subject = message.Subject or "(No Subject)"
                            sender = message.SenderName or "(Unknown Sender)"
                            date = message.ReceivedTime.strftime("%Y-%m-%d %H:%M")
                            filtered_att_count = len(filtered_attachments)

                            # Use iid for index
                            email_item = self.email_tree.insert("", tk.END, iid=str(i), values=(
                                subject[:50] + "..." if len(subject) > 50 else subject,
                                sender[:30] + "..." if len(sender) > 30 else sender,
                                date,
                                str(filtered_att_count)
                            ))

                            # Add filtered attachments to attachment tree
                            for att in filtered_attachments:
                                file_extension = Path(att.FileName).suffix.lower()
                                att_type = Path(att.FileName).suffix.upper() or "FILE"
                                email_date = message.ReceivedTime.strftime("%Y-%m-%d %H:%M")
                                # Use iid as "emailIndex:attIndex"
                                att_item = self.attachment_tree.insert("", tk.END, iid=f"{i}:{att.Index}", values=(
                                    att.FileName,
                                    email_date,
                                    att_type,
                                    subject[:40] + "..." if len(subject) > 40 else subject
                                ))
                                attachment_count += 1

                            email_count += 1

                except Exception as e:
                    logging.warning(f"Skipping problematic email at index {i}: {str(e)}")
                    continue  # Skip problematic emails

            # Show detailed results
            status_text = f"✅ Processed {processed_count} emails. Found {email_count} with {attachment_count} filtered attachments"
            if len(message_list) < 100:
                status_text += f"\n⚠️ Only {len(message_list)} total emails found - Outlook may have filters applied"
            self.status_label.config(text=status_text)
            logging.info(f"Successfully processed {processed_count} emails. Found {email_count} with {attachment_count} filtered attachments.")

        except Exception as e:
            error_msg = str(e)
            self.status_label.config(text=f"❌ Error loading Outlook data: {error_msg}")
            messagebox.showerror("Outlook Error", f"Failed to load Outlook data:\n{error_msg}")
            logging.exception("Error while fetching emails from Outlook:")

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
        """Select and save the chosen attachment"""
        selection = self.attachment_tree.selection()
        if not selection:
            return

        try:
            # Get attachment reference from iid
            att_ref = selection[0]
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

            # Store the path
            self.selected_attachment = temp_file

            messagebox.showinfo("Success", f"Attachment '{attachment.FileName}' selected successfully!")
            self.dialog.destroy()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to select attachment:\n{str(e)}")

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