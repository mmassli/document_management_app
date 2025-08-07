"""
Logging utilities for the application.
"""

import datetime
import tkinter as tk


class LoggingMixin:
    """Mixin class providing logging functionality"""

    def __init__(self):
        self.verbose_logging = False  # Control detailed logging

    def log_message(self, message, level="INFO"):
        """Log a message to the console with timestamp"""
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        
        # Filter out detailed logs unless verbose mode is enabled
        if not self.verbose_logging and self._is_detailed_log(message):
            return
            
        # Simplify emoji-heavy messages
        simplified_message = self._simplify_message(message)
        
        try:
            self.console.insert(tk.END, f"[{timestamp}] {simplified_message}\n")
            self.console.see(tk.END)
        except tk.TclError:
            # Console widget might be destroyed, ignore logging errors
            pass

    def _is_detailed_log(self, message):
        """Check if a message is a detailed log that should be filtered out"""
        detailed_indicators = [
            "🔍", "📊", "🚀", "✅", "❌", "⚠️", "🟢", "🔗", "📆", "ℹ️"
        ]
        return any(indicator in message for indicator in detailed_indicators)

    def _simplify_message(self, message):
        """Simplify verbose messages to essential information"""
        # Remove emojis and simplify common verbose patterns
        simplified = message
        
        # Replace emoji patterns with simple text
        emoji_replacements = {
            "🔍": "INFO:",
            "📊": "INFO:",
            "🚀": "INFO:",
            "✅": "SUCCESS:",
            "❌": "ERROR:",
            "⚠️": "WARNING:",
            "🟢": "INFO:",
            "🔗": "INFO:",
            "📆": "INFO:",
            "ℹ️": "INFO:"
        }
        
        for emoji, replacement in emoji_replacements.items():
            simplified = simplified.replace(emoji, replacement)
        
        # Simplify verbose patterns
        if "Detected V1.0 file:" in simplified:
            simplified = simplified.replace("Detected V1.0 file:", "V1.0 file detected:")
            
        if "Updating Excel tracking with doc_prefix:" in simplified:
            simplified = simplified.replace("Updating Excel tracking with doc_prefix:", "Updating Excel tracking:")
            
        if "Using enhanced operation for V1.0 file:" in simplified:
            simplified = simplified.replace("Using enhanced operation for V1.0 file:", "Enhanced operation:")
            
        if "Excel tracking updated successfully for" in simplified:
            simplified = simplified.replace("Excel tracking updated successfully for", "Excel tracking updated:")
            
        if "Successfully processed" in simplified:
            simplified = simplified.replace("Successfully processed", "Processed:")
            
        if "Could not remove attachment file:" in simplified:
            simplified = simplified.replace("Could not remove attachment file:", "Could not remove file:")
            
        if "Processing completed with" in simplified:
            simplified = simplified.replace("Processing completed with", "Processing completed:")
            
        if "All" in simplified and "unique files processed successfully" in simplified:
            simplified = simplified.replace("unique files processed successfully", "files processed")
            
        if "Critical error during processing:" in simplified:
            simplified = simplified.replace("Critical error during processing:", "Critical error:")
            
        if "Starting Excel tracking update..." in simplified:
            simplified = simplified.replace("Starting Excel tracking update...", "Updating Excel tracking...")
            
        if "Excel tracking file not found or not specified" in simplified:
            simplified = simplified.replace("Excel tracking file not found or not specified", "Excel file not found")
            
        if "Target directory not specified" in simplified:
            simplified = simplified.replace("Target directory not specified", "Target directory missing")
            
        if "Excel file updated and saved:" in simplified:
            simplified = simplified.replace("Excel file updated and saved:", "Excel file saved:")
            
        if "Excel file saved:" in simplified:
            simplified = simplified.replace("Excel file saved:", "Excel saved:")
            
        if "Error reading Excel file:" in simplified:
            simplified = simplified.replace("Error reading Excel file:", "Excel read error:")
            
        if "Error writing Excel file:" in simplified:
            simplified = simplified.replace("Error writing Excel file:", "Excel write error:")
            
        if "Error getting Excel info:" in simplified:
            simplified = simplified.replace("Error getting Excel info:", "Excel info error:")
            
        if "Error creating Excel summary:" in simplified:
            simplified = simplified.replace("Error creating Excel summary:", "Excel summary error:")
            
        if "Data exported to Excel:" in simplified:
            simplified = simplified.replace("Data exported to Excel:", "Data exported:")
            
        if "Error exporting to Excel:" in simplified:
            simplified = simplified.replace("Error exporting to Excel:", "Export error:")
            
        if "Checking deadlines for department" in simplified:
            simplified = simplified.replace("Checking deadlines for department", "Checking deadlines:")
            
        if "Cannot access file:" in simplified:
            simplified = simplified.replace("Cannot access file:", "File access error:")
            
        if "No matching sheet found for department" in simplified:
            simplified = simplified.replace("No matching sheet found for department", "No matching sheet:")
            
        if "Found matching deadline:" in simplified:
            simplified = simplified.replace("Found matching deadline:", "Found deadline:")
            
        if "Error processing row" in simplified:
            simplified = simplified.replace("Error processing row", "Row processing error:")
            
        if "No deadlines found for department" in simplified:
            simplified = simplified.replace("No deadlines found for department", "No deadlines found:")
            
        if "Generated department Excel with" in simplified:
            simplified = simplified.replace("Generated department Excel with", "Generated Excel:")
            
        if "Error generating deadline Excel:" in simplified:
            simplified = simplified.replace("Error generating deadline Excel:", "Excel generation error:")
            
        if "Deadline email sent successfully for" in simplified:
            simplified = simplified.replace("Deadline email sent successfully for", "Email sent:")
            
        if "Failed to send deadline email for" in simplified:
            simplified = simplified.replace("Failed to send deadline email for", "Email failed:")
            
        if "Error sending deadline email:" in simplified:
            simplified = simplified.replace("Error sending deadline email:", "Email error:")
            
        if "win32com not available, falling back to SMTP" in simplified:
            simplified = simplified.replace("win32com not available, falling back to SMTP", "Using SMTP fallback")
            
        if "Error sending email via Windows COM:" in simplified:
            simplified = simplified.replace("Error sending email via Windows COM:", "COM email error:")
            
        if "SMTP email sending not configured" in simplified:
            simplified = simplified.replace("SMTP email sending not configured", "SMTP not configured")
            
        if "Error sending email via SMTP:" in simplified:
            simplified = simplified.replace("Error sending email via SMTP:", "SMTP error:")
            
        if "Processing department:" in simplified:
            simplified = simplified.replace("Processing department:", "Processing:")
            
        if "Error processing department" in simplified:
            simplified = simplified.replace("Error processing department", "Department error:")
            
        if "Error sending all department deadlines:" in simplified:
            simplified = simplified.replace("Error sending all department deadlines:", "Batch email error:")
            
        if "Half-year tracking status reset" in simplified:
            simplified = simplified.replace("Half-year tracking status reset", "Status reset")
            
        if "Error resetting half-year status:" in simplified:
            simplified = simplified.replace("Error resetting half-year status:", "Reset error:")
            
        if "Error showing tracking status:" in simplified:
            simplified = simplified.replace("Error showing tracking status:", "Status error:")
            
        if "Generated:" in simplified:
            simplified = simplified.replace("Generated:", "Created:")
            
        if "Error generating Excel for" in simplified:
            simplified = simplified.replace("Error generating Excel for", "Excel error:")
            
        if "Error generating deadline Excel files:" in simplified:
            simplified = simplified.replace("Error generating deadline Excel files:", "Excel generation error:")
            
        if "Email sent successfully for" in simplified:
            simplified = simplified.replace("Email sent successfully for", "Email sent:")
            
        if "Failed to send email for" in simplified:
            simplified = simplified.replace("Failed to send email for", "Email failed:")
            
        if "Error sending email for" in simplified:
            simplified = simplified.replace("Error sending email for", "Email error:")
            
        if "Error sending deadline emails:" in simplified:
            simplified = simplified.replace("Error sending deadline emails:", "Email error:")
            
        if "Error in generate and send workflow:" in simplified:
            simplified = simplified.replace("Error in generate and send workflow:", "Workflow error:")
            
        if "Configuration loaded - persistent directories restored" in simplified:
            simplified = simplified.replace("Configuration loaded - persistent directories restored", "Configuration loaded")
            
        if "Error loading config:" in simplified:
            simplified = simplified.replace("Error loading config:", "Config error:")
            
        if "Default directories set" in simplified:
            simplified = simplified.replace("Default directories set", "Using default directories")
            
        if "Error saving config:" in simplified:
            simplified = simplified.replace("Error saving config:", "Config save error:")
            
        if "Spire.Doc watermark added to" in simplified:
            simplified = simplified.replace("Spire.Doc watermark added to", "Watermark added:")
            
        if "Error:" in simplified and "Spire.Doc" in simplified:
            simplified = simplified.replace("Error: Spire.Doc", "Watermark error:")
            
        if "File not found:" in simplified:
            simplified = simplified.replace("File not found:", "File missing:")
            
        if "Multiple PDF files found for" in simplified:
            simplified = simplified.replace("Multiple PDF files found for", "Multiple PDFs found:")
            
        if "Processed" in simplified and "unique files from" in simplified:
            simplified = simplified.replace("unique files from", "files from")
            
        if "No duplicates found - processing all" in simplified:
            simplified = simplified.replace("No duplicates found - processing all", "No duplicates - processing")
            
        if "Error processing duplicate files:" in simplified:
            simplified = simplified.replace("Error processing duplicate files:", "Duplicate processing error:")
            
        if "Could not remove file:" in simplified:
            simplified = simplified.replace("Could not remove file:", "File removal failed:")
            
        if "Verifying copy from" in simplified:
            simplified = simplified.replace("Verifying copy from", "Verifying:")
            
        if "Hash mismatch:" in simplified:
            simplified = simplified.replace("Hash mismatch:", "Verification failed:")
            
        if "Error accessing Outlook:" in simplified:
            simplified = simplified.replace("Error accessing Outlook:", "Outlook error:")
            
        if "Preview error:" in simplified:
            simplified = simplified.replace("Preview error:", "Preview failed:")
            
        if "Adding watermarks to" in simplified and "documents..." in simplified:
            simplified = simplified.replace("documents...", "files")
            
        if "Archiving completed with watermarks for" in simplified:
            simplified = simplified.replace("Archiving completed with watermarks for", "Archiving completed:")
            
        if "V1.0 file detected:" in simplified and "no previous versions to archive" in simplified:
            simplified = simplified.replace("no previous versions to archive", "no versions to archive")
            
        if "No files to archive for" in simplified:
            simplified = simplified.replace("No files to archive for", "No files to archive:")
            
        if "Replacement verification failed - files differ" in simplified:
            simplified = simplified.replace("Replacement verification failed - files differ", "Verification failed")
            
        if "Excel tracking update failed or no match found for" in simplified:
            simplified = simplified.replace("Excel tracking update failed or no match found for", "Excel update failed:")
            
        if "Error updating Excel tracking for" in simplified:
            simplified = simplified.replace("Error updating Excel tracking for", "Excel tracking error:")
            
        if "Attachment file" in simplified and "removed from local computer after processing." in simplified:
            simplified = simplified.replace("removed from local computer after processing.", "removed")
            
        if "Error processing" in simplified and ":" in simplified:
            # Keep the filename but simplify the message
            parts = simplified.split(":")
            if len(parts) >= 3:
                filename = parts[1].strip()
                error = parts[2].strip()
                simplified = f"ERROR: Processing {filename}: {error}"
        
        return simplified

    def update_status(self, message):
        """Update the status bar message"""
        self.status_message.config(text=message)

    def clear_logs(self):
        """Clear the console and history"""
        self.console.delete(1.0, tk.END)
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        self.log_message("Logs cleared")
        self.update_status("Ready")

    def record_operation(self, operation, status, details):
        """Record an operation in the history"""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.history_tree.insert("", tk.END, values=(timestamp, operation, status, details))
        self.operation_history.append((timestamp, operation, status, details))

    def set_verbose_logging(self, enabled):
        """Enable or disable verbose logging"""
        self.verbose_logging = enabled
        if enabled:
            self.log_message("INFO: Verbose logging enabled")
        else:
            self.log_message("INFO: Verbose logging disabled")