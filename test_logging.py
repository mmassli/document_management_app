#!/usr/bin/env python3
"""
Test script for the new simplified logging system.
"""

import tkinter as tk
from utils.logging import LoggingMixin

class TestApp(LoggingMixin):
    def __init__(self):
        super().__init__()
        self.root = tk.Tk()
        self.root.title("Logging Test")
        self.root.geometry("600x400")
        
        # Create console widget
        self.console = tk.Text(self.root, wrap=tk.WORD, height=20)
        self.console.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create buttons
        button_frame = tk.Frame(self.root)
        button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Button(button_frame, text="Test Detailed Logs", command=self.test_detailed_logs).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Test Simple Logs", command=self.test_simple_logs).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Toggle Verbose", command=self.toggle_verbose).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Clear", command=self.clear_logs).pack(side=tk.LEFT, padx=5)
        
        # Status label
        self.status_message = tk.Label(self.root, text="Ready")
        self.status_message.pack(side=tk.BOTTOM, pady=5)
        
        self.log_message("INFO: Test application started")
        self.log_message("INFO: Verbose logging is OFF by default")

    def test_detailed_logs(self):
        """Test detailed logging messages"""
        self.log_message("🔍 Detected V1.0 file: test_file.pdf")
        self.log_message("📊 Checking deadlines for department QK")
        self.log_message("🚀 Using enhanced operation for V1.0 file: test_file.pdf")
        self.log_message("✅ Excel tracking updated successfully for test_file.pdf")
        self.log_message("⚠️ Excel tracking update failed or no match found for test_file.pdf")
        self.log_message("❌ Error updating Excel tracking for test_file.pdf: File not found")
        self.log_message("🟢 In-place update: Found 'C' row with exact document ID match")
        self.log_message("🔗 Added hyperlink to column J: /path/to/file.pdf")
        self.log_message("📆 Filtering deadlines between 2025-01-01 and 2025-12-31")
        self.log_message("ℹ️ V1.0 file detected: test_file.pdf - no previous versions to archive")

    def test_simple_logs(self):
        """Test simple logging messages"""
        self.log_message("Processing files...")
        self.log_message("File copied successfully")
        self.log_message("Configuration loaded")
        self.log_message("Error: File not found")
        self.log_message("Warning: Invalid input")

    def toggle_verbose(self):
        """Toggle verbose logging"""
        self.verbose_logging = not self.verbose_logging
        self.set_verbose_logging(self.verbose_logging)
        status = "ON" if self.verbose_logging else "OFF"
        self.log_message(f"INFO: Verbose logging {status}")

    def run(self):
        """Run the test application"""
        self.root.mainloop()

if __name__ == "__main__":
    app = TestApp()
    app.run()
