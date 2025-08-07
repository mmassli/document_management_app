"""
Configuration management for the application.
"""

import json
import os
from pathlib import Path

# Constants
CONFIG_FILE = "file_replacer_config.json"
DEFAULT_TARGET = str(Path.home() / "Desktop")
DEFAULT_ARCHIVE = str(Path.home() / "Desktop" / "Archive")
LOG_FILE = "file_replacer.log"


class ConfigManager:
    """Manages application configuration"""

    def __init__(self, app):
        self.app = app

    def load_config(self):
        """Load configuration from file"""
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r') as f:
                    config = json.load(f)

                # NOTE: attachment (Link Attachment Directory) is intentionally NOT loaded
                # It should be cleared on each app restart
                
                # Load persistent directories only
                self.app.excel_entry.insert(0, config.get('excel', ''))
                self.app.target_entry.insert(0, config.get('target', DEFAULT_TARGET))
                self.app.archive_entry.insert(0, config.get('archive', DEFAULT_ARCHIVE))

                self.app.dark_mode = config.get('dark_mode', False)
                self.app.theme_var.set("dark" if self.app.dark_mode else "light")
                self.app.apply_theme()

                self.app.log_message("✅ Configuration loaded - persistent directories restored")
            else:
                # Set default values if no config file exists
                self.app.target_entry.insert(0, DEFAULT_TARGET)
                self.app.archive_entry.insert(0, DEFAULT_ARCHIVE)
                self.app.log_message("📁 Default directories set")
        except Exception as e:
            self.app.log_message(f"⚠️ Error loading config: {str(e)}")

    def save_config(self):
        """Save configuration to file"""
        try:
            config = {
                # NOTE: attachment (Link Attachment Directory) is intentionally NOT saved
                # It should be cleared on each app restart
                
                # Save only persistent directories
                'excel': self.app.excel_entry.get(),
                'target': self.app.target_entry.get(),
                'archive': self.app.archive_entry.get(),
                
                'dark_mode': self.app.dark_mode
            }

            with open(CONFIG_FILE, 'w') as f:
                json.dump(config, f, indent=2)

            self.app.log_message("💾 Configuration saved - persistent directories stored")
        except Exception as e:
            self.app.log_message(f"⚠️ Error saving config: {str(e)}")