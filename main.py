#!/usr/bin/env python3
"""
Enhanced File Replacer & Archiver - Professional Edition with Outlook Integration
Main entry point for the application.
"""

import tkinter as tk
from gui.app import FileReplacerApp
from pathlib import Path
import os
import sys
import logging

# Set working directory to script/exe location
if getattr(sys, 'frozen', False):
    os.chdir(os.path.dirname(sys.executable))
else:
    os.chdir(os.path.dirname(__file__))

# Set up logging
logging.basicConfig(filename='app_debug.log', level=logging.DEBUG)
logging.info('Application started. Working directory: %s', os.getcwd())


def main():
    root = tk.Tk()
    app = FileReplacerApp(root)

    # Set close handler
    root.protocol("WM_DELETE_WINDOW", app.on_close)

    root.mainloop()


if __name__ == "__main__":
    main()