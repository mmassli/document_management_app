import os
import sys
import winshell
from win32com.client import Dispatch
import logging

# === CONFIGURATION ===
# Name of your executable (change if you renamed it with PyInstaller)
EXE_NAME = 'main.exe'  # or 'FileReplacer.exe'
# Name for the shortcut
SHORTCUT_NAME = 'File Replacer & Archiver.lnk'
# Icon file (should be .ico, same as used for the exe)
ICON_FILE = 'app_icon.ico'

# === SCRIPT ===

def get_desktop():
    return winshell.desktop()

def create_shortcut():
    desktop = get_desktop()
    exe_path = os.path.abspath(os.path.join('dist', EXE_NAME))
    icon_path = os.path.abspath(ICON_FILE)
    shortcut_path = os.path.join(desktop, SHORTCUT_NAME)

    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.Targetpath = exe_path
    shortcut.WorkingDirectory = os.path.dirname(exe_path)  # Explicitly set working directory
    shortcut.IconLocation = icon_path
    shortcut.save()
    print(f"Shortcut created: {shortcut_path}")

if __name__ == '__main__':
    if not os.path.exists(os.path.join('dist', EXE_NAME)):
        print(f"Executable not found: dist/{EXE_NAME}. Build it with PyInstaller first.")
        sys.exit(1)
    if not os.path.exists(ICON_FILE):
        print(f"Icon file not found: {ICON_FILE}. Place your .ico file in the project directory.")
        sys.exit(1)
    logging.basicConfig(filename='app_debug.log', level=logging.DEBUG)
    logging.info(f"Current working directory: {os.getcwd()}")
    logging.info(f"EXE exists: {os.path.exists('dist/main.exe')}")
    logging.info(f"Icon exists: {os.path.exists('app_icon.ico')}")
    os.chdir(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__))
    try:
        import win32com.client
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            logging.info("Connected to Outlook.")
            namespace = outlook.GetNamespace("MAPI")
            OUTLOOK_AVAILABLE = True
            logging.info("Outlook COM connection successful.")
        except Exception as e:
            OUTLOOK_AVAILABLE = False
            logging.exception("Failed to connect to Outlook:")
    except ImportError:
        OUTLOOK_AVAILABLE = False
        logging.error("pywin32 not installed.")
    print(winshell.__VERSION__)
    create_shortcut()
    print("Done.")

# USAGE:
# 1. Build your exe with PyInstaller (see previous instructions).
# 2. Place this script in your project root.
# 3. Run: python create_shortcut.py
# 4. A shortcut will appear on your desktop with the correct icon. 