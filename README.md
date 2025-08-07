# File Replacer & Archiver - Professional Edition

## Overview

**File Replacer & Archiver** is a professional Windows desktop application with a modern GUI for automating the replacement and archiving of files, with advanced integration for Microsoft Outlook attachments and Excel-based tracking. It is designed to streamline document management workflows, especially in environments where document versioning, archiving, and traceability are critical.

---

## Features

- **Modern GUI**: User-friendly interface built with Tkinter.
- **File Replacement**: Replace files in a target directory with new versions (e.g., from email attachments).
- **Archiving**: Automatically archive replaced files to a specified directory.
- **Outlook Integration**: Select attachments directly from recent Outlook emails (requires Outlook and `pywin32`). Supports multiple attachment selection.
- **Multiple File Processing**: Support for selecting and processing multiple files in a single operation:
  - Select multiple files by holding Ctrl or Shift when browsing
  - Files are processed sequentially with progress tracking
  - If one file fails, processing continues with remaining files
  - Final summary shows success/failure count
- **Excel File Support**: Comprehensive Excel file operations using pandas and openpyxl:
  - Excel file selection and validation
  - Excel data processing and file information
  - Support for .xlsx, .xls, .xlsm, and .xlsb formats
  - Multi-sheet Excel file handling
- **Document Watermarking**: Automatically adds "UNGÜLTIG" watermark in red to Word, PDF, and Excel documents before archiving. Excel files use image-based watermarking with the ungültig_transparent.png image.
- **Dark Mode**: Optional dark mode for comfortable viewing.
- **Operation History**: View and manage a history of file operations.
- **Configurable**: Remembers last-used paths and settings via a config file.

---

## Installation

### Prerequisites
- **Windows OS** (required for Outlook integration)
- **Python 3.8+**
- **Microsoft Outlook** (for attachment integration)

### Required Python Packages
Install dependencies with pip:

```bash
pip install -r requirements.txt
```

Or install individually:
```bash
pip install tkinter pandas openpyxl pywin32 winshell python-docx reportlab
```

- `tkinter` (usually included with Python)
- `pandas` (for Excel data manipulation)
- `openpyxl` (for Excel file operations)
- `pywin32` (for Outlook integration and Excel PDF conversion)
- `winshell` (for shortcut creation)
- `python-docx` (for Word document watermarking)
- `reportlab` (for PDF generation - alternative method)
- `pdf2image` (for embedded PDF viewing)
- `Pillow` (for image processing)



### Optional: Build as Standalone EXE
You can use [PyInstaller](https://www.pyinstaller.org/) to build a standalone executable:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=app_icon.ico main.py
```

---

## Usage

1. **Start the Application**
   - Run `python main.py` or launch the built EXE.

2. **Select Files and Folders**
   - Choose one or more attachment files (or use the Outlook button to pick from recent emails). Both browser selection and Outlook selection support multiple files and will be processed sequentially.
   - Select an Excel file for data processing (optional, with preview functionality).
   - Set the target directory (where the files will be replaced).
   - Set the archive directory (where old files will be moved).

3. **Configure Options**
   - Toggle dark mode if desired.
   - Review operation history in the app.

4. **Excel File Operations**
   - Excel files are automatically validated for format and readability.
   - Support for multiple Excel formats (.xlsx, .xls, .xlsm, .xlsb).
   - View sheet information, dimensions, and data processing.

5. **Run the Replacement**
   - Click the action button to process the replacement and archiving.
   - If multiple files are selected, they will be processed one after another with progress tracking.
   - Word, PDF, and Excel documents will automatically receive a red "UNGÜLTIG" watermark before being archived. Excel files use image-based watermarking with the ungültig_transparent.png image.

5. **Review Logs and Status**
   - Check the output console and status bar for progress and results.

---

## Configuration

The app saves your last-used paths and settings in `file_replacer_config.json` for convenience.

---

## Troubleshooting
- **Outlook Integration**: Requires Microsoft Outlook and the `pywin32` package. If not available, the Outlook features will be disabled.
- **Excel Tracking**: Ensure the Excel file is not open in another program during updates.
- **Permissions**: Run as a user with access to the target and archive directories.

---

## License

This project is provided as-is for professional and internal use. See LICENSE file if available.

---

## Author

Developed by Mustafa Massli and contributors. 