"""
File operations logic including copy, archive, verify, and hash operations.
"""

import os
import shutil
import hashlib
import tempfile
from pathlib import Path
from logic.word_ops import WordOperations
from logic.pdf_ops import PDFOperations
from logic.excel_ops import ExcelOperations


class FileOperations:
    """Handles file operations for the application"""

    def __init__(self, app):
        self.app = app
        self.word_ops = WordOperations(app)
        self.pdf_ops = PDFOperations(app)
        self.excel_ops = ExcelOperations(app)

    def format_file_size(self, size_bytes):
        """Format file size in human readable format"""
        if size_bytes == 0:
            return "0 B"
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.1f} TB"

    def calculate_file_hash(self, filepath):
        """Calculate SHA-256 hash of a file"""
        sha256 = hashlib.sha256()
        with open(filepath, 'rb') as f:
            for block in iter(lambda: f.read(4096), b''):
                sha256.update(block)
        return sha256.hexdigest()

    def verify_file_copy(self, source, destination):
        """Verify that two files are identical by comparing hashes"""
        self.app.log_message(f"Verifying copy from {source} to {destination}")

        if not os.path.exists(destination):
            return False

        if os.path.getsize(source) != os.path.getsize(destination):
            return False

        # Compare file hashes
        source_hash = self.calculate_file_hash(source)
        dest_hash = self.calculate_file_hash(destination)

        if source_hash != dest_hash:
            self.app.log_message(f"Hash mismatch: {source_hash} != {dest_hash}")
            return False

        return True

    def find_matching_files(self, target_dir, attachment_name):
        """Find files in target directory that match the first 10 characters of attachment name"""
        target_dir = Path(target_dir)
        attachment_prefix = attachment_name[:10]
        matching_files = [f for f in target_dir.iterdir() if f.is_file() and f.name[:10] == attachment_prefix]
        return matching_files

    def process_duplicate_files(self, file_paths):
        """
        Process duplicate files with the following logic:
        1. Group files by name (ignoring extensions)
        2. For each group with duplicates:
           - ALL files are processed (replacement and watermarking)
           - Files are grouped for Excel tracking (one entry per base name)
           - PDF priority for hyperlinks in Excel
        3. Return list of files to process (all files) and grouping info
        
        Args:
            file_paths (list): List of file paths to process
            
        Returns:
            tuple: (files_to_process, file_groups) where:
                   - files_to_process: List of all files to process
                   - file_groups: Dict mapping base names to file lists for Excel tracking
        """
        try:
            # Group files by name (ignoring extensions)
            file_groups = {}
            
            self.app.log_message(f"🔍 Processing {len(file_paths)} files for duplicates...")
            
            for file_path in file_paths:
                file_path = Path(file_path)
                if not file_path.exists():
                    self.app.log_message(f"⚠️ File not found: {file_path}")
                    continue
                
                # Get filename without extension
                name_without_ext = file_path.stem
                
                # Debug: Log file source
                temp_dir = tempfile.gettempdir()
                is_outlook_file = str(file_path).startswith(temp_dir)
                self.app.log_message(f"🔍 File: {file_path.name} (Source: {'Outlook' if is_outlook_file else 'Browser'})")
                
                if name_without_ext not in file_groups:
                    file_groups[name_without_ext] = []
                
                file_groups[name_without_ext].append(file_path)
            
            # Process each group - ALL files are processed
            files_to_process = []
            duplicate_summary = []
            
            for name_without_ext, files in file_groups.items():
                if len(files) == 1:
                    # Single file - process normally
                    files_to_process.append(files[0])
                    self.app.log_message(f"📄 Single file: {files[0].name}")
                    
                    # Store single file info consistently
                    file_groups[name_without_ext] = {
                        'files': files,
                        'priority_file': files[0],
                        'has_multiple_formats': False
                    }
                else:
                    # Multiple files with same name - process ALL files
                    duplicate_summary.append(f"'{name_without_ext}' ({len(files)} files)")
                    
                    # Add ALL files to processing list
                    files_to_process.extend(files)
                    
                    # Log all files being processed
                    file_names = [f.name for f in files]
                    self.app.log_message(f"📄 Duplicate group '{name_without_ext}': Processing ALL versions: {', '.join(file_names)}")
                    
                    # Determine priority file for Excel hyperlink
                    pdf_files = [f for f in files if f.suffix.lower() == '.pdf']
                    docx_files = [f for f in files if f.suffix.lower() == '.docx']
                    xlsx_files = [f for f in files if f.suffix.lower() == '.xlsx']
                    
                    self.app.log_message(f"🔍 Priority selection for '{name_without_ext}': PDF={len(pdf_files)}, DOCX={len(docx_files)}, XLSX={len(xlsx_files)}")
                    
                    priority_file = None
                    if pdf_files:
                        priority_file = pdf_files[0]
                        self.app.log_message(f"🔗 Excel hyperlink will use PDF version: {priority_file.name}")
                    elif docx_files:
                        priority_file = docx_files[0]
                        self.app.log_message(f"🔗 Excel hyperlink will use Word version: {priority_file.name}")
                    elif xlsx_files:
                        priority_file = xlsx_files[0]
                        self.app.log_message(f"🔗 Excel hyperlink will use Excel version: {priority_file.name}")
                    else:
                        priority_file = files[0]
                        self.app.log_message(f"🔗 Excel hyperlink will use: {priority_file.name}")
                    
                    # Store priority file info in the group for Excel tracking
                    file_groups[name_without_ext] = {
                        'files': files,
                        'priority_file': priority_file,
                        'has_multiple_formats': True
                    }
                    
                    self.app.log_message(f"🔍 Stored priority file for '{name_without_ext}': {priority_file.name}")
            
            # Log summary
            if duplicate_summary:
                self.app.log_message(f"🔄 Duplicate handling summary:")
                for summary in duplicate_summary:
                    self.app.log_message(f"   - {summary}")
                self.app.log_message(f"✅ Processing {len(files_to_process)} total files from {len(file_paths)} input files")
            else:
                self.app.log_message(f"✅ No duplicates found - processing all {len(files_to_process)} files")
            
            return files_to_process, file_groups
            
        except Exception as e:
            self.app.log_message(f"❌ Error processing duplicate files: {str(e)}")
            # Fallback to original list if error occurs
            return file_paths, {}

    def archive_files(self, files_to_archive, archive_dir):
        """Archive files to the archive directory with watermark for Word, PDF, and Excel documents"""
        archive_dir = Path(archive_dir)
        archive_dir.mkdir(parents=True, exist_ok=True)

        # Add watermarks to Word documents before archiving
        watermarked_word_files = self.word_ops.add_watermark_to_archived_files(files_to_archive, archive_dir)
        
        # Add watermarks to PDF documents before archiving
        watermarked_pdf_files = self.pdf_ops.add_watermark_to_archived_pdfs(files_to_archive, archive_dir)
        
        # Add watermarks to Excel documents before archiving
        watermarked_excel_files = self.excel_ops.add_watermark_to_archived_excel_files(files_to_archive, archive_dir)
        
        # Combine all watermarked files
        all_watermarked_files = watermarked_word_files + watermarked_pdf_files + watermarked_excel_files

        for file_path in all_watermarked_files:
            file_path = Path(file_path)
            backup_path = archive_dir / file_path.name
            shutil.copy2(file_path, backup_path)
            self.app.log_message(f"Archived {file_path} to {backup_path}")
            # Remove the original file after archiving
            file_path.unlink()
            self.app.log_message(f"Removed original file {file_path} after archiving")

    def copy_file(self, source, destination):
        """Copy file from source to destination"""
        shutil.copy2(source, destination)
        self.app.log_message(f"Copied new file to {destination}")

    def cleanup_temp_file(self, file_path):
        """Clean up temporary file if it's from temp directory"""
        file_path_str = str(file_path)
        temp_dir = tempfile.gettempdir()
        if not file_path_str.startswith(temp_dir):
            try:
                os.remove(file_path_str)
                self.app.log_message(f"Attachment file '{file_path_str}' removed from local computer after processing.")
            except Exception as e:
                self.app.log_message(f"⚠️ Could not remove attachment file: {e}")