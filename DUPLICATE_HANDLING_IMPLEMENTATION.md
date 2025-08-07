# Duplicate File Handling Implementation

## Overview

This implementation adds intelligent duplicate file handling to the File Replacer & Archiver application. When multiple files are selected, the system now groups files by name (ignoring extensions) and processes only one file per unique name, with priority given to PDF versions.

## Features

### 1. File Grouping by Name
- Files are grouped by their name without extension
- Example: `contract.pdf`, `contract.docx`, `contract.xlsx` → grouped as "contract"

### 2. PDF Priority
- If multiple versions of a file exist, PDF versions are prioritized
- If no PDF version exists, any other version is used
- Only one file per unique name is processed

### 3. Comprehensive Logging
- Detailed logging of duplicate detection and selection
- Summary of duplicate groups processed
- Clear indication of which files were selected vs. skipped

## Implementation Details

### Core Logic (`logic/file_ops.py`)

The `process_duplicate_files()` method implements the duplicate handling logic:

```python
def process_duplicate_files(self, file_paths):
    """
    Process duplicate files with the following logic:
    1. Group files by name (ignoring extensions)
    2. For each group with duplicates:
       - If PDF exists, prioritize PDF
       - If no PDF, use any other version
    3. Return list of files to process (one per unique name)
    """
```

### Key Features:

1. **File Grouping**: Uses `Path(file_path).stem` to get filename without extension
2. **PDF Detection**: Checks for `.pdf` extension (case-insensitive)
3. **Fallback Logic**: If no PDF exists, uses the first available file
4. **Comprehensive Logging**: Detailed logs for transparency
5. **Error Handling**: Graceful fallback to original list if errors occur

### Integration Points

#### 1. Main Processing (`gui/app.py`)
- Updated `_process_files_thread()` to use duplicate handling
- Modified preview functionality to show duplicate processing
- Updated progress and status messages

#### 2. Preview Functionality
- Preview now shows total files vs. unique files to process
- Includes duplicate processing in preview mode

#### 3. Help Documentation
- Updated help text to explain duplicate handling
- Added clear examples of how duplicates are processed

## Example Scenarios

### Scenario 1: PDF Priority
**Input Files:**
- contract.pdf
- contract.docx
- contract.xlsx
- summary.xlsx
- notes.txt

**Output:**
- contract.pdf (selected - PDF priority)
- summary.xlsx (selected - no duplicates)
- notes.txt (selected - no duplicates)

**Logs:**
```
📄 Duplicate group 'contract': Selected PDF version: contract.pdf
⏭️ Skipped duplicates: contract.docx, contract.xlsx
📄 Single file: summary.xlsx
📄 Single file: notes.txt
```

### Scenario 2: No PDF Available
**Input Files:**
- contract.docx
- contract.xlsx
- summary.docx
- summary.xlsx
- notes.txt

**Output:**
- contract.docx (selected - first available)
- summary.docx (selected - first available)
- notes.txt (selected - no duplicates)

**Logs:**
```
📄 Duplicate group 'contract': No PDF found, using: contract.docx
⏭️ Skipped duplicates: contract.xlsx
📄 Duplicate group 'summary': No PDF found, using: summary.docx
⏭️ Skipped duplicates: summary.xlsx
📄 Single file: notes.txt
```

## User Experience

### 1. Transparent Processing
- Users see exactly which files are being processed
- Clear indication of duplicate detection and selection
- Summary of total vs. unique files

### 2. Consistent Behavior
- Same logic applies to both preview and actual processing
- Works with both file browser and Outlook attachment selection
- Maintains existing functionality for single files

### 3. Error Handling
- Graceful fallback if duplicate processing fails
- Continues with original file list if errors occur
- Detailed error logging for troubleshooting

## Benefits

1. **Efficiency**: Reduces processing time by eliminating duplicate files
2. **Intelligence**: Prioritizes PDF versions when available
3. **Transparency**: Clear logging shows exactly what's happening
4. **Flexibility**: Works with any file type combinations
5. **Backward Compatibility**: No changes to existing single-file workflows

## Testing

The implementation has been thoroughly tested with:
- Multiple file combinations
- PDF priority scenarios
- No-PDF fallback scenarios
- Error handling cases
- Integration with existing functionality

## Future Enhancements

Potential improvements could include:
1. User-configurable priority order (PDF, Word, Excel, etc.)
2. Custom duplicate detection rules
3. Batch processing statistics
4. Advanced duplicate detection (content-based)

## Conclusion

This implementation successfully addresses the requirement for intelligent duplicate file handling while maintaining the existing functionality and user experience. The system now efficiently processes multiple files by eliminating duplicates and prioritizing PDF versions when available. 