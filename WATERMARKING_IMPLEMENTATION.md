# Diagonal Red Text Field Watermarking Implementation

## Overview

This implementation adds diagonal red text field watermarking to Excel files using Excel automation (win32com). The watermark appears as a diagonal red text field with the text "UNGÜLTIG" that is rotated and positioned as a watermark overlay. This creates **actual floating text boxes** exactly like manual insertion in Excel. **All sheets in the workbook are watermarked automatically.**

## Features

- ✅ **Diagonal Text Field**: Creates a text box shape rotated at -45 degrees
- ✅ **Red Color**: Text is displayed in red color
- ✅ **Transparent Background**: Shape background is transparent for watermark effect
- ✅ **No Border**: Shape border is hidden for clean appearance
- ✅ **Configurable**: Text, font size, color, and transparency can be customized
- ✅ **Excel Automation**: Uses win32com to create actual floating text boxes
- ✅ **Multi-Sheet Support**: Automatically watermarks ALL sheets in the workbook
- ✅ **Fallback Support**: Falls back to Spire.XLS or openpyxl if win32com is not available

## Implementation Details

### Main Method: `add_watermark_to_excel()`

```python
def add_watermark_to_excel(self, file_path, watermark_text="UNGÜLTIG", font_size=36, font_color="FF0000", transparency=0.7):
```

**Parameters:**
- `file_path`: Path to the Excel file to watermark
- `watermark_text`: Text to display (default: "UNGÜLTIG")
- `font_size`: Font size in points (default: 36)
- `font_color`: Color in hex format (default: "FF0000" for red)
- `transparency`: Transparency level 0.0-1.0 (default: 0.7)

### Win32com Implementation: `_add_diagonal_text_field_with_win32com()`

This method creates actual floating text boxes using Excel automation on ALL sheets:

1. **Open Excel**: Launches Excel application in background
2. **Load Workbook**: Opens the existing Excel file
3. **Process Each Sheet**: Iterates through all worksheets in the workbook
4. **Create Text Box**: Adds a real floating text box shape to each sheet
5. **Set Text**: Configures the text content ("UNGÜLTIG")
6. **Format Text**: Sets font properties (Arial, size, red color, bold)
7. **Rotate Shape**: Rotates the entire shape by -45 degrees
8. **Make Transparent**: Sets background to transparent and hides border
9. **Save File**: Saves the modified Excel file with all sheets watermarked

### Key Code Features

```python
# Create Excel application object
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False  # Run in background

# Process each worksheet in the workbook
sheet_count = workbook.Worksheets.Count
watermarked_sheets = 0

for sheet_index in range(1, sheet_count + 1):
    worksheet = workbook.Worksheets(sheet_index)
    
    # Add a text box shape
    textbox = worksheet.Shapes.AddTextbox(
        Orientation=1,  # Horizontal text
        Left=center_col * 50,  # Position from left
        Top=center_row * 20,   # Position from top
        Width=300,             # Width of text box
        Height=100             # Height of text box
    )

# Set the text content
textbox.TextFrame.Characters().Text = watermark_text

# Format the text
textbox.TextFrame.Characters().Font.Name = "Arial"
textbox.TextFrame.Characters().Font.Size = font_size
textbox.TextFrame.Characters().Font.Color = 255  # Red color
textbox.TextFrame.Characters().Font.Bold = True

# Rotate the text box diagonally (-45 degrees)
textbox.Rotation = -45

# Make background transparent
textbox.Fill.Visible = False

# Hide the border
textbox.Line.Visible = False

watermarked_sheets += 1
```

## Dependencies

### Required Packages

Add to `requirements.txt`:
```
pywin32>=306
spire.xls>=11.0.0
```

### Import Handling

The code includes graceful handling for multiple methods:

```python
# Try to import win32com for Excel automation
try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False

# Try to import Spire.XLS for advanced watermarking
try:
    from spire.xls import *
    from spire.xls.common import *
    SPIRE_XLS_AVAILABLE = True
except ImportError:
    SPIRE_XLS_AVAILABLE = False
```

## Usage Examples

### Basic Usage

```python
from logic.excel_ops import ExcelOperations

# Create ExcelOperations instance
excel_ops = ExcelOperations(app)

# Add watermark to Excel file
success = excel_ops.add_watermark_to_excel("path/to/file.xlsx")
```

### Custom Configuration

```python
# Custom watermark with different text and settings
success = excel_ops.add_watermark_to_excel(
    file_path="path/to/file.xlsx",
    watermark_text="VOID",
    font_size=48,
    font_color="FF0000",
    transparency=0.8
)
```

## Testing

Run the test scripts to verify functionality:

```bash
# Test multi-sheet watermarking (creates floating text boxes on all sheets)
python test_multi_sheet_watermark.py

# Test win32com-based watermarking (creates actual floating text boxes)
python test_win32com_watermark.py

# Test general watermarking functionality
python test_watermark.py
```

This will:
1. Check if win32com and Spire.XLS are available
2. Create test Excel files (single and multi-sheet)
3. Apply the watermark using the best available method
4. Show the results for each sheet

## Fallback Behavior

The system uses a priority-based approach for watermarking:

1. **Primary**: win32com Excel automation (creates actual floating text boxes)
2. **Secondary**: Spire.XLS diagonal text field watermarking
3. **Fallback**: openpyxl text-based watermarking in cells

## Error Handling

The implementation includes comprehensive error handling:

- ✅ Import errors for Spire.XLS
- ✅ File loading/saving errors
- ✅ Shape creation and formatting errors
- ✅ Graceful fallback to alternative methods

## Visual Result

The watermark appears as:
- **Text**: "UNGÜLTIG" (or custom text)
- **Color**: Red
- **Orientation**: Diagonal (-45 degrees)
- **Style**: Bold Arial font
- **Background**: Transparent
- **Border**: Hidden
- **Position**: Centered in the worksheet

## Integration

This watermarking functionality integrates with the existing Excel operations:

- ✅ Works with `add_watermark_to_archived_excel_files()`
- ✅ Compatible with existing Excel tracking functionality
- ✅ Maintains file integrity and formatting
- ✅ Preserves existing data and formulas

## Performance Considerations

- **Spire.XLS**: Fast and efficient for shape-based watermarking
- **Memory Usage**: Minimal additional memory usage
- **File Size**: Slight increase due to shape objects
- **Processing Time**: Quick processing for typical Excel files

## Troubleshooting

### Common Issues

1. **Spire.XLS Not Available**
   - Install: `pip install spire.xls>=11.0.0`
   - Check import: `python -c "from spire.xls import *"`

2. **File Permission Errors**
   - Ensure write permissions on target directory
   - Close Excel file before processing

3. **Shape Creation Errors**
   - Verify Excel file is not corrupted
   - Check file format compatibility

### Debug Information

Enable detailed logging to troubleshoot issues:

```python
# Check Spire.XLS availability
print(f"Spire.XLS Available: {SPIRE_XLS_AVAILABLE}")

# Test with verbose logging
excel_ops.add_watermark_to_excel(file_path, verbose=True)
```

## Future Enhancements

Potential improvements:
- Multiple watermark positions
- Custom rotation angles
- Watermark opacity controls
- Batch processing optimizations
- Additional shape types
- Watermark templates 