# Print Area Configuration Module

This module provides comprehensive print area configuration for Excel worksheets used in invoice generation.

## Features

- ✅ **Dynamic Print Area Detection**: Automatically detects the range of non-empty cells
- ✅ **A4 Paper Size**: Sets paper size to A4 standard
- ✅ **Custom Margins**: Configurable margins (default: 0.1" left/right, 0.75" top/bottom)
- ✅ **Horizontal Centering**: Centers content horizontally on the page
- ✅ **Page Break Preview**: Shows page breaks in worksheet view for easy visualization
- ✅ **Grid Lines & Headers**: Configurable display of grid lines and row/column headers
- ✅ **Print Titles Support**: Can set repeating headers/columns
- ✅ **Custom Print Areas**: Override with specific cell ranges if needed

## Quick Start

### Basic Usage

```python
from openpyxl import load_workbook
from print_area_config import PrintAreaConfig

# Load your invoice workbook
wb = load_workbook('invoice.xlsx')
ws = wb.active

# Configure print settings
config = PrintAreaConfig()
config.configure_print_settings(ws)

# Save the configured invoice
wb.save('invoice_with_print_config.xlsx')
```

### Convenience Function

```python
from print_area_config import configure_print_area

# Quick one-liner configuration
configure_print_area(worksheet)
```

## Configuration Options

The `PrintAreaConfig` class can be customized:

```python
config = PrintAreaConfig()

# Custom margins (in inches)
config.margin_left = 0.2
config.margin_right = 0.2
config.margin_top = 1.0
config.margin_bottom = 1.0

# Centering options
config.center_horizontally = True   # Default: True
config.center_vertically = False    # Default: False

# Apply configuration
config.configure_print_settings(worksheet)
```

### View Options Configuration

```python
config = PrintAreaConfig()

# Configure view display options
config.set_view_options(
    show_page_breaks=True,    # Show page breaks in worksheet view
    show_grid_lines=True,     # Show grid lines
    show_headers=True         # Show row/column headers
)

# Apply configuration
config.configure_print_settings(worksheet)
```

## Advanced Features

### Page Break Preview
The module automatically enables **Page Break Preview** view in Excel, which:
- Shows blue dashed lines where pages will break
- Makes it easy to see how your invoice will print
- Helps optimize layout before printing

### Custom Print Area

```python
config.set_custom_print_area(worksheet, 'A1', 'H50')
```

### Print Titles (Repeating Headers)

```python
# Repeat first 2 rows on each page
config.set_print_titles(worksheet, title_rows='1:2')

# Repeat first column on each page
config.set_print_titles(worksheet, title_cols='A:A')

# Both rows and columns
config.set_print_titles(worksheet, title_rows='1:2', title_cols='A:B')
```

## Integration with Existing Code

### In `style_utils.py`

```python
# Add to your existing style application
from print_area_config import configure_print_area

def apply_all_styles_and_print_config(workbook, worksheet):
    # Your existing style code...
    apply_cell_style(cell, styling_config, context)

    # Add print configuration at the end
    configure_print_area(worksheet)
```

### In Invoice Generation Pipeline

```python
# In your invoice generation workflow
def generate_invoice_with_print_config(data, template_path, output_path):
    # Load template
    wb = load_workbook(template_path)

    # Fill invoice data...
    # (your existing code)

    # Configure print settings for all sheets
    from print_area_config import PrintAreaConfig
    config = PrintAreaConfig()

    for ws in wb.worksheets:
        config.configure_print_settings(ws)

    # Save final invoice
    wb.save(output_path)
```

## Dynamic Print Area Logic

The module automatically detects the print area by:

1. **Scanning all cells** in the worksheet
2. **Finding boundaries** of non-empty, non-whitespace content
3. **Setting print area** from first non-empty row/column to last non-empty row/column
4. **Converting to Excel range format** (e.g., 'A1:H25')

### Example Detection:
- Data in cells A1, B1, C1, A2, B2, C2
- Empty cells in D1, A3, B3, etc.
- **Detected print area**: A1:C2

## Testing

Run the test suite:

```bash
python test_print_area_config.py
```

Run integration example:

```bash
python integration_example.py
```

## Dependencies

- `openpyxl` - For Excel file manipulation
- `logging` - For debug/info messages (optional)

## Error Handling

The module includes comprehensive error handling:

- **Invalid ranges** are caught and logged
- **Missing data** scenarios are handled gracefully
- **File access errors** are properly reported
- **Configuration failures** don't break the entire process

## Logging

Enable logging to see detailed configuration steps:

```python
import logging
logging.basicConfig(level=logging.INFO)

# Now you'll see info messages during configuration
config.configure_print_settings(worksheet)
```

## Testing

Run the test suite:

```bash
python test_print_area_config.py
```

Run integration example:

```bash
python integration_example.py
```

Run page break demo:

```bash
python demo_page_breaks.py
```

## File Structure

```
invoice_gen/
├── print_area_config.py          # Main module
├── test_print_area_config.py     # Unit tests
├── integration_example.py        # Integration examples
├── demo_page_breaks.py           # Page break demonstration
└── README.md                     # This documentation
```

## Compatibility

- **Python**: 3.7+
- **openpyxl**: 3.0+
- **Excel formats**: .xlsx, .xlsm

## Troubleshooting

### Print Area Not Set Correctly
- Check that your data doesn't have leading/trailing whitespace
- Verify cells contain actual values (not just formulas)
- Use logging to see detected boundaries

### Margins Not Applied
- Ensure worksheet has page setup initialized
- Check that margin values are reasonable (0.1-2.0 inches)

### Paper Size Issues
- Verify A4 paper size is supported by your printer
- Some Excel versions may override paper size settings
