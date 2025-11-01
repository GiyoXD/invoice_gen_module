# Builder Architecture: `template_state_builder.py`

This document explains the structure and purpose of the `TemplateStateBuilder` class, which is responsible for capturing the complete state (content, formatting, merges, dimensions) of a template worksheet and restoring it to output worksheets.

## Overview

The `TemplateStateBuilder` acts as a "memory" system for template worksheets. It captures all styling, content, merges, row heights, and column widths from a template during initialization, then provides methods to selectively restore this state to target worksheets. This enables the separation of template (read-only) and output (writable) worksheets while preserving visual consistency.

- **Purpose**: To capture and restore complete worksheet state including content, formatting, merges, and dimensions.
- **Pattern**: Builder pattern - constructs worksheet state snapshots and reconstructs them elsewhere.
- **Key Responsibility**: Preserves template appearance and structure by capturing state once and restoring it as needed.

## Key Concepts

### Immediate Capture Strategy
The builder captures template state **immediately during initialization**, not lazily. This ensures the captured state represents the original, unmodified template before any processing occurs.

### Separate Template and Output
The builder enables a clean separation:
- **Template Worksheet**: Source of truth (read-only usage)
- **Output Worksheet**: Target for restoration (writable)

This separation completely avoids merge conflicts and template corruption.

## `TemplateStateBuilder` Class

### `__init__(...)` - The Constructor and Capture Phase

The constructor **immediately captures** the complete template state during initialization.

- **Purpose**: To initialize the builder and capture all template state from the source worksheet.
- **Parameters**:
    - `worksheet: Worksheet`: The template worksheet to capture state from (read-only usage).
    - `num_header_cols: int`: Number of columns in the header (determines capture width).
    - `header_end_row: int`: Last row of the header section in the template.
    - `footer_start_row: int`: First row of the footer section in the template.
    - `debug: bool`: Enable debug printing for troubleshooting (default: `False`).

- **Initialization Process**:
    1. **Store Parameters**: Save worksheet reference and boundary information.
    2. **Initialize State Storage**:
        - `header_state`: List of lists containing cell dictionaries for header rows.
        - `footer_state`: List of lists containing cell dictionaries for footer rows.
        - `header_merged_cells`: List of merge range strings for header.
        - `footer_merged_cells`: List of merge range strings for footer.
        - `row_heights`: Dictionary mapping row numbers to heights.
        - `column_widths`: Dictionary mapping column numbers to widths.
    3. **Create Default Style References**: Create a dummy workbook to get default style objects for comparison.
    4. **Calculate Max Column**: Scan entire worksheet to find the rightmost column with content or styling.
    5. **Calculate Max Row**: Scan entire worksheet to find the bottommost row with content or styling.
    6. **Capture Header State**: Call `_capture_header()` to capture all header rows.
    7. **Capture Footer State**: Call `_capture_footer()` to capture all footer rows.

- **Captured Information Per Cell**:
    - `value`: Cell value (text, number, formula)
    - `font`: Font object (copied) if non-default
    - `fill`: PatternFill object (copied) if non-default
    - `border`: Border object (copied) if non-default
    - `alignment`: Alignment object (copied) if non-default
    - `number_format`: Number format string

- **Debug Mode**: When `debug=True`, prints detailed information about capture and restoration operations.

### `_has_content_or_style(self, cell) -> bool` - Content Detection Helper

Determines if a cell has meaningful content or non-default styling.

- **Purpose**: To distinguish cells with actual content/styling from empty default cells.
- **Checks**:
    1. Cell has non-empty value
    2. Cell has non-default font
    3. Cell has non-default fill
    4. Cell has non-default border
    5. Cell has non-default alignment
- **Return Value**: `True` if cell has content or styling, `False` otherwise.
- **Use Case**: Used to calculate the actual bounds of the template (avoiding empty trailing rows/columns).

### `_is_default_style(self, style_obj, default_obj) -> bool` - Style Comparison Helper

Compares a style object to the default style to determine if it's been customized.

- **Purpose**: To identify whether a style has been modified from Excel's defaults.
- **Supported Style Types**:
    - **Font**: Compares name, size, bold, italic, underline, strike, color
    - **PatternFill**: Compares fill_type, start_color, end_color
    - **Border**: Compares left, right, top, bottom, diagonal sides
    - **Alignment**: Compares horizontal, vertical, rotation, wrap, shrink, indent
- **Return Value**: `True` if style matches default, `False` if customized.

### `_get_cell_info(self, worksheet, row, col) -> Dict[str, Any]` - Cell State Extractor

Extracts complete state information from a single cell.

- **Purpose**: To capture all styling and content from a cell for later restoration.
- **Merge Handling**: If cell is part of a merged range, retrieves styling from the top-left cell of the merge (openpyxl convention).
- **Return Value**: Dictionary containing `value`, `font`, `fill`, `border`, `alignment`, `number_format`.
- **Optimization**: Only copies style objects if they differ from defaults (saves memory and processing time).

### `_capture_header(self, end_row: int)` - Header State Capture

Captures all header rows including content, formatting, merges, and dimensions.

- **Purpose**: To capture the complete state of the header section for later restoration.
- **Process**:
    1. **Find Header Start**: Scan from row 1 to find first row with content.
    2. **Capture Cell Data**: For each row from start to `end_row`:
        - Call `_get_cell_info()` for every column
        - Store row data in `header_state`
        - Record row height in `row_heights`
    3. **Capture Header Merges**: Store all merge ranges that fall within header rows.
    4. **Capture Column Widths**: Record widths for all columns.
- **Debug Output**: Prints number of rows and merges captured when debug mode enabled.

### `_capture_footer(self, footer_start_row: int, max_possible_footer_row: int)` - Footer State Capture

Captures all footer rows with intelligent end-of-footer detection.

- **Purpose**: To capture the complete state of the footer section, automatically detecting where the footer ends.
- **Smart End Detection**:
    - Scans from `footer_start_row` up to 50 rows or `max_possible_footer_row`
    - Looks for rows with **actual values** or **merge participation**
    - Stops after finding 3 consecutive empty rows (configurable via `MAX_EMPTY_ROWS_BEFORE_STOP`)
    - This prevents capturing hundreds of styled-but-empty rows
- **Process**:
    1. **Detect Footer End**: Find the last row with meaningful content.
    2. **Capture Cell Data**: For each footer row:
        - Call `_get_cell_info()` for every column
        - Store row data in `footer_state`
        - Record row height in `row_heights`
    3. **Capture Footer Merges**: Store all merge ranges within footer bounds.
    4. **Capture Column Widths**: Record widths for all columns.
- **Debug Output**: Prints footer boundaries, row count, merges, and sample cell data.

### `restore_header_only(self, target_worksheet: Worksheet)` - Header-Only Restoration

Restores only the header section to a target worksheet.

- **Purpose**: To copy the template header to a fresh output worksheet without touching footer area.
- **Process**:
    1. **Restore Cell Values and Formatting**: For each captured header cell:
        - Write value to target cell
        - Apply font, fill, border, alignment (deep copied)
        - Set number format
    2. **Restore Header Merges**: Apply all captured merge ranges.
    3. **Restore Row Heights**: Set heights for header rows.
    4. **Restore Column Widths**: Set widths for all columns.
- **Use Case**: Called by `LayoutBuilder` to place template header before building data and footer.

### `restore_footer_only(self, target_worksheet: Worksheet, footer_start_row: int)` - Footer-Only Restoration

Restores only the footer section to a target worksheet at a specified position.

- **Purpose**: To place the template footer below dynamic content at a calculated position.
- **Parameters**:
    - `target_worksheet`: The worksheet to restore footer to
    - `footer_start_row`: The row where the template footer should start (typically after data footer)
- **Offset Calculation**:
    ```
    offset = footer_start_row - template_footer_start_row
    ```
    This adjusts all footer row numbers to the new position.
- **Process**:
    1. **Calculate Offset**: Determine how many rows to shift the footer.
    2. **Restore Cell Values and Formatting**: For each captured footer cell:
        - Write value to target cell at `original_row + offset`
        - Apply formatting (deep copied)
    3. **Restore Footer Merges with Offset**: Adjust merge range row numbers and apply.
    4. **Restore Row Heights with Offset**: Set heights for footer rows at new positions.
- **Use Case**: Called by `LayoutBuilder` to place template footer (signatures, static text) below dynamic data footer.

### `restore_state(self, target_worksheet: Worksheet, data_start_row: int, data_table_end_row: int, restore_footer_merges: bool = True)` - Complete State Restoration

Restores complete template state (formatting only, not values) to preserve structure.

- **Purpose**: To restore merges, heights, and widths without overwriting cell values that have already been written.
- **Parameters**:
    - `target_worksheet`: The worksheet to restore state to
    - `data_start_row`: Starting row of data (for offset calculation)
    - `data_table_end_row`: Ending row of data table (for offset calculation)
    - `restore_footer_merges`: Whether to restore footer merges (default: `True`). Set to `False` when `FooterBuilder` creates its own merges.
- **Process**:
    1. **Restore Header Merges**: Apply header merge ranges without offset.
    2. **Calculate Footer Offset**: `offset = (data_table_end_row + 1) - template_footer_start_row`
    3. **Restore Footer Merges (Optional)**: Apply footer merges with offset if requested.
    4. **Restore Row Heights**: Apply heights for header and footer rows (with offset for footer).
    5. **Restore Column Widths**: Apply widths for all columns.
- **Use Case**: Alternative restoration method for scenarios where cell values are written separately from formatting.

## State Storage Structure

### Header State Example
```python
header_state = [
    [  # Row 1
        {'value': 'Invoice No.', 'font': Font(...), 'border': Border(...), ...},
        {'value': 'Date', 'font': Font(...), 'border': Border(...), ...},
        # ... more cells
    ],
    [  # Row 2
        {'value': 'Description', 'font': Font(...), 'border': Border(...), ...},
        {'value': 'Quantity', 'font': Font(...), 'border': Border(...), ...},
        # ... more cells
    ]
]

header_merged_cells = ['A1:B1', 'C1:D1']
```

### Footer State Example
```python
footer_state = [
    [  # Footer row 1
        {'value': 'Manufacturer:', 'font': Font(...), 'alignment': Alignment(...), ...},
        {'value': 'ABC Company', 'font': Font(...), ...},
        # ... more cells
    ],
    # ... more footer rows
]

footer_merged_cells = ['A15:C15', 'D15:F15']
```

## Data Flow

```
Template Worksheet
        ↓
TemplateStateBuilder.__init__()
        ↓
    ┌───────┴───────┐
    ↓               ↓
_capture_header  _capture_footer
    ↓               ↓
Scan cells    Detect end
    ↓               ↓
Store state   Store state
    ↓               ↓
Store merges  Store merges
        ↓
State Captured (header_state, footer_state, merges, dimensions)
        ↓
    ┌───────┴────────────┐
    ↓                    ↓
restore_header_only  restore_footer_only
    ↓                    ↓
Target Worksheet ← Apply with offset
```

## Key Design Decisions

### 1. **Immediate Capture**
State is captured during `__init__`, not lazily. This ensures the snapshot represents the original template before any modifications occur.

### 2. **Deep Copy of Styles**
All style objects (Font, Fill, Border, Alignment) are deep-copied using `copy.copy()` to prevent reference issues when applying to multiple cells.

### 3. **Default Style Filtering**
Only non-default styles are captured, significantly reducing memory usage and restoration time for large templates.

### 4. **Intelligent Footer End Detection**
Instead of capturing all 180 rows of a typical template, the builder detects the actual footer end by looking for consecutive empty rows, capturing only relevant content.

### 5. **Merge-Aware Cell Info Extraction**
When extracting cell info from merged ranges, the builder retrieves styling from the top-left cell (openpyxl's convention), ensuring correct style application.

### 6. **Offset-Based Footer Restoration**
Footer restoration uses offset calculation to place footer content at any row position, enabling flexible document layouts.

### 7. **Selective Merge Restoration**
The `restore_footer_merges` parameter allows skipping footer merge restoration when `FooterBuilder` creates its own merges, avoiding conflicts.

## Footer Offset Calculation

The offset mechanism allows footer placement at any position:

```python
# Template has footer at row 20
template_footer_start_row = 20

# Data ends at row 50, so footer should start at row 51
new_footer_start_row = 51

# Calculate offset
offset = 51 - 20 = 31

# Apply offset to all footer elements:
# - Template footer row 20 → Output row 51
# - Template footer row 21 → Output row 52
# - Template merge A20:C20 → Output merge A51:C51
```

## Smart Footer End Detection Algorithm

```python
MAX_EMPTY_ROWS_BEFORE_STOP = 3
consecutive_empty_rows = 0

for row in range(footer_start_row, footer_start_row + 50):
    if row has value or row in merge:
        footer_end_row = row
        consecutive_empty_rows = 0
    else:
        consecutive_empty_rows += 1
        if consecutive_empty_rows >= 3:
            break  # Footer ends here
```

This algorithm:
- Scans up to 50 rows
- Tracks consecutive empty rows
- Stops after 3 consecutive empties
- Result: Captures only relevant footer content

## Dependencies

- **openpyxl**: Core Excel manipulation library
    - `Worksheet` - Worksheet objects
    - `Font`, `PatternFill`, `Border`, `Alignment` - Style objects
    - `get_column_letter` - Column index to letter conversion
    - `range_boundaries` - Parse merge range strings
- **copy**: Python standard library for deep copying style objects
- **typing**: Type hints for better code clarity

## Usage Example

### Basic Usage - Capture and Restore

```python
from invoice_generator.builders.template_state_builder import TemplateStateBuilder

# Load template workbook
template_workbook = openpyxl.load_workbook('template.xlsx')
template_worksheet = template_workbook['Invoice']

# Create state builder (captures immediately)
state_builder = TemplateStateBuilder(
    worksheet=template_worksheet,
    num_header_cols=6,
    header_end_row=2,      # Header is rows 1-2
    footer_start_row=20,   # Footer starts at row 20
    debug=False
)

# Create output worksheet
output_workbook = openpyxl.Workbook()
output_worksheet = output_workbook.active

# Restore header to output
state_builder.restore_header_only(output_worksheet)

# ... build data table (rows 3-50) ...

# Restore footer after data
state_builder.restore_footer_only(output_worksheet, footer_start_row=51)

# Save output
output_workbook.save('output.xlsx')
```

### Debug Mode for Troubleshooting

```python
# Enable debug mode for detailed output
state_builder = TemplateStateBuilder(
    worksheet=template_worksheet,
    num_header_cols=6,
    header_end_row=2,
    footer_start_row=20,
    debug=True  # Enable debug printing
)

# Output will include:
# [TemplateStateBuilder] Capturing template state during init
#   Header: rows 1-2, Footer: rows 20-180
# [TemplateStateBuilder] Header captured: 2 rows, 3 merges
# [TemplateStateBuilder] Footer ends at row 25 (found 6 footer rows)
# [TemplateStateBuilder] Footer captured: 6 rows, 4 merges, start row: 20
#   First footer row sample:
#     Cell 1: value=Manufacturer:, font=True, fill=False, border=True
# ...
```

### Integration with LayoutBuilder

```python
# Inside LayoutBuilder.build()

# Step 1: Create state builder and capture template
template_state_builder = TemplateStateBuilder(
    worksheet=template_worksheet,
    num_header_cols=num_header_cols,
    header_end_row=header_end_row,
    footer_start_row=footer_start_row
)

# Step 2: Restore header to output
template_state_builder.restore_header_only(target_worksheet=output_worksheet)

# Step 3: Build header content (HeaderBuilder)
# Step 4: Build data table (DataTableBuilder)
# Step 5: Build data footer (FooterBuilder)

# Step 6: Restore template footer after data footer
next_row = 52  # After data footer at row 51
template_state_builder.restore_footer_only(
    target_worksheet=output_worksheet,
    footer_start_row=next_row
)
```

## Notes

- The builder assumes the template worksheet is well-formed with a clear header-data-footer structure.
- Capture happens immediately in `__init__`, so the template must be in its original state when the builder is created.
- All style objects are deep-copied to prevent reference issues when applying to multiple cells.
- The builder only captures non-default styles, significantly improving performance for large templates.
- Footer end detection is intelligent but has a 50-row search limit to prevent excessive scanning.
- The offset mechanism allows footer placement at any position, but the caller must ensure adequate space.
- Debug mode is invaluable for troubleshooting state capture and restoration issues.
- Column widths and row heights are preserved exactly as in the template.
- Merged cells are stored as range strings (e.g., `"A1:B1"`) and restored with bounds adjustment for footer offset.
- The `restore_state()` method is an alternative approach that restores only formatting (merges, dimensions) without overwriting values.











