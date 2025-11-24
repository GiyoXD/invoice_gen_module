# Builder Architecture: `header_builder.py`

This document explains the structure and purpose of the `HeaderBuilder` class, which is responsible for constructing table headers in invoice worksheets, including multi-row headers with cell merging and styling.

## Overview

The `HeaderBuilder` is a foundational component that constructs the header section of data tables. It processes a layout configuration to create headers that can span multiple rows and columns, applies appropriate styling, and generates column mapping information used by other builders.

- **Purpose**: To construct and style table headers while creating column reference mappings for subsequent data operations.
- **Pattern**: Builder pattern - constructs header structures from configuration specifications.
- **Key Responsibility**: Translates header layout configuration into formatted Excel header rows with proper merging, styling, and metadata.

## `HeaderBuilder` Class

### `__init__(...)` - The Constructor

The constructor initializes the builder with the worksheet, position, layout configuration, and styling specifications.

- **Purpose**: To configure the header builder with all information needed to construct the header section.
- **Parameters**:
    - `worksheet: Worksheet`: The `openpyxl` Worksheet object where the header will be written.
    - `start_row: int`: The Excel row index where the header should begin (typically row 1 or 2).
    - `header_layout_config: List[Dict[str, Any]]`: List of cell configuration dictionaries defining the header structure. Each dictionary contains:
        - `text`: The display text for the header cell (e.g., `"Description"`, `"Quantity"`).
        - `row`: Row offset from `start_row` (0-based, e.g., `0` for first row, `1` for second row).
        - `col`: Column offset (0-based, e.g., `0` for column A, `1` for column B).
        - `id`: Column identifier string (e.g., `"col_desc"`, `"col_qty"`) used for referencing this column elsewhere.
        - `rowspan`: Number of rows this cell should span (default: `1`).
        - `colspan`: Number of columns this cell should span (default: `1`).
    - `sheet_styling_config: Optional[StylingConfigModel]`: The styling configuration model containing header-specific styling rules (fonts, alignments, borders, row heights).

### `build(self) -> Optional[Dict[str, Any]]` - The Main Build Method

This is the primary method that constructs the complete header section and returns metadata about the header structure.

- **Purpose**: To build the header rows from configuration, apply styling and merging, then return information about the header structure for use by other builders.
- **Process Flow**:
    1. **Validation**: Check if `header_layout_config` is provided and `start_row` is valid.
    2. **Dimension Calculation**: Call `calculate_header_dimensions()` to determine total rows and columns needed.
    3. **Unmerge Preparation**: Unmerge the entire header block to ensure clean slate for new header.
    4. **Initialization**: Set up tracking variables for row indices, column count, and column mappings.
    5. **Cell Processing Loop**: For each cell configuration:
        - Calculate absolute cell position (row and column).
        - Update tracking variables (last row index, max column).
        - Write cell value to worksheet.
        - Apply header styling using `apply_header_style()`.
        - Apply context-aware cell styling using `apply_cell_style()` with `is_header: True` flag.
        - Build column mappings if cell has an `id`.
        - Apply cell merging if `rowspan > 1` or `colspan > 1`.
    6. **Return Metadata**: Return dictionary containing header structure information.

- **Return Value**: `Optional[Dict[str, Any]]`
    - Returns `None` if header config is invalid or empty.
    - On success, returns a dictionary with the following keys:
        - `first_row_index` (`int`): The Excel row index of the first header row.
        - `second_row_index` (`int`): The Excel row index of the last header row (data starts after this).
        - `column_map` (`Dict[str, str]`): Maps column text to column letters (e.g., `{"Description": "A", "Quantity": "B"}`).
        - `column_id_map` (`Dict[str, int]`): Maps column IDs to 1-based column indices (e.g., `{"col_desc": 1, "col_qty": 2}`).
        - `num_columns` (`int`): Total number of columns in the header.

- **Metadata Usage**: The returned dictionary is critical for downstream operations:
    - `DataTableBuilder` uses `column_id_map` to write data to correct columns.
    - `FooterBuilder` uses `column_id_map` to place totals and formulas.
    - `second_row_index` determines where data writing begins.
    - `num_columns` defines the table width for styling and merging operations.

## Header Layout Configuration

The header layout configuration is a list of cell configuration dictionaries that define the complete header structure.

### Single-Row Header Example

```python
header_layout_config = [
    {"text": "No", "row": 0, "col": 0, "id": "col_no", "colspan": 1},
    {"text": "Description", "row": 0, "col": 1, "id": "col_desc", "colspan": 1},
    {"text": "Quantity", "row": 0, "col": 2, "id": "col_qty", "colspan": 1},
    {"text": "Unit Price", "row": 0, "col": 3, "id": "col_price", "colspan": 1},
    {"text": "Total", "row": 0, "col": 4, "id": "col_total", "colspan": 1}
]
```

This creates a single-row header:
```
| No | Description | Quantity | Unit Price | Total |
```

### Multi-Row Header Example

```python
header_layout_config = [
    # First row - spanning headers
    {"text": "Product Information", "row": 0, "col": 0, "id": None, "colspan": 2, "rowspan": 1},
    {"text": "Pricing", "row": 0, "col": 2, "id": None, "colspan": 2, "rowspan": 1},
    {"text": "Pallet\nNo", "row": 0, "col": 4, "id": "col_pallet", "rowspan": 2},  # Spans 2 rows
    
    # Second row - detail headers
    {"text": "No", "row": 1, "col": 0, "id": "col_no", "colspan": 1},
    {"text": "Description", "row": 1, "col": 1, "id": "col_desc", "colspan": 1},
    {"text": "Quantity", "row": 1, "col": 2, "id": "col_qty", "colspan": 1},
    {"text": "Total", "row": 1, "col": 3, "id": "col_total", "colspan": 1}
]
```

This creates a two-row header with merging:
```
|  Product Information  |      Pricing      | Pallet |
|    No   | Description | Quantity | Total  |   No   |
```

## Data Flow

```
Header Layout Config (List of Dicts)
        ↓
calculate_header_dimensions() → num_rows, num_cols
        ↓
Unmerge Header Block
        ↓
Initialize Tracking Variables
        ↓
For Each Cell Config:
    ├─ Calculate Position
    ├─ Write Cell Value
    ├─ Apply Header Style
    ├─ Apply Context-Aware Style
    ├─ Update Column Mappings
    └─ Apply Cell Merging (if needed)
        ↓
Return Header Info Dictionary
        ↓
Used by DataTableBuilder, FooterBuilder, etc.
```

## Key Design Decisions

### 1. **Offset-Based Positioning**
Cell positions use offsets (`row` and `col`) relative to `start_row`, allowing headers to be placed at any vertical position in the worksheet without hardcoding absolute positions.

### 2. **Dual Column Mapping**
The builder creates two types of column mappings:
- **Text-to-Letter Map**: For human-readable references and debugging.
- **ID-to-Index Map**: For programmatic column referencing in builders and processors.

This dual mapping provides both readability and flexibility.

### 3. **Pre-Unmerge Strategy**
The entire header block is unmerged before building, ensuring that re-running the builder produces consistent results without merge conflicts from previous runs.

### 4. **Layered Styling**
The builder applies two levels of styling:
- **General Header Styling**: Applied to all header cells via `apply_header_style()`.
- **Context-Aware Styling**: Applied per-cell via `apply_cell_style()` with `is_header: True` flag.

This allows both uniform header appearance and column-specific styling overrides.

### 5. **Optional Column IDs**
Not all header cells require an `id`. Cells used purely for visual organization (like spanning labels) can omit the `id` field, keeping the column mapping clean and focused on data columns.

## Column ID Conventions

Column IDs follow a consistent naming pattern for clarity:

| Column ID | Purpose | Example Header Text |
|-----------|---------|---------------------|
| `col_no` | Item number/sequence | "No", "Item No" |
| `col_desc` | Product description | "Description", "Product" |
| `col_qty` | Quantity | "Quantity", "Qty" |
| `col_price` | Unit price | "Unit Price", "Price" |
| `col_total` | Total amount | "Total", "Amount" |
| `col_weight` | Weight information | "Net Weight", "G.W." |
| `col_pallet` | Pallet information | "Pallet No", "Pallet Info" |
| `col_hs` | HS code | "HS Code", "H.S. Code" |

These IDs are used throughout the codebase for consistent column referencing.

## Dependencies

- **Layout Utilities**: `invoice_generator.utils.layout` - Provides dimension calculation and unmerging functions.
- **Style Application**: `invoice_generator.styling.style_applier` - Applies header and cell styling.
- **Style Configuration**: `invoice_generator.styling.models.StylingConfigModel` - Type-safe styling configuration model.
- **openpyxl**: Core Excel manipulation library for cells, merging, and column letter conversion.

## Usage Example

### Simple Single-Row Header

```python
from invoice_generator.builders.header_builder import HeaderBuilder

# Define header layout
header_config = [
    {"text": "No", "row": 0, "col": 0, "id": "col_no"},
    {"text": "Description of Goods", "row": 0, "col": 1, "id": "col_desc"},
    {"text": "Quantity", "row": 0, "col": 2, "id": "col_qty"},
    {"text": "Unit Price", "row": 0, "col": 3, "id": "col_price"},
    {"text": "Amount", "row": 0, "col": 4, "id": "col_total"}
]

# Create and build header
builder = HeaderBuilder(
    worksheet=worksheet,
    start_row=2,  # Header starts at row 2
    header_layout_config=header_config,
    sheet_styling_config=styling_config
)

header_info = builder.build()

if header_info:
    print(f"Header built from row {header_info['first_row_index']} to {header_info['second_row_index']}")
    print(f"Total columns: {header_info['num_columns']}")
    print(f"Column ID map: {header_info['column_id_map']}")
    # Output:
    # Header built from row 2 to 2
    # Total columns: 5
    # Column ID map: {'col_no': 1, 'col_desc': 2, 'col_qty': 3, 'col_price': 4, 'col_total': 5}
```

### Complex Multi-Row Header with Merging

```python
# Define multi-row header with spans
header_config = [
    # Row 0 - Category headers
    {"text": "Item", "row": 0, "col": 0, "id": None, "rowspan": 2},
    {"text": "Product Details", "row": 0, "col": 1, "id": None, "colspan": 3},
    {"text": "Shipping", "row": 0, "col": 4, "id": None, "colspan": 2},
    
    # Row 1 - Detail headers (under "Product Details")
    {"text": "Description", "row": 1, "col": 1, "id": "col_desc"},
    {"text": "Quantity", "row": 1, "col": 2, "id": "col_qty"},
    {"text": "Price", "row": 1, "col": 3, "id": "col_price"},
    
    # Row 1 - Detail headers (under "Shipping")
    {"text": "Weight", "row": 1, "col": 4, "id": "col_weight"},
    {"text": "Pallet", "row": 1, "col": 5, "id": "col_pallet"}
]

builder = HeaderBuilder(
    worksheet=worksheet,
    start_row=1,
    header_layout_config=header_config,
    sheet_styling_config=styling_config
)

header_info = builder.build()

# Result:
# Row 1: | Item |      Product Details      |    Shipping    |
# Row 2: |      | Description | Qty | Price | Weight | Pallet |
```

## Integration with Other Builders

The `header_info` dictionary returned by `build()` is used extensively by other builders:

### DataTableBuilder Integration
```python
# HeaderBuilder creates the header and returns info
header_info = header_builder.build()

# DataTableBuilder uses the header info
data_builder = DataTableBuilder(
    worksheet=worksheet,
    header_info=header_info,  # Passed here
    mapping_rules=mapping_rules,
    data_source=data
)

# DataTableBuilder uses:
# - header_info['second_row_index'] to know where to start writing data
# - header_info['column_id_map'] to write data to correct columns
# - header_info['num_columns'] for row operations
```

### FooterBuilder Integration
```python
# FooterBuilder uses header info for column references
footer_builder = FooterBuilder(
    worksheet=worksheet,
    footer_row_num=footer_row,
    header_info=header_info,  # Passed here
    footer_config=footer_config,
    sum_ranges=sum_ranges,
    pallet_count=total_pallets
)

# FooterBuilder uses:
# - header_info['column_id_map'] to place totals and formulas
# - header_info['num_columns'] for merge operations
```

## Notes

- The builder assumes the worksheet already exists; it does not create new worksheets.
- Header rows are not automatically sized; row heights are applied separately via styling configuration.
- Column widths are not set by `HeaderBuilder`; they are managed by `LayoutBuilder` using `apply_column_widths()`.
- The `id` field is optional for header cells; only cells that need to be referenced later (data columns) require IDs.
- Cell merging is applied immediately after the cell is written, ensuring consistent visual appearance.
- The builder supports unlimited rows and columns, though typical headers are 1-3 rows deep.
- Header text can include newline characters (`\n`) for multi-line header cells (e.g., `"Pallet\nNo"`).
- The returned `column_map` uses column text as keys, which may not be unique if headers have duplicate text; `column_id_map` is always unique.


