# Builder Architecture: `footer_builder.py`

This document explains the structure and purpose of the `FooterBuilder` class, which is responsible for generating footer rows in invoice worksheets, including totals, pallet counts, SUM formulas, and optional summary sections.

## Overview

The `FooterBuilder` is a specialized component that constructs footer rows at the end of data table sections. It handles different footer types (regular vs. grand total), applies SUM formulas across multiple data ranges, displays pallet counts, and can generate additional summary sections for specific scenarios.

- **Purpose**: To build and style footer rows that summarize data tables with totals, formulas, and metadata.
- **Pattern**: Builder pattern - constructs footer structures with different configurations.
- **Key Responsibility**: Generates footer content with proper formulas, text labels, styling, and optional add-on sections.

## `FooterBuilder` Class

### `__init__(...)` - The Constructor

The constructor initializes the builder with all information needed to construct an appropriate footer for the data table.

- **Purpose**: To configure the footer builder with position, data ranges, configuration, and context for footer generation.
- **Parameters**:
    - `worksheet: Worksheet`: The `openpyxl` Worksheet object where the footer will be written.
    - `footer_row_num: int`: The Excel row index where the footer should be placed.
    - `header_info: Dict[str, Any]`: Header structure information containing:
        - `num_columns`: Total number of columns in the table.
        - `column_id_map`: Maps column IDs (e.g., `"col_total"`) to their indices.
        - Other header metadata.
    - `sum_ranges: List[Tuple[int, int]]`: List of row range tuples to include in SUM formulas. Each tuple is `(start_row, end_row)`. Multiple ranges are combined with commas in the SUM formula.
    - `footer_config: Dict[str, Any]`: Footer configuration containing:
        - `type`: Footer type (`"regular"` or `"grand_total"`).
        - `total_text`: Text label for the footer (e.g., `"TOTAL:"`).
        - `total_text_column_id`: Column ID or index for the total text label.
        - `pallet_count_column_id`: Column ID or index for pallet count display.
        - `sum_column_ids`: List of column IDs where SUM formulas should be applied.
        - `merge_rules`: List of merge rule dictionaries with `start_column_id` and `colspan`.
        - `add_ons`: List of add-on features (e.g., `["summary"]`).
    - `pallet_count: int`: The total number of pallets to display in the footer.
    - `override_total_text: Optional[str]`: If provided, overrides the default total text from config (default: `None`).
    - `DAF_mode: bool`: Delivery At Frontier mode flag that affects summary add-on behavior (default: `False`).
    - `sheet_styling_config: Optional[StylingConfigModel]`: Styling configuration model for footer cell styles and row heights.
    - `all_tables_data: Optional[Dict[str, Any]]`: Complete data for all tables (used by summary add-on in multi-table scenarios).
    - `table_keys: Optional[List[str]]`: List of table keys for multi-table processing (used by summary add-on).
    - `mapping_rules: Optional[Dict[str, Any]]`: Mapping rules for data columns (used by summary add-on).
    - `sheet_name: Optional[str]`: Name of the sheet being processed (used to determine if summary add-on applies).
    - `is_last_table: bool`: Flag indicating if this is the last table in a multi-table sheet (controls summary add-on display).
    - `dynamic_desc_used: bool`: Flag indicating if dynamic descriptions were used in the data table (affects summary add-on eligibility).

### `build(self) -> int` - The Main Build Method

This is the primary method that orchestrates footer construction and returns the next available row.

- **Purpose**: To build the complete footer section including main footer row and any add-ons, then return the next available row index.
- **Process Flow**:
    1. **Validation**: Check if `footer_config` exists and `footer_row_num` is valid.
    2. **Footer Type Determination**: Read `type` from footer config (defaults to `"regular"`).
    3. **Footer Building**: Call appropriate private method:
        - `_build_regular_footer()` for standard footers.
        - `_build_grand_total_footer()` for grand total footers.
    4. **Row Height Application**: Apply configured row height to the footer row.
    5. **Add-on Processing**: Check for add-ons in config:
        - If `"summary"` is in add-ons list, call `_build_summary_add_on()`.
    6. **Return Next Row**: Return the row index after the footer (and any add-ons).

- **Return Value**: `int`
    - The Excel row index of the next available row after the footer section.
    - Returns `-1` if footer config is invalid or an error occurred.

- **Error Handling**:
    - Missing or invalid footer config: Returns `-1`.
    - Exceptions during footer building: Prints error message and returns `-1`.

### `_build_regular_footer(self, current_footer_row: int)` - Regular Footer Builder

Builds a standard footer row with "TOTAL:" label, pallet count, and SUM formulas.

- **Purpose**: To construct a regular footer for single-table or intermediate table sections.
- **Process**:
    1. **Unmerge Row**: Clear any existing merges in the footer row.
    2. **Total Text**: Write total label (e.g., `"TOTAL:"`) to the configured column.
    3. **Pallet Count**: Write pallet count text (e.g., `"5 PALLETS"`) if pallet count > 0.
    4. **SUM Formulas**: For each column ID in `sum_column_ids`:
        - Build a SUM formula covering all ranges in `sum_ranges`.
        - Formula format: `=SUM(A5:A10,A15:A20)` for multiple ranges.
    5. **Apply Styles**: Apply footer cell styling to all cells in the footer row.
    6. **Apply Merges**: Apply merge rules from footer config.

- **Column ID Resolution**: Supports three formats for column IDs:
    - **String ID** (preferred): `"col_total"` → looked up in `column_id_map`.
    - **String Number**: `"3"` → converted to integer, then to 1-based index.
    - **Integer**: `3` → converted to 1-based index (adds 1).

### `_build_grand_total_footer(self, current_footer_row: int)` - Grand Total Footer Builder

Builds a grand total footer row, typically used at the end of multi-table sections.

- **Purpose**: To construct a grand total footer that summarizes all tables in a sheet.
- **Implementation**: Nearly identical to `_build_regular_footer()` except:
    - Default total text is `"TOTAL OF:"` instead of `"TOTAL:"`.
    - Override text from constructor still takes precedence.
- **Use Case**: Used when `footer_config["type"]` is `"grand_total"`, typically for the final footer in Packing List sheets with multiple product tables.

### `_build_summary_add_on(self, current_footer_row: int) -> int` - Summary Add-on Builder

Builds an optional summary section below the footer, listing all products with their totals.

- **Purpose**: To generate a detailed summary table in DAF mode for Packing List sheets after all product tables.
- **Conditions**: Only executes if ALL of the following are true:
    - `DAF_mode` is `True`
    - `dynamic_desc_used` is `True` (data used dynamic descriptions, not static fallbacks)
    - `sheet_name` is `"Packing list"`
    - `is_last_table` is `True` (this is the last table in the sheet)
    - `all_tables_data`, `table_keys`, and `mapping_rules` are all provided
- **Delegation**: Calls `write_summary_rows()` from `utils.layout` module, passing all necessary context.
- **Return Value**: Returns the next available row index after the summary section.

### `_apply_footer_cell_style(self, cell, col_id)` - Helper Method

Applies styling to a single footer cell.

- **Purpose**: To apply consistent footer styling to individual cells using the styling configuration.
- **Parameters**:
    - `cell`: The openpyxl Cell object to style.
    - `col_id`: The column ID for context-aware styling.
- **Context**: Creates a context dictionary with:
    - `col_id`: Column identifier for column-specific styling.
    - `col_idx`: Column index from the cell.
    - `is_footer`: Flag set to `True` to trigger footer-specific style rules.
- **Delegation**: Calls `apply_cell_style()` from `styling.style_applier` module.

### `_apply_footer_row_height(self, footer_row: int)` - Helper Method

Applies the configured row height to the footer row.

- **Purpose**: To set the footer row height based on configuration, with logic to optionally match header height.
- **Parameters**:
    - `footer_row: int`: The Excel row index of the footer row.
- **Height Resolution Logic**:
    1. Check if `footer_matches_header_height` flag is `True` (default: `True`).
    2. If `True`, use header height from `rowHeights.header`.
    3. If `False` or header height not found, use explicit `rowHeights.footer` value.
    4. Apply the determined height to the row dimensions.
- **Graceful Handling**: Silently skips if configuration is missing or height values are invalid.

## Footer Types Comparison

| Aspect | Regular Footer | Grand Total Footer |
|--------|---------------|-------------------|
| **Default Text** | `"TOTAL:"` | `"TOTAL OF:"` |
| **Use Case** | Single tables, intermediate tables | Final footer in multi-table sheets |
| **SUM Ranges** | Sums assigned data ranges | Sums all table ranges combined |
| **Config Type** | `"type": "regular"` | `"type": "grand_total"` |

Both types support the same features: pallet counts, SUM formulas, merge rules, and styling.

## Data Flow

```
FooterBuilder.build()
        ↓
Determine Footer Type
        ↓
    ┌───────────┴───────────┐
    ↓                       ↓
_build_regular_footer   _build_grand_total_footer
    │                       │
    ├─ Unmerge Row          ├─ Unmerge Row
    ├─ Write Total Text     ├─ Write Total Text
    ├─ Write Pallet Count   ├─ Write Pallet Count
    ├─ Write SUM Formulas   ├─ Write SUM Formulas
    ├─ Apply Cell Styles    ├─ Apply Cell Styles
    └─ Apply Merge Rules    └─ Apply Merge Rules
        ↓                       ↓
        └───────────┬───────────┘
                    ↓
        Apply Row Height
                    ↓
        Check for Add-ons
                    ↓
        _build_summary_add_on (if applicable)
                    ↓
        Return Next Row Index
```

## Key Design Decisions

### 1. **Flexible Column ID Resolution**
The builder supports multiple formats for specifying columns (string IDs, string numbers, integers) to accommodate different configuration styles and legacy configurations.

### 2. **Multi-Range SUM Support**
SUM formulas can span multiple non-contiguous ranges, enabling footers to sum across separated data sections in multi-table sheets.

### 3. **Conditional Summary Add-on**
The summary section only appears in very specific scenarios (DAF mode, Packing List, last table), preventing unnecessary clutter in other contexts.

### 4. **Footer Type Abstraction**
Regular and grand total footers share the same implementation structure, differing only in default text, making the code maintainable and consistent.

### 5. **Context-Aware Styling**
Footer cells are styled with `is_footer: True` context, allowing the styling system to apply footer-specific borders, fonts, and alignments.

## SUM Formula Generation

The builder generates Excel SUM formulas that can span multiple ranges:

```python
# Single range:
sum_ranges = [(5, 10)]
# Formula: =SUM(A5:A10)

# Multiple ranges (multi-table scenario):
sum_ranges = [(5, 10), (15, 20), (25, 30)]
# Formula: =SUM(A5:A10,A15:A20,A25:A30)
```

Each range tuple represents `(start_row, end_row)` inclusive.

## Dependencies

- **Layout Utilities**: `invoice_generator.utils.layout` - Handles row unmerging and summary row writing.
- **Style Application**: `invoice_generator.styling.style_applier` - Applies cell-level styling.
- **Style Configuration**: `invoice_generator.styling.models.StylingConfigModel` - Type-safe styling model.
- **openpyxl**: Core Excel manipulation library for cells, merging, and formulas.

## Usage Example

```python
from invoice_generator.builders.footer_builder import FooterBuilder

# Configuration for a regular footer
footer_config = {
    "type": "regular",
    "total_text": "TOTAL:",
    "total_text_column_id": "col_desc",
    "pallet_count_column_id": "col_pallet",
    "sum_column_ids": ["col_qty", "col_total", "col_weight"],
    "merge_rules": [
        {"start_column_id": "col_desc", "colspan": 3}
    ]
}

# Data ranges to sum (multiple tables)
sum_ranges = [
    (5, 10),   # First table: rows 5-10
    (15, 20),  # Second table: rows 15-20
    (25, 30)   # Third table: rows 25-30
]

# Create and build the footer
builder = FooterBuilder(
    worksheet=worksheet,
    footer_row_num=35,
    header_info=header_info,
    sum_ranges=sum_ranges,
    footer_config=footer_config,
    pallet_count=150,
    sheet_styling_config=styling_config
)

next_row = builder.build()

if next_row > 0:
    print(f"Footer built at row 35, next available row: {next_row}")
else:
    print("Failed to build footer")
```

## Example with Grand Total Footer and Summary Add-on

```python
# Grand total footer for final table in DAF Packing List
footer_config = {
    "type": "grand_total",
    "total_text_column_id": "col_desc",
    "pallet_count_column_id": "col_pallet",
    "sum_column_ids": ["col_qty", "col_total"],
    "add_ons": ["summary"]  # Enable summary section
}

builder = FooterBuilder(
    worksheet=worksheet,
    footer_row_num=50,
    header_info=header_info,
    sum_ranges=[(5, 15), (20, 30), (35, 45)],  # All three tables
    footer_config=footer_config,
    pallet_count=300,
    DAF_mode=True,
    sheet_name="Packing list",
    is_last_table=True,
    dynamic_desc_used=True,
    all_tables_data=all_tables_data,
    table_keys=["PRODUCT_A", "PRODUCT_B", "PRODUCT_C"],
    mapping_rules=mapping_rules,
    sheet_styling_config=styling_config
)

next_row = builder.build()
# Footer at row 50, summary section from row 51 to next_row - 1
```

## Notes

- The builder assumes the footer row is already inserted by the caller (typically `LayoutBuilder` or `DataTableBuilder`).
- Footer merging is applied after content and styling, ensuring merge boundaries don't affect content population.
- The summary add-on is highly specialized for DAF Packing List sheets and will be skipped in most scenarios.
- Column ID resolution supports legacy numeric formats for backward compatibility with older configurations.
- Row height application uses the same logic as `DataTableBuilder`, ensuring visual consistency between data rows and footer rows.
- The `override_total_text` parameter allows dynamic text based on context (e.g., "SUBTOTAL:" for intermediate footers).

