# Builder Architecture: `data_table_builder.py`

This document explains the structure and purpose of the `DataTableBuilder` class, which is responsible for populating worksheet data tables with invoice data, applying styles, formulas, and handling both static and dynamic content.

## Overview

The `DataTableBuilder` is a core component of the invoice generation system that bridges data preparation and worksheet rendering. It takes prepared data and configuration, then writes it to Excel worksheets with proper formatting, formulas, merging, and styling.

- **Purpose**: To construct and populate data tables in Excel worksheets with invoice data while applying appropriate styling and calculations.
- **Pattern**: Builder pattern with **Pure Bundle Architecture** - uses config bundles with `@property` accessors.
- **Key Responsibility**: Translates business data into formatted Excel rows with formulas, merges, and styles.

## Architecture: Pure Bundle Pattern

`DataTableBuilder` uses the **Bundle Cascade Pattern** where configuration bundles flow from `LayoutBuilder` without extraction. This provides:
- ✅ **Tiny constructor** (5 lines vs 50 lines)
- ✅ **Zero maintenance** for new configs
- ✅ **Clean access** via `@property` decorators
- ✅ **Infinite extensibility**

See `docs/BUNDLE_CASCADE_PATTERN.md` for full details.

## `DataTableBuilder` Class

### `__init__(...)` - The Constructor

The constructor uses the pure bundle pattern: stores bundles without extraction, uses properties for access.

- **Purpose**: To configure the builder with config bundles from LayoutBuilder (Bundle Cascade Pattern).
- **Parameters**:
    - `worksheet: Worksheet`: The `openpyxl` Worksheet object where data will be written.
    - `style_config: Dict[str, Any]`: Style configuration bundle containing:
        - `styling_config`: StylingConfigModel instance
    - `context_config: Dict[str, Any]`: Runtime context bundle containing:
        - `sheet_name`: Sheet being processed
        - `all_sheet_configs`: All sheet configurations
        - `grand_total_pallets`: Global pallet count
        - `args`: Command-line arguments (for DAF/custom flags)
        - `is_last_table`: Whether this is the last table
    - `layout_config: Dict[str, Any]`: Layout configuration bundle containing:
        - `sheet_config`: Complete sheet configuration
        - `add_blank_after_header`: Insert blank row after header
        - `static_content_after_header`: Static content after header
        - `add_blank_before_footer`: Insert blank row before footer
        - `static_content_before_footer`: Static content before footer
        - `merge_rules_after_header`: Merge rules after header
        - `merge_rules_before_footer`: Merge rules before footer
        - `merge_rules_footer`: Footer merge rules
        - `data_cell_merging_rules`: Data cell merge rules
        - `max_rows_to_fill`: Maximum rows to fill
    - `data_config: Dict[str, Any]`: Data source bundle containing:
        - `data_source`: Actual data to write
        - `data_source_type`: Type of data source (aggregation, DAF, etc.)
        - `header_info`: Header structure information
        - `mapping_rules`: Data-to-column mapping rules
        - `all_tables_data`: All tables data (multi-table)
        - `table_keys`: Table keys (multi-table)

- **Instance Variables Initialized**:
    - **Config Bundles**: `style_config`, `context_config`, `layout_config`, `data_config` stored as-is
    - **Output State Variables** (build results, not configs):
        - `actual_rows_to_process`, `data_rows_prepared`, `col1_index`, `num_static_labels`
        - `desc_col_idx`, `local_chunk_pallets`, `dynamic_desc_used`
        - `row_after_header_idx`, `data_start_row`, `data_end_row`, `row_before_footer_idx`, `footer_row_final`

- **Property Accessors** (21 properties for clean access):
    - `sheet_name`, `sheet_config`, `all_sheet_configs`
    - `data_source`, `data_source_type`, `header_info`, `mapping_rules`
    - `sheet_styling_config`, `add_blank_after_header`, `static_content_after_header`
    - `add_blank_before_footer`, `static_content_before_footer`
    - `merge_rules_after_header`, `merge_rules_before_footer`, `merge_rules_footer`
    - `max_rows_to_fill`, `grand_total_pallets`, `custom_flag`, `data_cell_merging_rules`
    - `DAF_mode`, `all_tables_data`, `table_keys`, `is_last_table`

### `build(self) -> Tuple[bool, int, int, int, int]` - The Main Build Method

This is the primary method that orchestrates the entire data table building process from start to finish.

- **Purpose**: To construct the complete data table by inserting rows, filling data, applying styles, formulas, and merges, then returning position information for subsequent processing.
- **Process Flow**:
    1. **Initialization**: Calculate pallet counts from data source and initialize tracking variables.
    2. **Validation**: Verify that `header_info` contains required keys (`second_row_index`, `column_map`, `num_columns`).
    3. **Configuration Parsing**: Extract column indices, parse mapping rules into static values, dynamic mappings, and formula definitions.
    4. **Data Preparation**: Call `prepare_data_rows()` to transform raw data into write-ready row dictionaries.
    5. **Row Calculation**: Calculate total rows needed including blank rows, data rows, and footer.
    6. **Bulk Row Insertion**: Insert all required rows at once for single-table modes (improves performance).
    7. **Data Writing Loop**: For each row:
        - Write static label values (first column).
        - Write dynamic data values from prepared data.
        - Apply cell-level styling using `apply_cell_style()`.
        - Write formula cells (e.g., calculations like subtotals).
        - Apply data cell merging rules if specified.
    8. **Column Merging**: Merge contiguous cells in description, pallet info, and HS code columns.
    9. **Special Row Filling**: Fill the row before footer with static content and special styling.
    10. **Footer Height Application**: Apply configured row height to footer row.
    11. **Merge Application**: Apply merge rules to after-header row, before-footer row, and footer row.
    12. **Row Height Application**: Apply all row heights for header, data, and footer rows.

- **Return Value**: `Tuple[bool, int, int, int, int]`
    - `[0] bool`: Success status - `True` if data table was built successfully, `False` if critical error occurred.
    - `[1] int`: Footer row position - the Excel row index where the footer was placed.
    - `[2] int`: Data start row - the Excel row index where data rows begin.
    - `[3] int`: Data end row - the Excel row index where data rows end.
    - `[4] int`: Local chunk pallets - the sum of pallet counts for this data chunk.

- **Error Handling**:
    - Invalid `header_info`: Returns `(False, -1, -1, -1, 0)`.
    - Bulk insert/unmerge errors: Returns `(False, fallback_row, -1, -1, 0)`.
    - Data filling errors: Returns `(False, footer_row_final + 1, data_start_row, data_end_row, 0)`.

### `_apply_footer_row_height(self, footer_row: int)` - Helper Method

A private helper method that applies the configured row height to a footer row.

- **Purpose**: To set the footer row height based on configuration, with logic to optionally match the header height.
- **Parameters**:
    - `footer_row: int`: The Excel row index of the footer row to apply height to.
- **Logic**:
    1. Checks if `sheet_styling_config` and `rowHeights` configuration exist.
    2. Checks `footer_matches_header_height` flag (default: `True`).
    3. If flag is `True`, uses header height for footer; otherwise uses explicit footer height.
    4. Applies the calculated height to the worksheet row dimensions.
- **Graceful Handling**: Silently skips if configuration is missing or height values are invalid.

## Data Flow

```
Data Source (Dict/List)
        ↓
parse_mapping_rules() → Static values, Dynamic mappings, Formulas
        ↓
prepare_data_rows() → List of row dictionaries
        ↓
Row Calculation → Determine positions of all rows
        ↓
Bulk Insert → Insert all rows at once
        ↓
Data Writing Loop → Write values + formulas + styles
        ↓
Merging & Heights → Apply visual formatting
        ↓
Return positions → Footer, Data Start, Data End
```

## Key Design Decisions

### 1. **Single Bulk Insert**
Instead of inserting rows one at a time, the builder calculates the total rows needed and inserts them all at once. This significantly improves performance for large datasets.

### 2. **Formula as Data**
Formulas are defined in mapping rules and written during the data loop with dynamic cell references. This allows flexible formula definitions without hardcoding cell positions.

### 3. **Separation of Concerns**
- **Data Preparation**: Handled by `data_preparer.py` functions.
- **Data Writing**: Handled by `DataTableBuilder`.
- **Footer Building**: Now handled by `FooterBuilder` (called by `LayoutBuilder`).
- **Styling**: Handled by `styling/style_applier.py` functions.

### 4. **Multi-Mode Support**
The builder supports multiple data source types:
- `"aggregation"`: Single-table standard data.
- `"DAF_aggregation"`: Delivery At Frontier mode with special styling.
- `"custom_aggregation"`: Custom processing with special rules.

### 5. **Flexible Row Insertion**
The builder can insert optional blank rows and static content rows before/after the data section, with independent merge rules for each.

## Dependencies

- **Data Preparation**: `invoice_generator.data.data_preparer` - Prepares raw data into write-ready format.
- **Layout Utilities**: `invoice_generator.utils.layout` - Handles cell merging, unmerging, and column widths.
- **Style Application**: `invoice_generator.styling.style_applier` - Applies fonts, borders, alignments, and row heights.
- **Style Configuration**: `invoice_generator.styling.models.StylingConfigModel` - Type-safe styling configuration model.
- **openpyxl**: Core Excel manipulation library for worksheets, cells, and styles.

## Usage Example

```python
from invoice_generator.builders.data_table_builder import DataTableBuilder
from invoice_generator.styling.models import StylingConfigModel

# Assuming worksheet, config, and data are already prepared
builder = DataTableBuilder(
    worksheet=worksheet,
    sheet_name="Invoice",
    sheet_config=invoice_config,
    all_sheet_configs=all_configs,
    data_source=invoice_data,
    data_source_type="aggregation",
    header_info=header_info,
    mapping_rules=mapping_rules,
    sheet_styling_config=styling_config,
    add_blank_before_footer=True,
    static_content_before_footer={"1": "Subtotal:"},
    merge_rules_footer={"1": 3}  # Merge first 3 columns in footer
)

# Build the data table
success, footer_row, data_start, data_end, pallets = builder.build()

if success:
    print(f"Data table built: rows {data_start} to {data_end}, footer at {footer_row}")
    print(f"Total pallets: {pallets}")
else:
    print("Failed to build data table")
```

## Notes

- The builder is designed to work with pre-configured templates where headers already exist.
- Row insertion only occurs in single-table modes; multi-table modes expect rows to be pre-inserted by the orchestrator.
- The builder returns position information rather than directly calling `FooterBuilder`, following the Director pattern where `LayoutBuilder` orchestrates the sequence.
- The `DAF_mode` flag affects both styling rules and content presentation throughout the build process.

