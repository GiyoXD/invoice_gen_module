# Builder Architecture: `data_table_builder.py`

This document explains the structure and purpose of the `DataTableBuilder` class, which is responsible for populating worksheet data tables with invoice data, applying styles, formulas, and handling both static and dynamic content.

## Overview

The `DataTableBuilder` is a core component of the invoice generation system that bridges data preparation and worksheet rendering. It takes prepared data and configuration, then writes it to Excel worksheets with proper formatting, formulas, merging, and styling.

- **Purpose**: To construct and populate data tables in Excel worksheets with invoice data while applying appropriate styling and calculations.
- **Pattern**: Builder pattern - constructs complex data table structures step-by-step.
- **Key Responsibility**: Translates business data into formatted Excel rows with formulas, merges, and styles.

## `DataTableBuilder` Class

### `__init__(...)` - The Constructor

The constructor initializes the builder with all necessary data, configuration, and context needed to build a complete data table section in an Excel worksheet.

- **Purpose**: To configure the builder with worksheet context, data sources, styling rules, and layout specifications.
- **Parameters**:
    - `worksheet: Worksheet`: The `openpyxl` Worksheet object where data will be written.
    - `sheet_name: str`: The name of the worksheet being processed (for logging and debugging).
    - `sheet_config: Dict[str, Any]`: The complete configuration for this specific sheet, including footer configurations and merge rules.
    - `all_sheet_configs: Dict[str, Any]`: The full configuration dictionary for all sheets (allows cross-sheet references if needed).
    - `data_source: Union[Dict[str, List[Any]], Dict[Tuple, Dict[str, Any]]]`: The actual data to be written, can be a flat dictionary of lists or a nested dictionary structure.
    - `data_source_type: str`: Indicates the type of data source (e.g., `"aggregation"`, `"DAF_aggregation"`, `"custom_aggregation"`).
    - `header_info: Dict[str, Any]`: Information about the header structure including:
        - `second_row_index`: The row index of the second header row (data starts after this).
        - `column_map`: Maps column names to their indices.
        - `column_id_map`: Maps column IDs (e.g., `"col_desc"`) to their indices.
        - `num_columns`: Total number of columns in the table.
    - `mapping_rules: Dict[str, Any]`: Rules for mapping data to columns, including static values, dynamic mappings, and formula definitions.
    - `sheet_styling_config: Optional[StylingConfigModel]`: The styling configuration model containing font, alignment, border, and row height specifications.
    - `add_blank_after_header: bool`: Whether to insert a blank row immediately after the header (default: `False`).
    - `static_content_after_header: Optional[Dict[str, Any]]`: Content to populate in the row after the header, if any.
    - `add_blank_before_footer: bool`: Whether to insert a blank row immediately before the footer (default: `False`).
    - `static_content_before_footer: Optional[Dict[str, Any]]`: Content to populate in the row before the footer, if any.
    - `merge_rules_after_header: Optional[Dict[str, int]]`: Cell merge rules for the row after header.
    - `merge_rules_before_footer: Optional[Dict[str, int]]`: Cell merge rules for the row before footer.
    - `merge_rules_footer: Optional[Dict[str, int]]`: Cell merge rules for the footer row itself.
    - `max_rows_to_fill: Optional[int]`: Maximum number of data rows to process (used for template constraints).
    - `grand_total_pallets: int`: The global pallet count total (may be used in footers).
    - `custom_flag: bool`: Indicates if custom processing mode is active.
    - `data_cell_merging_rules: Optional[Dict[str, Any]]`: Rules for merging specific data cells within rows.
    - `DAF_mode: Optional[bool]`: Delivery At Frontier mode flag that affects styling and content rules (default: `False`).
    - `all_tables_data: Optional[Dict[str, Any]]`: Data for all tables (used in multi-table scenarios).
    - `table_keys: Optional[List[str]]`: List of table keys for multi-table processing.
    - `is_last_table: bool`: Flag indicating if this is the last table being processed (affects footer behavior).

- **Instance Variables Initialized**:
    - **Tracking Variables**: `actual_rows_to_process`, `data_rows_prepared`, `col1_index`, `num_static_labels`, `desc_col_idx`, `local_chunk_pallets`, `dynamic_desc_used`
    - **Row Position Trackers**: `row_after_header_idx`, `data_start_row`, `data_end_row`, `row_before_footer_idx`, `footer_row_final`

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

