# Utility Modules Documentation

This document provides a detailed breakdown of the two main utility modules, `invoice_utils.py` and `packing_list_utils.py`. Understanding their contents is crucial for the refactoring process.

---

## 1. `invoice_utils.py`

This is a large, multi-purpose module that acts as the primary toolkit for all Excel manipulation. It contains a mix of generic, low-level functions and highly specific, complex business logic. It is the main target for our refactoring.

### Core Responsibilities:
- Writing data to cells.
- Applying styles (fonts, borders, alignment, number formats).
- Creating complex headers and footers from configuration.
- Calculating and writing summaries and totals.
- Handling merged cells.
- Preparing and transforming data before it's written to the sheet.

### Function Breakdown:

#### Low-Level Cell & Range Manipulation
- **`unmerge_row(worksheet, row_num, num_cols)`**: Unmerges any cell ranges that overlap with a specific row.
- **`unmerge_block(worksheet, start_row, end_row, num_cols)`**: Unmerges all cell ranges within a specified block of rows and columns.
- **`safe_unmerge_block(...)`**: A safer version of `unmerge_block` that is more careful about not unmerging cells that are completely outside the target range.
- **`merge_contiguous_cells_by_id(...)`**: A specialized function that finds and merges vertical cells within a single column that have the same value (e.g., merging all "Item A" cells in the "Item" column).
- **`apply_horizontal_merge(worksheet, row_num, ...)`**: Applies horizontal merges to a single row based on a set of rules from the configuration.
- **`apply_explicit_data_cell_merges_by_id(...)`**: Applies horizontal merges to data cells in a specific row, driven by column IDs.

#### Styling and Formatting
- **`_apply_cell_style(cell, ...)`**: A private helper that applies font, alignment, and number format to a single cell based on its column ID and the sheet's styling configuration.
- **`_style_row_before_footer(...)`**: A private helper with a very specific purpose: to apply a unique set of styles and borders to the static row that appears just before a footer.
- **`apply_column_widths(worksheet, ...)`**: Sets the width for multiple columns based on the styling configuration.
- **`apply_row_heights(worksheet, ...)`**: Sets the height for header, data, and footer rows based on the styling configuration.

#### Data Preparation and Transformation
- **`_to_numeric(value)`**: A private helper that safely attempts to convert any value (including strings with commas) into a numeric type (`int` or `float`).
- **`_apply_fallback(row_dict, ...)`**: A private helper that applies a fallback value to a data row if the primary data source is empty, with different logic for DAF mode.
- **`parse_mapping_rules(...)`**: A critical function that reads the complex `mappings` section of the configuration and translates it into a structured set of rules that the data-filling functions can understand.
- **`prepare_data_rows(...)`**: A highly complex "kitchen" function. It takes the raw data source and the mapping rules and transforms the data into a list of row dictionaries, ready to be written to the worksheet. It handles different data source types (`aggregation`, `DAF_aggregation`, `processed_tables`) and applies fallback values.

#### Content Writing and Layout
- **`fill_static_row(worksheet, ...)`**: Fills a single row with static content (e.g., a blank row with a label in one cell).
- **`find_footer(worksheet, ...)`**: Scans the worksheet to locate a footer section by searching for a specific "marker text" defined in the configuration.
- **`calculate_header_dimensions(header_layout)`**: A simple helper that calculates the total row and column span of a header based on its layout configuration.
- **`write_header(worksheet, ...)`**: A major function responsible for writing a complete, potentially multi-row header block based on a detailed layout configuration. It writes text, applies styles, and handles all cell merges for the header.
- **`write_footer_row(worksheet, ...)`**: A major function for writing a complete footer row. It writes labels, calculates and inserts `SUM` formulas, displays a pallet count, and applies all necessary styles and merges based on the footer configuration.
- **`write_summary_rows(worksheet, ...)`**: A specific business logic function that calculates and writes the "TOTAL OF: BUFFALO LEATHER" and "TOTAL OF: COW LEATHER" summary rows for packing lists.
- **`write_grand_total_weight_summary(worksheet, ...)`**: Another specific business logic function that calculates the grand total net and gross weights and writes a two-row summary.
- **`write_configured_rows(worksheet, ...)`**: A modern, structured function for writing rows. It takes a list of row configurations and handles writing content, applying styles, and managing merges for each row. This represents a more advanced approach compared to the older functions.

#### The Core Orchestrator Function
- **`fill_invoice_data(...)`**: This is the largest and most complex function in the module. It is the original "god function" that orchestrates the entire process of populating a sheet's data section. It calls many of the other helpers to perform its job.
    - It determines the number of rows to insert.
    - It calls `prepare_data_rows` to get the data ready.
    - It loops through each data row and writes the values to the cells.
    - It handles special columns like "No." and "Pallet Info".
    - It applies styles and borders to every cell.
    - It calls `write_footer_row` to add the footer.
    - It triggers post-processing steps like merging contiguous cells.
    - **This function is the primary candidate for being broken down and replaced by the new Manager classes.**

---

## 2. `packing_list_utils.py`

This module is much more focused than `invoice_utils`. It contains high-level functions specifically for generating a packing list, which typically involves multiple repeating tables. It appears to be an older utility module, and its logic has largely been superseded by the `MultiTableProcessor`.

### Core Responsibilities:
- Calculating the total size of a packing list.
- Orchestrating the generation of a full packing list in a single, monolithic function.

### Function Breakdown:

- **`calculate_rows_to_generate(packing_list_data, ...)`**:
    - **Purpose**: To pre-calculate the total number of rows that will be needed for the entire packing list.
    - **Logic**: It counts the number of tables, the number of data rows in each table, and adds rows for headers, footers, and spacing. This is the same logic used by `MultiTableProcessor` to pre-insert rows.

- **`generate_full_packing_list(worksheet, ...)`**:
    - **Purpose**: This is a single, large function that generates the *entire* packing list from start to finish.
    - **Logic**:
        1. It loops through each table in the `packing_list_data`.
        2. For each table, it calls `invoice_utils.write_header`.
        3. It then enters a nested loop to write each data row for that table.
        4. It calls `style_utils.apply_cell_style` for each cell.
        5. It calls `merge_utils.merge_vertical_cells_in_range` to merge cells after writing.
        6. It writes a pre-footer and a main footer for the table by calling `invoice_utils.write_footer_row`.
        7. After processing all tables, it writes a final "GRAND TOTAL" footer.
    - **Analysis**: This function is a prime example of the old, monolithic approach. The `MultiTableProcessor` was created to break down this exact logic into a more structured, step-by-step process.
