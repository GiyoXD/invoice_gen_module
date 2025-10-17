# Function Inventory and Testing Plan for `invoice_utils.py` and `packing_list_utils.py`

This document provides a clear inventory of all functions within `invoice_utils.py` and `packing_list_utils.py`. It also outlines a basic approach for testing each function to improve robustness and facilitate refactoring.

## 1. `invoice_utils.py`

### Function Inventory & Testing Plan

#### **`unmerge_row(worksheet, row_num, num_cols)`**
-   **Description:** Unmerges any merged cells that overlap with a specific row.
-   **Testing Approach:**
    -   Create a workbook with various merged cells (e.g., `A1:C1`, `D2:E3`).
    -   Call `unmerge_row` on a row that intersects one or more of these merges.
    -   Assert that `worksheet.merged_cells` no longer contains the unmerged range.
    -   Test edge cases: a row with no merges, an invalid row number.

#### **`unmerge_block(worksheet, start_row, end_row, num_cols)`**
-   **Description:** Unmerges all merged cells within a specified block of rows and columns.
-   **Testing Approach:**
    -   Similar to `unmerge_row`, but call `unmerge_block` on a range of rows.
    -   Assert that all merges within that block are gone, while merges outside the block remain.

#### **`safe_unmerge_block(...)`**
-   **Description:** A more careful version of `unmerge_block` that avoids unmerging cells completely outside the target range.
-   **Testing Approach:**
    -   Create a merge that is partially inside and partially outside the target block.
    -   Call the function and assert that the merge is correctly removed.

#### **`fill_static_row(worksheet, row_num, num_cols, static_content_dict)`**
-   **Description:** Fills a single row with static content from a dictionary.
-   **Testing Approach:**
    -   Create a blank worksheet.
    -   Call `fill_static_row` with a dictionary like `{'1': 'Label', '3': 123}`.
    -   Assert that `worksheet.cell(row=row_num, column=1).value` is `'Label'`.
    -   Assert that `worksheet.cell(row=row_num, column=3).value` is `123`.

#### **`apply_row_merges(worksheet, row_num, num_cols, merge_rules)`**
-   **Description:** Applies horizontal merges to a single row based on a configuration dictionary.
-   **Testing Approach:**
    -   Provide a `merge_rules` dictionary (e.g., `{'1': 3}` to merge A:C).
    -   Call the function and assert that the specified range is now in `worksheet.merged_cells`.

#### **`_apply_cell_style(cell, column_id, sheet_styling_config, DAF_mode)`**
-   **Description:** (Private helper) Applies font, alignment, and number format to a single cell.
-   **Testing Approach:**
    -   Create a single cell and a `sheet_styling_config` dictionary.
    -   Call the function on the cell.
    -   Assert that `cell.font`, `cell.alignment`, and `cell.number_format` have been updated correctly.
    -   Test with and without `DAF_mode`.

#### **`write_grand_total_weight_summary(...)`**
-   **Description:** Calculates and writes a two-row summary for grand total net and gross weights.
-   **Testing Approach:**
    -   Create mock `processed_tables_data` with net/gross weights.
    -   Provide `header_info` and `weight_config`.
    -   Call the function and check if the correct values are written to the correct cells in the worksheet.
    -   Verify that the styling is applied as expected.

#### **`write_header(worksheet, start_row, header_layout_config, sheet_styling_config)`**
-   **Description:** Writes a complete, potentially multi-row header block based on a detailed layout configuration.
-   **Testing Approach:**
    -   This is a larger function. A good test would involve:
    -   Creating a `header_layout_config` list of dictionaries.
    -   Calling `write_header`.
    -   Asserting that the returned `header_info` dictionary is correct (e.g., `column_id_map`).
    -   Checking a few key cells in the worksheet to ensure values, styles, and merges are correct.

#### **`merge_contiguous_cells_by_id(...)`**
-   **Description:** Merges vertical cells in a column that have the same value.
-   **Testing Approach:**
    -   Create a column with data like `['A', 'A', 'B', 'C', 'C', 'C']`.
    -   Call the function on this column.
    -   Assert that the correct merge ranges (`A1:A2`, `C4:C6`) have been created.

#### **`find_footer(worksheet, footer_rules)`**
-   **Description:** Scans the worksheet to locate a footer section by searching for a specific "marker text".
-   **Testing Approach:**
    -   Create a worksheet and place a `marker_text` in a specific cell.
    -   Call `find_footer` with the appropriate rules.
    -   Assert that the function returns the correct row index.
    -   Test cases where the marker is not found.

#### **`write_configured_rows(...)`**
-   **Description:** Writes one or more rows with specified content, styling, and merges based on a configuration list.
-   **Testing Approach:**
    -   Create a `rows_config_list` with various content and styling rules.
    -   Provide `calculated_totals` data.
    -   Call the function and verify that the rows are written and styled correctly.

#### **`apply_explicit_data_cell_merges_by_id(...)`**
-   **Description:** Applies horizontal merges to data cells in a specific row based on column IDs.
-   **Testing Approach:**
    -   Provide `merge_rules_data_cells` with `rowspan` values.
    -   Call the function and check if the correct cells are merged in the specified row.

#### **`_to_numeric(value)`**
-   **Description:** (Private helper) Safely attempts to convert a value to a float or int.
-   **Testing Approach:**
    -   Test with various inputs: integers, floats, string numbers, strings with commas, and non-numeric strings.
    -   Assert that the correct numeric or original value is returned.

#### **`_apply_fallback(row_dict, target_col_idx, mapping_rule, DAF_mode)`**
-   **Description:** (Private helper) Applies a fallback value to a row dictionary based on the DAF mode.
-   **Testing Approach:**
    -   Create a `mapping_rule` with `fallback_on_DAF` and `fallback_on_none` values.
    -   Call the function in both DAF and non-DAF modes and assert that the correct fallback value is applied.

#### **`prepare_data_rows(...)`**
-   **Description:** A highly complex function that transforms raw data into a list of row dictionaries ready to be written to the sheet.
-   **Testing Approach:**
    -   This is a pure data transformation function, making it ideal for unit testing.
    -   Create mock inputs: `data_source`, `dynamic_mapping_rules`, `column_id_map`, etc.
    -   Call `prepare_data_rows`.
    -   Assert that the returned `data_rows_prepared` list has the expected structure and values.
    -   Create separate tests for each `data_source_type` (`aggregation`, `DAF_aggregation`, `processed_tables`).

#### **`parse_mapping_rules(...)`**
-   **Description:** Parses mapping rules from a configuration dictionary.
-   **Testing Approach:**
    -   Create a `mapping_rules` dictionary with different types of rules.
    -   Call the function and assert that the returned dictionary contains the correctly parsed rules.

#### **`write_summary_rows(...)`**
-   **Description:** Calculates and writes ID-driven summary rows for different leather types.
-   **Testing Approach:**
    -   Provide mock `all_tables_data` and `table_keys`.
    -   Call the function and verify that the summary rows are calculated and written correctly.

#### **`write_footer_row(...)`**
-   **Description:** Writes a fully configured footer row, including styling, borders, merges, and summed totals.
-   **Testing Approach:**
    -   Provide `header_info`, `sum_ranges`, and `footer_config`.
    -   Call the function and check if the footer row is generated with the correct formulas, text, and styling.

#### **`_style_row_before_footer(...)`**
-   **Description:** (Private helper) Applies column-specific styles and borders to the static row before the footer.
-   **Testing Approach:**
    -   Create a worksheet with a row of data.
    -   Call the function and verify that the correct styles and borders are applied to the cells in that row.

#### **`fill_invoice_data(...)`**
-   **Description:** The "god function" that orchestrates the entire process of populating a sheet's data section.
-   **Testing Approach:**
    -   **Unit testing this function directly is very difficult due to its high complexity and many side effects (modifying the worksheet).**
    -   This function should be tested via **integration tests**. Create a full test case with a template file, data, and config, run the `SingleTableProcessor` (which calls this function), and then inspect the final generated Excel file to see if it's correct.
    -   The primary goal should be to refactor this function into smaller, testable units as outlined in `docs/invoice_util/refactoring_plan.md`.

#### **`apply_column_widths(worksheet, sheet_styling_config, header_map)`**
-   **Description:** Sets column widths based on the configuration.
-   **Testing Approach:**
    -   Create a `sheet_styling_config` with `column_widths`.
    -   Provide a `header_map`.
    -   Call the function and assert that the column widths are set correctly in the worksheet.

#### **`apply_row_heights(...)`**
-   **Description:** Sets row heights based on the configuration for header, data, footer, and specific rows.
-   **Testing Approach:**
    -   Create a `sheet_styling_config` with `row_heights`.
    -   Call the function and verify that the row heights are set correctly for different sections of the sheet.

---

## 2. `packing_list_utils.py`

### Function Inventory & Testing Plan

#### **`calculate_rows_to_generate(packing_list_data, sheet_config)`**
-   **Description:** Pre-calculates the total number of rows that will be needed for the entire packing list.
-   **Testing Approach:**
    -   This is a pure calculation function, perfect for unit testing.
    -   Create mock `packing_list_data` and `sheet_config` dictionaries.
    -   Call the function and assert that it returns the expected integer value.
    -   Test with different numbers of tables and rows per table.

#### **`generate_full_packing_list(worksheet, start_row, packing_list_data, sheet_config)`**
-   **Description:** A single, large function that generates the *entire* packing list from start to finish. It has been largely superseded by the `MultiTableProcessor`.
-   **Testing Approach:**
    -   Similar to `fill_invoice_data`, this monolithic function is very difficult to unit test.
    -   It should be tested with an **integration test**: create a test case that runs this function and then validates the output Excel file.
    -   Since its logic has been replaced by `MultiTableProcessor`, extensive testing might not be necessary unless it's still being used in some parts of the application. If it is, the focus should be on refactoring it to use the processor pattern.