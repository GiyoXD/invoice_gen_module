# Concrete Processors: `multi_table_processor.py`

This document details the `MultiTableProcessor`, a concrete implementation of the `SheetProcessor` abstract base class designed for more complex layouts.

## `MultiTableProcessor`

This processor handles worksheets that contain multiple, repeating blocks of tables. A common example is a packing list where each "block" represents the contents of a single pallet or container, and each block has its own header, data rows, and footer.

- **Purpose**: To manage the layout and population of worksheets with repeating, variable-height data sections.
- **Inherits from**: `SheetProcessor`

### `process(self) -> bool`

This method orchestrates the complex process of building a multi-table sheet. Its primary challenge is managing the vertical space, as the number of rows for each table is not known in advance.

- **Purpose**: To execute the specific sequence of operations required to generate a multi-table sheet.
- **Process**:
    1.  **Data Retrieval**: It retrieves the `processed_tables_data` dictionary from the main `invoice_data`. This dictionary contains the data for each individual table, indexed by keys (e.g., `"1"`, `"2"`, etc.).
    2.  **Row Pre-calculation and Insertion (`_pre_calculate_and_insert_rows`)**:
        - This is the most critical and complex part of the processor.
        - It first iterates through all the tables *without writing anything*. Its only goal is to calculate how many rows each table will occupy (header + data rows + footer + blank spaces).
        - Based on these calculations, it inserts the required number of blank rows into the worksheet *before* writing any data. This "makes space" for all the tables at once, preventing the need to insert rows one by one, which is inefficient and breaks cell formatting in `openpyxl`.
        - This pre-calculation step is essential for preserving the template's layout and styling.
    3.  **Main Write Loop**:
        - After the rows have been inserted, the processor initializes a `write_pointer_row` to keep track of the current row.
        - It then loops through each table's data again.
        - **For each table**:
            - It calls `invoice_utils.write_header()` to write the table's header at the current `write_pointer_row`.
            - It calls the main `invoice_utils.fill_invoice_data()` function to populate the data rows and footer for just that single table.
            - It updates the `write_pointer_row` to the next available row after the current table is finished.
    4.  **Summary and Finalization**:
        - After the loop finishes, it may write a final summary section (e.g., grand totals).
        - It applies final styling, such as column widths.
        - It returns `True` to signal successful completion.

### Helper Methods

This processor uses internal helper methods (often prefixed with an underscore) to break down its complex logic.

- **`_pre_calculate_and_insert_rows(self, ...)`**:
    - **Purpose**: To calculate the total number of rows required for all tables and insert them in a single bulk operation.
    - **Logic**:
        1. It iterates through each table's data to count the number of header rows, data rows, footer rows, and spacer rows.
        2. It sums these counts to get a total number of rows to insert.
        3. It calls `worksheet.insert_rows()` once at the beginning of the process.
    - **Analysis**: This "pre-calculation" is the key to this processor's success. By inserting all rows at once, it avoids the performance issues and formatting corruption associated with inserting rows inside a loop.

- **`_write_grand_total_row(self, start_row, header_info, sum_ranges, pallet_count)`**:
    - **Purpose**: To write the final "TOTAL OF:" row that summarizes all the tables processed on the sheet.
    - **Logic**:
        1. It receives the `start_row` where it should write, the `header_info` for column mapping, a list of all `sum_ranges` (the data rows from every table), and the total `pallet_count`.
        2. It calls the generic `invoice_utils.write_footer_row` function to perform the actual writing.
        3. It passes a specific `override_total_text="TOTAL OF:"` to ensure the correct label is used for this summary row.
    - **Analysis**: This method acts as a specialized wrapper around the more generic footer-writing utility, configuring it for the specific purpose of a grand total summary.