# Concrete Processors: `single_table_processor.py`

This document details the `SingleTableProcessor`, a concrete implementation of the `SheetProcessor` abstract base class.

## `SingleTableProcessor`

This processor is designed to handle worksheets that have a relatively simple structure: a header, a single body of data (which can be a table or an aggregation), and an optional footer. It is the default processor used by the orchestrator.

- **Purpose**: To manage the entire lifecycle of populating a standard, single-table worksheet.
- **Inherits from**: `SheetProcessor`

### `process(self) -> bool`

This method contains the step-by-step logic for building the sheet. It acts as a procedural script that calls various utility functions to perform the work.

- **Purpose**: To execute the specific sequence of operations required to generate a single-table sheet.
- **Process**:
    1.  **Configuration Loading**: It begins by extracting all necessary configuration values from `self.sheet_config`. This includes flags, static content definitions, merge rules, and styling information.
    2.  **Header Writing**:
        - It calls `invoice_utils.write_header()` to write the main table header at the `start_row` defined in the configuration.
        - If the header cannot be written, it returns `False` to halt the process.
    3.  **Data Source Selection**:
        - This is a critical step where the processor determines exactly which piece of data to use from the main `invoice_data` dictionary.
        - It uses `self.data_source_indicator` as the primary key but also checks for command-line flags (`--custom`, `--DAF`) to select the correct data source (e.g., `custom_aggregation_results`, `standard_aggregation_results`, or a specific table from `processed_tables_data`).
    4.  **Data Filling**:
        - If a valid data source is found, it calls the large and complex `invoice_utils.fill_invoice_data()` function.
        - This single function is responsible for:
            - Looping through the data.
            - Writing data rows to the worksheet.
            - Applying styles.
            - Writing footer rows.
            - Handling various merging and spacing rules.
        - It passes a large number of parameters, delegating most of the sheet's construction to this utility function.
    5.  **Finalization**:
        - After `fill_invoice_data()` completes, it performs final adjustments, such as applying column widths using `invoice_utils.apply_column_widths()`.
        - It returns `True` to signal successful completion.
