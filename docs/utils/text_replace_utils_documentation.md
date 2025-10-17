# Utility Module Documentation: `text_replace_utils.py`

This document provides a detailed breakdown of the `text_replace_utils.py` module. This module is responsible for finding and replacing text and placeholders throughout the workbook.

---

## `text_replace_utils.py`

This module contains a powerful, two-pass replacement engine and the task-runner functions that use it. It is designed to handle simple text swaps, data-driven replacements from the main `invoice_data`, and even complex formula construction.

### Core Responsibilities:
- Finding and replacing text based on a set of rules.
- Handling different match modes (`exact` vs. `substring`).
- Intelligently parsing and formatting date values.
- Safely retrieving data from nested dictionaries and lists.
- Assembling and placing Excel formulas based on the locations of other placeholders.

### Function Breakdown:

#### Section 1: Core Helper Functions

- **`excel_number_to_datetime(excel_num)`**:
    - **Purpose**: To convert an Excel serial number for a date (e.g., `45788`) into a standard Python `datetime` object.
    - **Logic**: It correctly handles Excel's 1900 leap year bug.

- **`format_cell_as_date_smarter(cell, value)`**:
    - **Purpose**: A robust, intelligent function for handling date values.
    - **Logic**:
        1. It checks if the input `value` is already a `datetime` object.
        2. If it's a string, it uses the powerful `dateutil.parser` to try and parse it. This can handle many formats (e.g., "2025-05-11T00:00:00", "11/10/2025").
        3. If it's a number, it assumes it's an Excel serial date and uses `excel_number_to_datetime` to convert it.
        4. If a date is successfully parsed, it places the `datetime` object into the cell and applies the "dd/mm/yyyy" number format, ensuring Excel treats it as a proper date.

- **`_get_nested_data(data_dict, path)`**:
    - **Purpose**: A safe way to retrieve a value from a deeply nested data structure (like `invoice_data`) without causing an error if an intermediate key doesn't exist.
    - **Logic**: It takes a `path` (a list of keys and/or indices, like `["processed_tables_data", "1", "inv_no", 0]`) and traverses the dictionary one level at a time. If any key is not found, it returns `None` instead of crashing.

#### Section 2: The Replacement Engine

- **`find_and_replace(workbook, rules, limit_rows, limit_cols, invoice_data)`**:
    - **Purpose**: This is the core engine that drives all replacement operations. It works in two distinct passes to handle complex dependencies.
    - **Pass 1: Locate Placeholders & Apply Simple Replacements**:
        1. It iterates through every cell within the specified row/column limits.
        2. It checks if the cell's value matches any placeholder defined in the `rules`. If it does, it stores the placeholder and its cell coordinate (e.g., `{'[[NET]]': 'G35'}`) in a `placeholder_locations` dictionary.
        3. At the same time, it applies any *simple* (non-formula) replacement rules. It replaces the cell's content with either a hardcoded `replace` value or a dynamic value retrieved from `invoice_data` using `_get_nested_data`.
        4. If a rule has `is_date: True`, it uses `format_cell_as_date_smarter` to handle the replacement.
    - **Pass 2: Build and Apply Formulas**:
        1. After Pass 1 has found the location of *all* placeholders, this pass iterates through only the `formula_rules`.
        2. For each formula rule, it finds the target cell where the formula should go (e.g., the cell containing `[[GRAND_TOTAL]]`).
        3. It then looks at the `formula_template` (e.g., `SUM({[[NET]]}, {[[TAX]]})`) and identifies the dependent placeholders (`[[NET]]` and `[[TAX]]`).
        4. It uses the `placeholder_locations` dictionary to look up the real cell coordinates for these dependencies (e.g., 'G35' and 'H35').
        5. It replaces the placeholders in the template with the real coordinates to build the final formula string (e.g., `=SUM(G35, H35)`).
        6. Finally, it places this complete formula into the target cell.

#### Section 3: Task-Runner Functions

These are simple, high-level functions that define specific sets of rules and then call the `find_and_replace` engine to execute them.

- **`run_invoice_header_replacement_task(workbook, invoice_data)`**:
    - **Purpose**: To replace all the placeholders typically found in the header of an invoice (e.g., invoice number, date, customer name).
    - **Logic**: It defines a list of data-driven rules that map placeholders like `JFINV` and `[[CUSTOMER_NAME]]` to their corresponding paths in the `invoice_data` dictionary. It then calls the engine, limiting the search to the top 14 rows of the sheet.

- **`run_DAF_specific_replacement_task(workbook)`**:
    - **Purpose**: To perform a series of hardcoded text replacements required for DAF (Delivered at Frontier) invoices.
    - **Logic**: It defines a long list of simple find-and-replace rules (e.g., find "BINH PHUOC", replace with "BAVET"). It then calls the engine to apply these rules across a larger area of the sheet.
