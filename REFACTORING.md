# Consolidated Refactoring Plan for `invoice_generator`

## 1. Objective

To refactor the `invoice_generator` from a monolithic script into a modular, component-based architecture. This will improve maintainability, testability, and extensibility, using the Strategy and Builder design patterns. The immediate focus is the systematic decommissioning of the legacy `invoice_utils.py` and other old files by migrating their logic into the new architecture.

## 2. Target Directory Structure

The `invoice_generator` directory will be reorganized to separate concerns into the following specialized packages:

```
invoice_generator/
|
|-- config/
|   |-- __init__.py
|   |-- loader.py               # Loads and validates the JSON config files
|   `-- models.py               # Defines Pydantic models for config structure
|
|-- processors/
|   |-- __init__.py
|   |-- base_processor.py         # Abstract base class for document generation strategies
|   |-- single_table_processor.py # Strategy for standard invoices
|   `-- multi_table_processor.py  # Strategy for packing lists
|
|-- builders/
|   |-- __init__.py
|   |-- layout_builder.py       # Director: Orchestrates the other builders
|   |-- header_builder.py       # Builds the header section
|   |-- footer_builder.py       # Builds the footer and summary sections
|   `-- table_builder.py        # Builds the main data table(s)
|
|-- styling/
|   |-- __init__.py
|   |-- style_applier.py        # Applies pre-defined styles to cells/rows
|   `-- style_config.py         # Defines reusable style objects (Fonts, Borders, etc.)
|
|-- data/
|   |-- __init__.py
|   `-- data_preparer.py        # Transforms raw data into a writable format
|
|-- utils/
|   |-- __init__.py
|   |-- data_processing.py      # Helpers for data cleaning and preparation
|   |-- excel_operations.py     # Low-level, generic Excel operations (merge, unmerge, etc.)
|   `-- layout.py               # Helpers for row/column manipulation
|
`-- generate_invoice.py           # Main entry point, uses a processor to run the process
```

## 3. Component Responsibilities

*   **`config`:** Loads, validates, and models the JSON configuration files using Pydantic.
*   **`processors`:** Implements the "Strategy" pattern to define the high-level algorithm for creating a document.
*   **`builders`:** Implements the "Builder" pattern to construct the different visual parts of the Excel document (header, table, footer).
*   **`styling`:** Centralizes all styling and formatting logic (fonts, borders, alignments).
*   **`data`:** Handles all data manipulation and preparation before it is written to the sheet.
*   **`utils`:** Provides low-level, reusable helper functions that are not specific to any part of the invoice.

## 4. Execution Plan: Decommissioning Legacy Files

The logic currently in legacy files like `invoice_utils.py` will be migrated into the target architecture as follows:

| Original Logic | Target Module & File | Justification |
| :--- | :--- | :--- |
| `prepare_data_rows`, `parse_mapping_rules` | `data/data_preparer.py` | **Data Preparation:** This logic transforms raw input data into a structured format. |
| Main data writing loop in `fill_invoice_data` | `builders/table_builder.py` | **Content Writing:** This is the core logic for writing the main data table. |
| `write_footer_row`, `write_summary_rows` | `builders/footer_builder.py` | **Footer Generation:** These functions construct the footer and summary sections. |
| `_style_row_before_footer`, style constants | `styling/style_applier.py` | **Styling:** Centralizes all visual formatting logic. |
| `safe_unmerge_block`, `apply_horizontal_merge` | `utils/excel_operations.py` | **Generic Utilities:** Low-level, reusable `openpyxl` helper functions. |
| `_to_numeric`, `_apply_fallback` | `utils/data_processing.py` | **Data Cleaning:** Small helpers for cleaning individual data points. |


## 5. Step-by-Step Workflow & Status

We will execute the refactoring in the following order to ensure the application remains functional at every step:

1.  **Migrate Utilities:** `[Completed]` Move the most independent, low-level functions from legacy files into the appropriate modules in `invoice_generator/utils/`. Update imports in the old files to point to the new locations.
2.  **Consolidate Data Logic:** `[Ready to Start]` Populate `data/data_preparer.py` with data transformation logic. Update the legacy functions to call this new module.
3.  **Build the Builders:** `[Not Started]` Create and populate the `TableBuilder` and `FooterBuilder` with logic from the old `fill_invoice_data` function.
4.  **Update the Processor:** `[Not Started]` Modify the `SingleTableProcessor` to use the new builders instead of calling the old functions.
5.  **Final Cleanup:** `[Not Started]` Once all logic has been migrated and legacy functions are no longer called, the old files (`invoice_utils.py`, etc.) will be deleted.