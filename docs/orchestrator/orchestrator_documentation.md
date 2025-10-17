# Orchestrator Logic: `generate_invoice.py`

This document breaks down the step-by-step process of the main orchestrator script, `generate_invoice.py`. The script is responsible for setting up the environment, loading data, and coordinating the invoice generation process using a processor-based pattern.

## Step 1: Initialization and Argument Parsing

The script begins by preparing for execution and reading user input from the command line.

- **Purpose**: To configure the generation task based on user-provided paths and flags.
- **Process**:
    1.  An `ArgumentParser` is created to define the expected command-line arguments.
    2.  **Arguments defined**:
        - `input_data_file`: (Required) The path to the primary data source (`.json` or `.pkl`).
        - `-o`, `--output`: The desired path for the final generated Excel file.
        - `-t`, `--templatedir`: The directory where template `.xlsx` files are stored.
        - `-c`, `--configdir`: The directory where configuration `.json` files are stored.
        - `--DAF`: A boolean flag to trigger a special text replacement task.
        - `--custom`: A boolean flag for potential custom logic.
    3.  `parser.parse_args()` is called to read the arguments provided by the user.

## Step 2: Asset Loading (Paths, Config, Data)

This phase is responsible for locating and loading all necessary files and data into memory. The script will exit if any of these critical assets cannot be found.

- **Purpose**: To validate and load all required inputs before processing begins.
- **Process**:
    1.  **`derive_paths()`**:
        - Takes the input data file path and uses its name to intelligently find the corresponding template and configuration files.
        - It tries an exact name match first, then attempts a prefix-based match (e.g., `JLFHM_..._data.json` matches `JLFHM.xlsx` and `JLFHM_config.json`).
        - Returns a dictionary of `Path` objects for the data, template, and config files.
    2.  **`load_config()`**:
        - Reads the JSON configuration file identified by `derive_paths()`.
        - This file contains all the rules for processing, such as which sheets to process, how data is mapped, and styling information.
    3.  **`load_data()`**:
        - Reads the primary data file (`.json` or `.pkl`) identified by `derive_paths()`.
        - It includes a special sub-process to correctly parse string representations of `Decimal` objects into a usable format.

## Step 3: Template Preparation

The script prepares the output file by creating a safe working copy of the template.

- **Purpose**: To ensure the original template file is never modified.
- **Process**:
    1.  The output directory is created if it does not already exist.
    2.  `shutil.copy()` is used to copy the template `.xlsx` file (found in Step 2) to the final output path specified by the user.
    3.  All subsequent operations are performed on this copied file.

## Step 4: Main Processing and Workbook Manipulation

This is the core of the orchestrator, where the Excel file is actively modified. It is wrapped in a `try...except...finally` block to ensure the workbook is properly handled and closed, even if errors occur.

- **Purpose**: To systematically apply data and formatting to the workbook, sheet by sheet.
- **Process**:
    1.  **Load Workbook**: The copied Excel file is loaded into an `openpyxl` workbook object.
    2.  **Sheet Selection**: The `sheets_to_process` list from the configuration is filtered to ensure only sheets that actually exist in the workbook are processed.
    3.  **Pre-Processing**:
        - **Text Replacement**: Global text placeholders (e.g., invoice numbers, dates) are replaced throughout the workbook using `run_invoice_header_replacement_task`.
        - **DAF Task**: If the `--DAF` flag was used, a DAF-specific replacement task is also run.
    4.  **State & Config Mapping**:
        - `store_original_merges()` is called to create a backup of all merged cell ranges. This is critical for restoring them later.
        - The `sheet_data_map` and `data_mapping` configurations are retrieved, which will be used to direct the processing of each sheet.
    5.  **Global Calculation**: A `final_grand_total_pallets` value is calculated by summing pallet counts from the input data. This pre-calculated global value is then available to all sheet processors.
    6.  **Core Processing Loop**: The script iterates through each sheet designated for processing.
        - **Processor Factory**: Inside the loop, a "factory" pattern is used to decide which processor to use for the current sheet.
            - If the sheet's `data_source_indicator` is `processed_tables_multi`, it instantiates `MultiTableProcessor`.
            - Otherwise, it defaults to instantiating `SingleTableProcessor`.
        - **Execute Processing**: The `process()` method of the selected processor instance is called. This is where the main work for that sheet (writing data, inserting rows, etc.) happens.
        - **Error Handling**: If any processor's `process()` method returns `False`, the loop is immediately terminated.
    7.  **Merge Restoration**: After the loop, `find_and_restore_merges_heuristic()` is called to re-apply the merged cell configurations that were saved earlier. This fixes any merges that were broken by row insertions.

## Step 5: Save and Finalize

The final step is to save the modified workbook and report the outcome.

- **Purpose**: To persist the changes to disk and provide feedback to the user.
- **Process**:
    1.  If all processors completed successfully, the workbook is saved to the output path.
    2.  If an error occurred, the script still attempts to save the (potentially incomplete) workbook for debugging purposes.
    3.  The `workbook.close()` method is called within a `finally` block to release the file handle.
    4.  Total execution time is calculated and printed to the console.
