# Utility Module Documentation: `merge_utils.py`

This document provides a detailed breakdown of the `merge_utils.py` module. This module centralizes all logic related to cell merging and unmerging, providing a range of utilities from forceful cleaning to heuristic-based restoration.

---

## `merge_utils.py`

This module is critical for managing the complex state of merged cells in `openpyxl`, which can be easily corrupted by row insertions.

### Core Responsibilities:
- Storing the state of merged cells before modification.
- Restoring merged cells after data has been written.
- Providing utilities for targeted and bulk unmerging.
- Applying new merges based on rules or data.

### Function Breakdown:

- **`store_original_merges(workbook, sheet_names)`**:
    - **Purpose**: To create a backup of horizontal merge configurations from specified sheets before any modifications are made.
    - **Logic**:
        1. It iterates through the specified `sheet_names`.
        2. It scans for all merged cell ranges on each sheet.
        3. **Crucially, it filters out any merges that start above row 16**, focusing only on the data area. It also ignores merges that span multiple rows.
        4. For each valid horizontal merge, it stores a tuple containing its column span (`colspan`), the value of its top-left cell, and the height of its row.
    - **Analysis**: This function is the first step in the merge-restoration workflow. By saving the state (without coordinates), it provides the necessary information for the `find_and_restore_merges_heuristic` function to work later.

- **`find_and_restore_merges_heuristic(workbook, stored_merges, processed_sheet_names, search_range_str)`**:
    - **Purpose**: To intelligently re-apply merged cells after row insertions have likely broken them.
    - **Logic**:
        1. It iterates through the `stored_merges` data for each sheet.
        2. For each stored merge (value, span, and height), it searches for the `value` in the worksheet, starting from the bottom of the `search_range_str` and moving up.
        3. When it finds a cell with the matching value, it assumes this is the top-left anchor of the merge that needs to be restored.
        4. It proactively unmerges any existing cells in the target range to prevent conflicts.
        5. It then applies a new merge with the stored `colspan`, re-applies the stored `value`, and sets the row to the stored `height`.
    - **Analysis**: This is a "heuristic" because it relies on finding the unique value to locate the merge's position, rather than using original coordinates. The bottom-up search helps find the correct location in cases where the same value might appear elsewhere.

- **`force_unmerge_from_row_down(worksheet, start_row)`**:
    - **Purpose**: To provide a clean slate for writing data by removing all merges in a specific area of the worksheet.
    - **Logic**: It iterates through all merged ranges on the sheet and removes any merge that **starts on or after** the given `start_row`.
    - **Analysis**: This is a safer alternative to unmerging the entire sheet, as it leaves header merges (which are typically above the data area) intact.

- **`apply_horizontal_merge(worksheet, row_num, num_cols, merge_rules)`**:
    - **Purpose**: To apply a set of horizontal merges to a single row based on a configuration dictionary.
    - **Logic**: It takes a dictionary of `merge_rules` where the key is the starting column and the value is the number of columns to span (`colspan`). It then iterates through these rules and applies them to the specified `row_num`.
    - **Analysis**: A straightforward utility for applying declarative merge rules, often used for footers or static rows.

- **`merge_vertical_cells_in_range(worksheet, scan_col, start_row, end_row)`**:
    - **Purpose**: To scan a single column and merge adjacent cells that contain the exact same value.
    - **Logic**:
        1. It iterates from the `start_row` to the `end_row` within the specified `scan_col`.
        2. It keeps track of the current value it's looking for.
        3. When it finds a sequence of two or more identical, non-empty values, it merges them into a single vertical block.
    - **Analysis**: This is used for cleaning up data presentation after writing, such as merging repeated "Pallet No" or "Item" values to make the table easier to read.