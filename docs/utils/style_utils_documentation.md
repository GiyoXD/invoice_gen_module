# Utility Module Documentation: `style_utils.py`

This document provides a detailed breakdown of the `style_utils.py` module. This module is intended to centralize styling logic, but as of now, it contains a mix of generic and highly specific logic.

---

## `style_utils.py`

This module contains functions responsible for applying visual styles—such as fonts, alignments, borders, and row heights—to cells and rows within a worksheet.

### Core Responsibilities:
- Applying a complete set of styles to a single cell based on its context.
- Setting the height for various rows (header, data, footer).

### Function Breakdown:

- **`apply_cell_style(cell, styling_config, context)`**:
    - **Purpose**: To be a single, comprehensive function that applies all necessary styles to a given cell.
    - **Logic**:
        1.  It receives a `cell` object, the main `styling_config` from the JSON, and a `context` dictionary.
        2.  The `context` dictionary is crucial; it provides information about the cell's location and role (e.g., its `col_id`, `col_idx`, and whether it's in a `pre_footer` row).
        3.  **Font, Alignment, Number Format**: It first applies the base styles. It looks up the `col_id` in the `column_id_styles` section of the configuration and applies the specified font, alignment, and number format. It also uses default styles from the config if no column-specific style is found.
        4.  **Conditional Borders**: This is the most complex part. It applies different border styles based on the `context`:
            - If the cell is in a "pre-footer" row, it applies a full border to most cells but only side borders to the designated "static column."
            - For regular data rows, it applies only side borders to the "static column" and a full grid border to all other columns.
    - **Analysis**: This function encapsulates a significant amount of business logic related to layout and styling, making it a key function for understanding how the final document is styled.

- **`apply_row_heights(worksheet, styling_config, headers, data_ranges, footer_rows)`**:
    - **Purpose**: To set the height for all the different types of rows in the document in one pass.
    - **Logic**:
        1.  It reads the `row_heights` dictionary from the `styling_config`.
        2.  **Header Height**: It gets the `header` height and applies it to all rows that are part of any header block.
        3.  **Data Height**: It gets the `data_default` height and applies it to all rows that fall within the specified data ranges.
        4.  **Footer Height**: It gets the `footer` height and applies it to all specified footer rows.
    - **Analysis**: This is a straightforward utility function that cleanly separates the logic for applying row heights.
