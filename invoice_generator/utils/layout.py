from openpyxl.worksheet.worksheet import Worksheet
from typing import List, Dict, Any, Optional, Tuple
from openpyxl.utils import get_column_letter

def unmerge_row(worksheet: Worksheet, row_num: int, num_cols: int):
    """
    Unmerges any merged cells that overlap with the specified row within the given column range.

    Args:
        worksheet: The openpyxl Worksheet object.
        row_num: The 1-based row index to unmerge.
        num_cols: The number of columns to check for merges.
    """
    if row_num <= 0:
        return
    merged_ranges_copy = list(worksheet.merged_cells.ranges) # Copy ranges before modification
    merged_ranges_to_remove = []

    # Identify ranges that overlap with the target row
    for merged_range in merged_ranges_copy:
        # Check if the range's row span includes the target row_num
        # And if the range's column span overlaps with columns 1 to num_cols
        overlap = (merged_range.min_row <= row_num <= merged_range.max_row and
                   max(merged_range.min_col, 1) <= min(merged_range.max_col, num_cols))
        if overlap:
            merged_ranges_to_remove.append(str(merged_range))

    if merged_ranges_to_remove:
        for range_str in merged_ranges_to_remove:
            try:
                worksheet.unmerge_cells(range_str)
            except KeyError:
                # Range might have been removed by unmerging an overlapping one
                pass
            except Exception as unmerge_err:
                # Log or handle other potential errors if needed
                pass
    else:
        # No overlapping merges found for this row
        pass


def unmerge_block(worksheet: Worksheet, start_row: int, end_row: int, num_cols: int):
    """
    Unmerges any merged cells that overlap with the specified row range and column range.
    Args:
        worksheet: The openpyxl Worksheet object.
        start_row: The 1-based starting row index of the block.
        end_row: The 1-based ending row index of the block.
        num_cols: The number of columns to check for merges.
    """
    if start_row <= 0 or end_row < start_row:
        return
    merged_ranges_copy = list(worksheet.merged_cells.ranges) # Copy ranges before modification
    merged_ranges_to_remove = []

    # Identify ranges that overlap with the target block
    for merged_range in merged_ranges_copy:
        mr_min_row, mr_min_col, mr_max_row, mr_max_col = merged_range.bounds
        row_overlap = max(mr_min_row, start_row) <= min(mr_max_row, end_row)
        col_overlap = max(mr_min_col, 1) <= min(mr_max_col, num_cols)

        if row_overlap and col_overlap:
            range_str = str(merged_range)
            if range_str not in merged_ranges_to_remove: # Avoid duplicates
                merged_ranges_to_remove.append(range_str)

    if merged_ranges_to_remove:
        for range_str in merged_ranges_to_remove:
            try:
                worksheet.unmerge_cells(range_str)
            except KeyError:
                # Range might have been removed by unmerging an overlapping one
                pass
            except Exception as unmerge_err:
                # Log or handle other potential errors if needed
                pass
    else:
        # No overlapping merges found in this block
        pass


def safe_unmerge_block(worksheet: Worksheet, start_row: int, end_row: int, num_cols: int):
    """
    Safely unmerges only cells within the specific target range, preventing unintended unmerging
    of cells completely outside the block.
    """
    if start_row <= 0 or end_row < start_row:
        return

    # Only process merges that actually intersect with our target range
    for merged_range in list(worksheet.merged_cells.ranges):
        # Check if this merge intersects our target range
        if (merged_range.min_row <= end_row and
            merged_range.max_row >= start_row and
            merged_range.min_col <= num_cols and
            merged_range.max_col >= 1):
            try:
                worksheet.unmerge_cells(merged_range.coord)
            except (KeyError, ValueError, AttributeError):
                # Ignore errors if the range is somehow invalid or already unmerged
                continue

    return True


from ..styling.models import StylingConfigModel

def apply_column_widths(worksheet: Worksheet, sheet_styling_config: Optional[StylingConfigModel], header_map: Optional[Dict[str, int]]):
    """
    Sets column widths based on the configuration.

    Args:
        worksheet: The openpyxl Worksheet object.
        sheet_styling_config: Styling configuration containing the 'column_widths' dictionary.
        header_map: Dictionary mapping header text to column index (1-based).
    """
    if not sheet_styling_config or not header_map: return
    column_widths_cfg = sheet_styling_config.column_id_widths
    if not column_widths_cfg or not isinstance(column_widths_cfg, dict): return
    for header_text, width in column_widths_cfg.items():
        col_idx = header_map.get(header_text)
        if col_idx:
            col_letter = get_column_letter(col_idx)
            try:
                width_val = float(width)
                if width_val > 0: worksheet.column_dimensions[col_letter].width = width_val
                else: pass # Ignore non-positive widths
            except (ValueError, TypeError): pass # Ignore invalid width values
            except Exception as width_err: pass # Log other errors?
        else: pass # Header text not found in map


def apply_row_heights(worksheet: Worksheet, sheet_styling_config: Optional[StylingConfigModel], header_info: Dict[str, Any], data_row_indices: List[int], footer_row_index: int, row_after_header_idx: int, row_before_footer_idx: int):
    """
    Sets row heights based on the configuration for header, data, footer, and specific rows.
    Footer height can now optionally match the header height.

    Args:
        worksheet: The openpyxl Worksheet object.
        sheet_styling_config: Styling configuration containing the 'row_heights' dictionary.
        header_info: Dictionary with header row indices.
        data_row_indices: List of 1-based indices for the actual data rows written.
        footer_row_index: 1-based index of the footer row.
        row_after_header_idx: 1-based index of the static/blank row after the header (-1 if none).
        row_before_footer_idx: 1-based index of the static/blank row before the footer (-1 if none).
    """
    if not sheet_styling_config: return
    row_heights_cfg = sheet_styling_config.row_heights
    if not row_heights_cfg or not isinstance(row_heights_cfg, dict): return

    actual_header_height = None # Store the applied header height

    def set_height(r_idx, height_val, desc): # Helper function
        nonlocal actual_header_height # Ensure actual_header_height is modified
        if r_idx <= 0: return
        try:
            h_val = float(height_val)
            if h_val > 0:
                worksheet.row_dimensions[r_idx].height = h_val
                if desc == "header": # Store the height applied to the header
                    actual_header_height = h_val
            else: pass # Ignore non-positive heights
        except (ValueError, TypeError): pass # Ignore invalid height values
        except Exception as height_err: pass # Log other errors?

    # Apply Heights Based on Config
    header_height = row_heights_cfg.get("header")
    if header_height is not None and header_info:
        h_start = header_info.get('first_row_index', -1); h_end = header_info.get('second_row_index', -1)
        if h_start > 0 and h_end >= h_start:
            for r in range(h_start, h_end + 1): set_height(r, header_height, "header")

    after_header_height = row_heights_cfg.get("after_header")
    if after_header_height is not None and row_after_header_idx > 0: set_height(row_after_header_idx, after_header_height, "after_header")
    data_default_height = row_heights_cfg.get("data_default")
    if data_default_height is not None and data_row_indices:
        for r in data_row_indices: set_height(r, data_default_height, "data_default")
    before_footer_height = row_heights_cfg.get("before_footer")
    if before_footer_height is not None and row_before_footer_idx > 0: set_height(row_before_footer_idx, before_footer_height, "before_footer")

    # --- Footer Height Logic ---
    footer_height_config = row_heights_cfg.get("footer")
    match_header_height_flag = row_heights_cfg.get("footer_matches_header_height", True) # Default to True

    final_footer_height = None
    if match_header_height_flag and actual_header_height is not None:
        final_footer_height = actual_header_height # Use header height if flag is true and header height was set
    elif footer_height_config is not None:
        final_footer_height = footer_height_config # Otherwise, use specific footer height if defined

    if final_footer_height is not None and footer_row_index > 0:
        set_height(footer_row_index, final_footer_height, "footer")
    # --- End Footer Height Logic ---

    specific_heights = row_heights_cfg.get("specific_rows")
    if isinstance(specific_heights, dict):
        for row_str, height_val in specific_heights.items():
            try: row_num = int(row_str); set_height(row_num, height_val, f"specific_row_{row_num}")
            except ValueError: pass # Ignore invalid row numbers

def calculate_header_dimensions(header_layout: List[Dict[str, Any]]) -> Tuple[int, int]:
    """
    Calculates the total number of rows and columns a header will occupy.
    """
    if not header_layout:
        return (0, 0)
    num_rows = max(cell.get('row', 0) + cell.get('rowspan', 1) for cell in header_layout)
    num_cols = max(cell.get('col', 0) + cell.get('colspan', 1) for cell in header_layout)
    return (num_rows, num_cols)
