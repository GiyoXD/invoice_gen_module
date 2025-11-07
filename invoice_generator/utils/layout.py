from openpyxl.worksheet.worksheet import Worksheet
from typing import List, Dict, Any, Optional, Tuple
from openpyxl.utils import get_column_letter
from ..styling.models import StylingConfigModel
from ..styling.style_applier import apply_cell_style, apply_header_style
from ..styling.style_config import THIN_BORDER, NO_BORDER, CENTER_ALIGNMENT, LEFT_ALIGNMENT, BOLD_FONT
from decimal import Decimal, InvalidOperation
import re
import traceback
import logging



def apply_column_widths(worksheet: Worksheet, sheet_styling_config: Optional[StylingConfigModel], header_map: Optional[Dict[str, int]]):
    """
    Sets column widths based on the configuration.

    Args:
        worksheet: The openpyxl Worksheet object.
        sheet_styling_config: Styling configuration containing the 'column_widths' dictionary.
        header_map: Dictionary mapping header text to column index (1-based).
    """
    if not sheet_styling_config or not header_map: return
    column_widths_cfg = sheet_styling_config.columnIdWidths
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

def calculate_header_dimensions(header_layout: List[Dict[str, Any]]) -> Tuple[int, int]:
    """
    Calculates the total number of rows and columns a header will occupy.
    """
    if not header_layout:
        return (0, 0)
    num_rows = max(cell.get('row', 0) + cell.get('rowspan', 1) for cell in header_layout)
    num_cols = max(cell.get('col', 0) + cell.get('colspan', 1) for cell in header_layout)
    return (num_rows, num_cols)

def fill_static_row(worksheet: Worksheet, row_num: int, num_cols: int, static_content_dict: Dict[str, Any], sheet_styling_config: Optional[StylingConfigModel] = None):
    """
    Fills a specific row with static content defined in a dictionary.
    Delegates styling to the apply_cell_style function.

    Args:
        worksheet: The openpyxl Worksheet object.
        row_num: The 1-based row index to fill.
        num_cols: The total number of columns in the table context (for bounds checking).
        static_content_dict: Dictionary where keys are column indices (as strings or ints)
                             and values are the static content to write.
        sheet_styling_config: The styling configuration for the sheet.
    """
    if not static_content_dict:
        return  # Nothing to do
    if row_num <= 0:
        return

    for col_key, value in static_content_dict.items():
        target_col_index = None
        try:
            # Attempt to convert key to integer column index
            target_col_index = int(col_key)
            # Check if the column index is within the valid range
            if 1 <= target_col_index <= num_cols:
                cell = worksheet.cell(row=row_num, column=target_col_index)
                cell.value = value
                
                # Delegate styling to apply_cell_style
                context = {
                    "col_idx": target_col_index,
                    "is_static_row": True 
                }
                apply_cell_style(cell, sheet_styling_config, context)

            else:
                # Column index out of range, log warning?
                pass
        except (ValueError, TypeError) as e:
            # Invalid column key, log warning?
            pass
        except Exception as cell_err:
            # Error accessing cell, log warning?
            pass

def apply_row_merges(worksheet: Worksheet, row_num: int, num_cols: int, merge_rules: Optional[Dict[str, int]]):
    """
    Applies horizontal merges to a specific row based on rules.

    Args:
        worksheet: The openpyxl Worksheet object.
        row_num: The 1-based row index to apply merges to.
        num_cols: The total number of columns in the table context.
        merge_rules: Dictionary where keys are starting column indices (as strings or ints)
                     and values are the number of columns to span (colspan).
    """
    if not merge_rules or row_num <= 0:
        return # No rules or invalid row

    try:
        # Convert string keys to integers and sort for predictable application order
        rules_with_int_keys = {int(k): v for k, v in merge_rules.items()}
        sorted_keys = sorted(rules_with_int_keys.keys())
    except (ValueError, TypeError) as e:
        # Invalid key format in merge_rules
        return

    for start_col in sorted_keys:
        colspan_val = rules_with_int_keys[start_col]
        try:
            # Ensure colspan is an integer
            colspan = int(colspan_val)
        except (ValueError, TypeError):
            # Invalid colspan value
            continue

        # Basic validation for start column and colspan
        if not isinstance(start_col, int) or not isinstance(colspan, int) or start_col < 1 or colspan < 1:
            continue

        # Calculate end column, ensuring it doesn't exceed the table width
        end_col = start_col + colspan - 1
        if end_col > num_cols:
            end_col = num_cols
            # Check if clamping made the range invalid (start > end)
            if start_col > end_col:
                continue

        try:
            worksheet.merge_cells(start_row=row_num, start_column=start_col, end_row=row_num, end_column=end_col)
            # Apply alignment to the top-left cell of the merged range
            top_left_cell = worksheet.cell(row=row_num, column=start_col)
            if not top_left_cell.alignment or top_left_cell.alignment.horizontal is None:
                top_left_cell.alignment = CENTER_ALIGNMENT # Apply center alignment if none exists
        except ValueError as ve:
            # This can happen if trying to merge over an existing merged cell that wasn't properly unmerged
            pass
        except Exception as merge_err:
            # Log or handle other merge errors
            pass

def merge_contiguous_cells_by_id(
    worksheet: Worksheet,
    start_row: int,
    end_row: int,
    col_id_to_merge: str,
    column_id_map: Dict[str, int]
):
    """
    Finds and merges contiguous vertical cells within a column that have the same value.
    This is called AFTER all data has been written to the sheet.
    """
    col_idx = column_id_map.get(col_id_to_merge)
    if not col_idx or start_row >= end_row:
        return

    current_merge_start_row = start_row
    value_to_match = worksheet.cell(row=start_row, column=col_idx).value

    for row_idx in range(start_row + 1, end_row + 2):
        cell_value = worksheet.cell(row=row_idx, column=col_idx).value if row_idx <= end_row else object()
        if cell_value != value_to_match:
            if row_idx - 1 > current_merge_start_row:
                if value_to_match is not None and str(value_to_match).strip():
                    try:
                        worksheet.merge_cells(
                            start_row=current_merge_start_row,
                            start_column=col_idx,
                            end_row=row_idx - 1,
                            end_column=col_idx
                        )
                    except Exception as e:
                        print(f"Could not merge cells for ID {col_id_to_merge} from row {current_merge_start_row} to {row_idx - 1}. Error: {e}")
            
            current_merge_start_row = row_idx
            if row_idx <= end_row:
                value_to_match = cell_value

def apply_explicit_data_cell_merges_by_id(
    worksheet: Worksheet,
    row_num: int,
    column_id_map: Dict[str, int],  # Maps column ID to its 1-based column index
    num_total_columns: int,
    merge_rules_data_cells: Dict[str, Dict[str, Any]], # e.g., {'col_item': {'rowspan': 2}}
    sheet_styling_config: Optional[StylingConfigModel] = None,
    DAF_mode: Optional[bool] = False
):
    """
    Applies horizontal merges to data cells in a specific row based on column IDs.
    """
    if not merge_rules_data_cells or row_num <= 0:
        return

    # Loop through rules where the key is now the column ID
    for col_id, rule_details in merge_rules_data_cells.items():
        colspan_to_apply = rule_details.get("rowspan")

        if not isinstance(colspan_to_apply, int) or colspan_to_apply <= 1:
            continue
        
        # Get column index from the ID map
        start_col_idx = column_id_map.get(col_id)
        if not start_col_idx:
            print(f"Warning: Could not find column for merge rule with ID '{col_id}'.")
            continue
            
        end_col_idx = start_col_idx + colspan_to_apply - 1
        end_col_idx = min(end_col_idx, num_total_columns)

        if start_col_idx >= end_col_idx:
            continue

        try:
            # Apply the new merge
            worksheet.merge_cells(start_row=row_num, start_column=start_col_idx,
                                  end_row=row_num, end_column=end_col_idx)
            
            # Style the anchor cell of the new merged range
            anchor_cell = worksheet.cell(row=row_num, column=start_col_idx)
            context = {
                "col_id": col_id,
                "col_idx": start_col_idx,
                "DAF_mode": DAF_mode
            }
            apply_cell_style(anchor_cell, sheet_styling_config, context)

        except Exception as e:
            print(f"Error applying explicit data cell merge for ID '{col_id}' on row {row_num}: {e}")

def write_summary_rows(
    worksheet: Worksheet,
    start_row: int,
    header_info: Dict[str, Any],
    all_tables_data: Dict[str, Any],
    table_keys: List[str],
    footer_config: Dict[str, Any],
    mapping_rules: Dict[str, Any],
    styling_config: Optional[StylingConfigModel] = None,
    DAF_mode: Optional[bool] = False,
    grand_total_pallets: int = 0
) -> int:
    """
    Calculates and writes ID-driven summary rows for different leather types.
    This function is specifically designed to handle "BUFFALO LEATHER" and "COW LEATHER"
    summaries, calculating totals for specified numeric columns and pallet counts.
    It applies styling based on the provided styling configuration.
    """
    buffalo_summary_row = start_row
    leather_summary_row = start_row + 1
    next_available_row = start_row + 2

    try:
        # --- Configuration and Data Extraction ---
        column_id_map = header_info.get('column_id_map', {})
        idx_to_id_map = {v: k for k, v in column_id_map.items()}
        data_map = mapping_rules.get('data_map', {})
        
        # Define which column IDs should be summed up
        numeric_ids_to_sum = ["col_qty_pcs", "col_qty_sf", "col_net", "col_gross", "col_cbm"]
        id_to_data_key_map = {v['id']: k for k, v in data_map.items() if v.get('id') in numeric_ids_to_sum}
        ids_to_sum = list(id_to_data_key_map.keys())

        # --- Totals Initialization ---
        grand_totals = {col_id: 0 for col_id in ids_to_sum}
        buffalo_totals = {col_id: 0 for col_id in ids_to_sum}
        cow_totals = {col_id: 0 for col_id in ids_to_sum}
        grand_pallet_total = grand_total_pallets # Initialize with the passed parameter
        buffalo_pallet_total = 0
        cow_pallet_total = 0

        for table_key in table_keys:
            table_data = all_tables_data.get(str(table_key), {})
            descriptions = table_data.get("description", [])
            pallet_counts = table_data.get("pallet_count", [])

            for i in range(len(descriptions)):
                raw_val = descriptions[i]
                desc_val = str(raw_val)
                is_buffalo = "BUFFALO" in desc_val.upper()

                logging.debug(f"DEBUG: desc_val: {desc_val}, is_buffalo: {is_buffalo}")

                try:
                    pallet_val = int(pallet_counts[i]) if i < len(pallet_counts) else 0
                    # REMOVED: grand_pallet_total += pallet_val
                    if is_buffalo:
                        buffalo_pallet_total += pallet_val
                    else: # If not buffalo, it's cow leather
                        cow_pallet_total += pallet_val
                    logging.debug(f"DEBUG: Pallet - grand: {grand_pallet_total}, buffalo: {buffalo_pallet_total}, cow: {cow_pallet_total}")
                except (ValueError, TypeError):
                    logging.debug(f"DEBUG: Error processing pallet_val for desc_val: {desc_val}")
                    pass

                for col_id in ids_to_sum:
                    data_key = id_to_data_key_map.get(col_id)
                    if not data_key: continue
                    
                    data_list = table_data.get(data_key, [])
                    if i < len(data_list):
                        try:
                            value_to_add = data_list[i]
                            numeric_value = 0
                            if isinstance(value_to_add, (int, float)):
                                numeric_value = float(value_to_add)
                            elif isinstance(value_to_add, str) and value_to_add.strip():
                                numeric_value = float(value_to_add.replace(',', ''))
                            
                            grand_totals[col_id] += numeric_value
                            if is_buffalo:
                                buffalo_totals[col_id] += numeric_value
                            else: # If not buffalo, it's cow leather
                                cow_totals[col_id] += numeric_value
                            logging.debug(f"DEBUG: {col_id} - grand: {grand_totals[col_id]}, buffalo: {buffalo_totals[col_id]}, cow: {cow_totals[col_id]}") # Changed to logging.debug
                        except (ValueError, TypeError, IndexError):
                            logging.debug(f"DEBUG: Error processing numeric_value for {col_id} and desc_val: {desc_val}") # Changed to logging.debug
                            pass

        # --- Writing to Worksheet ---
        num_columns = header_info['num_columns']
        desc_col_idx = column_id_map.get("col_desc")
        label_col_idx = column_id_map.get("col_pallet") or 2

        # Write Buffalo Summary Row
        cell = worksheet.cell(row=buffalo_summary_row, column=label_col_idx, value="TOTAL OF:")
        apply_cell_style(cell, styling_config, {"col_id": "col_pallet", "col_idx": label_col_idx, "is_footer": True})
        cell = worksheet.cell(row=buffalo_summary_row, column=label_col_idx + 1, value="BUFFALO LEATHER")
        apply_cell_style(cell, styling_config, {"col_id": "col_pallet", "col_idx": label_col_idx + 1, "is_footer": True})
        if desc_col_idx:
            cell = worksheet.cell(row=buffalo_summary_row, column=desc_col_idx, value=f"{buffalo_pallet_total} PALLETS")
            apply_cell_style(cell, styling_config, {"col_id": "col_desc", "col_idx": desc_col_idx, "is_footer": True})
        for col_id, total_value in buffalo_totals.items():
            col_idx = column_id_map.get(col_id)
            if col_idx:
                cell = worksheet.cell(row=buffalo_summary_row, column=col_idx, value=total_value)
                apply_cell_style(cell, styling_config, {"col_id": col_id, "col_idx": col_idx, "is_footer": True})

        # Write Cow Leather Summary Row
        cell = worksheet.cell(row=leather_summary_row, column=label_col_idx, value="TOTAL OF:")
        apply_cell_style(cell, styling_config, {"col_id": "col_pallet", "col_idx": label_col_idx, "is_footer": True})
        cell = worksheet.cell(row=leather_summary_row, column=label_col_idx + 1, value="COW LEATHER")
        apply_cell_style(cell, styling_config, {"col_id": "col_pallet", "col_idx": label_col_idx + 1, "is_footer": True})
        if desc_col_idx:
            cell = worksheet.cell(row=leather_summary_row, column=desc_col_idx, value=f"{cow_pallet_total} PALLETS")
            apply_cell_style(cell, styling_config, {"col_id": "col_desc", "col_idx": desc_col_idx, "is_footer": True})
        for col_id, total_value in cow_totals.items():
            col_idx = column_id_map.get(col_id)
            if col_idx:
                cell = worksheet.cell(row=leather_summary_row, column=col_idx, value=total_value)
                apply_cell_style(cell, styling_config, {"col_id": col_id, "col_idx": col_idx, "is_footer": True})

        return next_available_row

    except Exception as summary_err:
        logging.error(f"Warning: Failed processing summary rows: {summary_err}") # Changed to logging.error
        traceback.print_exc()
        return start_row + 2
