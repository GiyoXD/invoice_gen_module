# style_utils.py
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Alignment, Border, Side, Font
from typing import Dict, Any, Optional, List, Tuple

def apply_cell_style(cell: Worksheet.cell, styling_config: dict, context: dict):
    """
    Applies all styles to a single cell, including fonts, alignments,
    and complex conditional borders, based on its context.
    """
    # --- Get Context ---
    col_id = context.get("col_id")
    col_idx = context.get("col_idx")
    static_col_idx = context.get("static_col_idx")
    is_pre_footer = context.get("is_pre_footer", False)

    # --- 1. Apply Font, Alignment, and Number Formats ---
    if col_id and styling_config:
        default_font_cfg = styling_config.get("default_font", {})
        default_align_cfg = styling_config.get("default_alignment", {})
        column_styles = styling_config.get("column_id_styles", {})
        col_specific_style = column_styles.get(col_id, {})
        
        final_font_cfg = {**default_font_cfg, **col_specific_style.get("font", {})}
        if final_font_cfg: cell.font = Font(**final_font_cfg)
        
        final_align_cfg = {**default_align_cfg, **col_specific_style.get("alignment", {})}
        if final_align_cfg: cell.alignment = Alignment(**final_align_cfg)
        
        if "number_format" in col_specific_style:
            cell.number_format = col_specific_style["number_format"]

    # --- 2. Apply Conditional Borders ---
    thin_side = Side(border_style="thin", color="000000")
    
    # Special handling for the pre-footer row
    if is_pre_footer:
        if col_idx == static_col_idx:
            cell.border = Border(left=thin_side, right=thin_side)
        else:
            cell.border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        return

    # UPDATED: Simplified logic for main data rows
    if col_idx == static_col_idx:
        # The static column ONLY ever gets side borders.
        cell.border = Border(left=thin_side, right=thin_side)
    elif col_idx: 
        # All other columns get a full grid.
        cell.border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)


def apply_row_heights(worksheet: Worksheet, styling_config: dict, headers: List[dict], data_ranges: List[Tuple[int, int]], footer_rows: List[int]):
    """
    Applies row heights for all headers, data rows, and footers.
    """
    print("  Applying all row heights...")
    row_heights_cfg = styling_config.get("row_heights", {})
    
    if h := row_heights_cfg.get('header'):
        for header_info in headers:
            for r in range(header_info['first_row_index'], header_info['second_row_index'] + 1):
                worksheet.row_dimensions[r].height = h
    
    if h := row_heights_cfg.get('data_default'):
        for start, end in data_ranges:
            for r in range(start, end + 1):
                worksheet.row_dimensions[r].height = h

    if h := row_heights_cfg.get('footer'):
        for r_num in footer_rows:
            worksheet.row_dimensions[r_num].height = h