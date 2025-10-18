# invoice_generator/styling/style_applier.py
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Alignment, Border, Side, Font
from typing import Dict, Any, Optional, List, Tuple

# --- Style Constants ---
thin_side = Side(border_style="thin", color="000000")
thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
no_border = Border(left=None, right=None, top=None, bottom=None)
center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
left_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
bold_font = Font(bold=True)

from .models import StylingConfigModel

def apply_cell_style(cell: Worksheet.cell, styling_config: StylingConfigModel, context: dict):
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
        col_specific_style = styling_config.columnIdStyles.get(col_id)
        
        if col_specific_style:
            if col_specific_style.font:
                cell.font = Font(**col_specific_style.font.model_dump(exclude_none=True))
            elif styling_config.defaultFont:
                cell.font = Font(**styling_config.defaultFont.model_dump(exclude_none=True))

            if col_specific_style.alignment:
                cell.alignment = Alignment(**col_specific_style.alignment.model_dump(exclude_none=True))
            elif styling_config.defaultAlignment:
                cell.alignment = Alignment(**styling_config.defaultAlignment.model_dump(exclude_none=True))

            if col_specific_style.numberFormat:
                cell.number_format = col_specific_style.numberFormat

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


def apply_row_heights(worksheet: Worksheet, styling_config: StylingConfigModel, headers: List[dict], data_ranges: List[Tuple[int, int]], footer_rows: List[int]):
    """
    Applies row heights for all headers, data rows, and footers.
    """
    print("  Applying all row heights...")
    
    if styling_config.rowHeights:
        if h := styling_config.rowHeights.get('header'):
            for header_info in headers:
                for r in range(header_info['first_row_index'], header_info['second_row_index'] + 1):
                    worksheet.row_dimensions[r].height = h
        
        if h := styling_config.rowHeights.get('data_default'):
            for start, end in data_ranges:
                for r in range(start, end + 1):
                    worksheet.row_dimensions[r].height = h

        if h := styling_config.rowHeights.get('footer'):
            for r_num in footer_rows:
                worksheet.row_dimensions[r_num].height = h
