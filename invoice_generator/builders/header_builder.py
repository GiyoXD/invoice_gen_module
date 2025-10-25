from typing import Any, Dict, List, Optional
from openpyxl.worksheet.worksheet import Worksheet

from ..styling.models import StylingConfigModel
from ..styling.style_applier import apply_header_style, apply_cell_style
from ..utils.layout import unmerge_block, calculate_header_dimensions
from openpyxl.utils import get_column_letter

class HeaderBuilder:
    def __init__(
        self,
        worksheet: Worksheet,
        start_row: int,
        header_layout_config: List[Dict[str, Any]],
        sheet_styling_config: Optional[StylingConfigModel] = None,
    ):
        self.worksheet = worksheet
        self.start_row = start_row
        self.header_layout_config = header_layout_config
        self.sheet_styling_config = sheet_styling_config

    def build(self) -> Optional[Dict[str, Any]]:
        if not self.header_layout_config or self.start_row <= 0:
            return None

        num_header_rows, num_header_cols = calculate_header_dimensions(self.header_layout_config)
        if num_header_rows > 0 and num_header_cols > 0:
            unmerge_block(self.worksheet, self.start_row, self.start_row + num_header_rows - 1, num_header_cols)

        first_row_index = self.start_row
        last_row_index = self.start_row
        max_col = 0
        column_map = {}
        column_id_map = {}

        for cell_config in self.header_layout_config:
            row_offset = cell_config.get('row', 0)
            col_offset = cell_config.get('col', 0)
            text = cell_config.get('text', '')
            cell_id = cell_config.get('id')
            rowspan = cell_config.get('rowspan', 1)
            colspan = cell_config.get('colspan', 1)

            cell_row = self.start_row + row_offset
            cell_col = 1 + col_offset

            last_row_index = max(last_row_index, cell_row + rowspan - 1)
            max_col = max(max_col, cell_col + colspan - 1)

            cell = self.worksheet.cell(row=cell_row, column=cell_col, value=text)
            apply_header_style(cell, self.sheet_styling_config)
            
            context = {
                "col_id": cell_id,
                "col_idx": cell_col,
                "is_header": True
            }
            apply_cell_style(cell, self.sheet_styling_config, context)

            if cell_id:
                column_map[text] = get_column_letter(cell_col)
                column_id_map[cell_id] = cell_col

            if rowspan > 1 or colspan > 1:
                self.worksheet.merge_cells(start_row=cell_row, start_column=cell_col,
                                      end_row=cell_row + rowspan - 1, end_column=cell_col + colspan - 1)

        return {
            'first_row_index': first_row_index,
            'second_row_index': last_row_index,
            'column_map': column_map,
            'column_id_map': column_id_map,
            'num_columns': max_col
        }
