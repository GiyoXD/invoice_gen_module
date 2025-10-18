import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from typing import List, Dict, Any, Tuple

class TemplateStateBuilder:
    """
    A builder responsible for capturing and restoring the state of a template file.
    This includes the header, footer, and other static content.
    """

    def __init__(self, worksheet: Worksheet, num_header_cols: int):
        self.worksheet = worksheet
        self.header_state: List[List[Dict[str, Any]]] = []
        self.footer_state: List[List[Dict[str, Any]]] = []
        self.header_merged_cells: List[str] = []
        self.footer_merged_cells: List[str] = []
        self.row_heights: Dict[int, float] = {}
        self.column_widths: Dict[int, float] = {}
        self.template_footer_start_row: int = -1
        self.min_row = 1
        self.max_row = self.worksheet.max_row
        self.min_col = 1
        self.num_header_cols = num_header_cols
        self.max_col = min(self.worksheet.max_column, self.num_header_cols) if self.num_header_cols > 0 else self.worksheet.max_column

    def _get_cell_info(self, worksheet, row, col) -> Dict[str, Any]:
        cell = worksheet.cell(row=row, column=col)
        top_left_cell = cell
        for merged_cell_range in worksheet.merged_cells.ranges:
            if cell.coordinate in merged_cell_range:
                top_left_cell = worksheet.cell(row=merged_cell_range.min_row, column=merged_cell_range.min_col)
                break

        return {
            'value': cell.value,
            'font': top_left_cell.font.copy() if top_left_cell.font else None,
            'fill': top_left_cell.fill.copy() if top_left_cell.fill else None,
            'border': top_left_cell.border.copy() if top_left_cell.border else None,
            'alignment': top_left_cell.alignment.copy() if top_left_cell.alignment else None,
            'number_format': top_left_cell.number_format,
        }

    def capture_header(self, end_row: int):
        """
        Captures the state of the header section.
        """
        # Determine the actual start row of the header by finding the first row with content
        header_start_row = 1
        for r_idx in range(1, end_row + 1):
            if any(self.worksheet.cell(row=r_idx, column=c_idx).value is not None
                   for c_idx in range(1, self.max_col + 1)):
                header_start_row = r_idx
                break

        for r_idx in range(header_start_row, end_row + 1):
            row_data = []
            for c_idx in range(1, self.max_col + 1):
                row_data.append(self._get_cell_info(self.worksheet, r_idx, c_idx))
            self.header_state.append(row_data)
            self.row_heights[r_idx] = self.worksheet.row_dimensions[r_idx].height

        # Capture merged cells within the header range
        for merged_cell_range in self.worksheet.merged_cells.ranges:
            min_col, min_row, max_col, max_row = merged_cell_range.bounds
            if header_start_row <= min_row <= end_row and header_start_row <= max_row <= end_row:
                self.header_merged_cells.append(str(merged_cell_range))

        # Capture column widths
        for c_idx in range(1, self.max_col + 1):
            self.column_widths[c_idx] = self.worksheet.column_dimensions[get_column_letter(c_idx)].width

    def capture_footer(self, data_end_row: int):
        """
        Captures the state of the footer section.
        The footer is assumed to start after the data_end_row.
        """
        footer_start_row = data_end_row + 1
        # Find the actual first row with content after data_end_row
        for r_idx in range(data_end_row + 1, self.worksheet.max_row + 1):
            if any(self.worksheet.cell(row=r_idx, column=c_idx).value is not None
                   for c_idx in range(1, self.max_col + 1)):
                footer_start_row = r_idx
                break
        else:
            # No content found after data_end_row, so no footer to capture
            return

        self.template_footer_start_row = footer_start_row

        footer_end_row = self.worksheet.max_row
        for r_idx in range(self.worksheet.max_row, footer_start_row - 1, -1):
            if any(self.worksheet.cell(row=r_idx, column=c_idx).value is not None
                   for c_idx in range(1, self.max_col + 1)):
                footer_end_row = r_idx
                break

        for r_idx in range(footer_start_row, footer_end_row + 1):
            row_data = []
            for c_idx in range(1, self.max_col + 1):
                row_data.append(self._get_cell_info(self.worksheet, r_idx, c_idx))
            self.footer_state.append(row_data)
            self.row_heights[r_idx] = self.worksheet.row_dimensions[r_idx].height
            self.row_heights[r_idx] = self.worksheet.row_dimensions[r_idx].height

        # Capture merged cells within the footer range
        for merged_cell_range in self.worksheet.merged_cells.ranges:
            min_col, min_row, max_col, max_row = merged_cell_range.bounds
            if footer_start_row <= min_row <= footer_end_row and footer_start_row <= max_row <= footer_end_row:
                self.footer_merged_cells.append(str(merged_cell_range))

        # Capture column widths
        for c_idx in range(1, self.max_col + 1):
            self.column_widths[c_idx] = self.worksheet.column_dimensions[get_column_letter(c_idx)].width

    def restore_state(self, target_worksheet: Worksheet, data_start_row: int):
        """
        Restores the captured state to a new worksheet.
        """
        # Restore header merged cells without offset
        for merged_cell_range_str in self.header_merged_cells:
            target_worksheet.merge_cells(merged_cell_range_str)

        # Calculate the offset for footer rows and merged cells
        footer_start_row_in_new_sheet = target_worksheet.max_row + 1
        offset = footer_start_row_in_new_sheet - self.template_footer_start_row if self.template_footer_start_row != -1 else 0

        # Restore footer merged cells with offset
        for merged_cell_range_str in self.footer_merged_cells:
            from openpyxl.utils.cell import range_boundaries
            min_col, min_row, max_col, max_row = range_boundaries(merged_cell_range_str)
            
            # Adjust row numbers for all footer merged cells
            min_row += offset
            max_row += offset
            adjusted_range_str = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
            target_worksheet.merge_cells(adjusted_range_str)

        # Restore header
        current_row = 1
        for row_data in self.header_state:
            for c_idx, cell_info in enumerate(row_data, 1):
                cell = target_worksheet.cell(row=current_row, column=c_idx)
                if isinstance(cell, openpyxl.cell.cell.MergedCell):
                    continue
                cell.value = cell_info['value']
                if cell_info['font']: cell.font = cell_info['font']
                if cell_info['fill']: cell.fill = cell_info['fill']
                if cell_info['border']: cell.border = cell_info['border']
                if cell_info['alignment']: cell.alignment = cell_info['alignment']
                if cell_info['number_format']: cell.number_format = cell_info['number_format']
            target_worksheet.row_dimensions[current_row].height = self.row_heights.get(current_row, None)
            current_row += 1

        # Restore footer (adjust row numbers)
        for r_offset, row_data in enumerate(self.footer_state):
            r_idx = footer_start_row_in_new_sheet + r_offset
            for c_idx, cell_info in enumerate(row_data, 1):
                cell = target_worksheet.cell(row=r_idx, column=c_idx)
                if isinstance(cell, openpyxl.cell.cell.MergedCell):
                    continue
                cell.value = cell_info['value']
                if cell_info['font']: cell.font = cell_info['font']
                if cell_info['fill']: cell.fill = cell_info['fill']
                if cell_info['border']: cell.border = cell_info['border']
                if cell_info['alignment']: cell.alignment = cell_info['alignment']
                if cell_info['number_format']: cell.number_format = cell_info['number_format']
            original_footer_row_idx = self.max_row - len(self.footer_state) + r_offset + 1
            target_worksheet.row_dimensions[r_idx].height = self.row_heights.get(original_footer_row_idx, None)

        # Restore column widths
        for c_idx, width in self.column_widths.items():
            target_worksheet.column_dimensions[get_column_letter(c_idx)].width = width