import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from typing import List, Dict, Any, Tuple
import copy

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
        self.template_footer_end_row: int = -1 # Added to store the end row of the footer
        self.min_row = 1
        self.max_row = self.worksheet.max_row
        self.min_col = 1
        self.num_header_cols = num_header_cols

        # Store default style objects for comparison
        default_workbook = openpyxl.Workbook()
        default_cell = default_workbook.active['A1']
        self.default_font = default_cell.font
        self.default_fill = default_cell.fill
        self.default_border = default_cell.border
        self.default_alignment = default_cell.alignment
        default_workbook.close() # Close the dummy workbook

        # Calculate max_col based on the maximum column with content in the entire worksheet
        max_col_with_content = 0
        max_row_with_content = 0 # Initialize max_row_with_content
        for r_idx in range(1, self.worksheet.max_row + 1):
            for c_idx in range(1, self.worksheet.max_column + 1):
                cell = self.worksheet.cell(row=r_idx, column=c_idx)
                if self._has_content_or_style(cell):
                    max_col_with_content = max(max_col_with_content, c_idx)
                    max_row_with_content = max(max_row_with_content, r_idx) # Update max_row_with_content
        self.max_col = max(max_col_with_content, self.num_header_cols) # Ensure it's at least num_header_cols
        self.max_row = max(max_row_with_content, self.max_row) # Update self.max_row with max_row_with_content

    def _has_content_or_style(self, cell) -> bool:
        if cell.value is not None and cell.value != '':
            return True
        # Check if any style is applied (not default)
        if cell.font and not self._is_default_style(cell.font, self.default_font): return True
        if cell.fill and not self._is_default_style(cell.fill, self.default_fill): return True
        if cell.border and not self._is_default_style(cell.border, self.default_border): return True
        if cell.alignment and not self._is_default_style(cell.alignment, self.default_alignment): return True
        return False

    def _is_default_style(self, style_obj, default_obj) -> bool:
        if style_obj is None:
            return True
        if default_obj is None: # Should not happen if default_obj is properly initialized
            return False
        
        # Compare relevant attributes for each style type
        if isinstance(style_obj, Font):
            return (
                style_obj.name == default_obj.name and
                style_obj.size == default_obj.size and
                style_obj.bold == default_obj.bold and
                style_obj.italic == default_obj.italic and
                style_obj.underline == default_obj.underline and
                style_obj.strike == default_obj.strike and
                style_obj.color == default_obj.color
            )
        elif isinstance(style_obj, PatternFill):
            return (
                style_obj.fill_type == default_obj.fill_type and
                style_obj.start_color == default_obj.start_color and
                style_obj.end_color == default_obj.end_color
            )
        elif isinstance(style_obj, Border):
            return (
                style_obj.left == default_obj.left and
                style_obj.right == default_obj.right and
                style_obj.top == default_obj.top and
                style_obj.bottom == default_obj.bottom and
                style_obj.diagonal == default_obj.diagonal
            )
        elif isinstance(style_obj, Alignment):
            return (
                style_obj.horizontal == default_obj.horizontal and
                style_obj.vertical == default_obj.vertical and
                style_obj.text_rotation == default_obj.text_rotation and
                style_obj.wrap_text == default_obj.wrap_text and
                style_obj.shrink_to_fit == default_obj.shrink_to_fit and
                style_obj.indent == default_obj.indent
            )
        
        return False # If type not recognized, assume not default

    def _get_cell_info(self, worksheet, row, col) -> Dict[str, Any]:
        cell = worksheet.cell(row=row, column=col)
        top_left_cell = cell
        for merged_cell_range in worksheet.merged_cells.ranges:
            if cell.coordinate in merged_cell_range:
                top_left_cell = worksheet.cell(row=merged_cell_range.min_row, column=merged_cell_range.min_col)
                break

        return {
            'value': cell.value,
            'font': copy.copy(top_left_cell.font) if top_left_cell.font and not self._is_default_style(top_left_cell.font, self.default_font) else None,
            'fill': copy.copy(top_left_cell.fill) if top_left_cell.fill and not self._is_default_style(top_left_cell.fill, self.default_fill) else None,
            'border': copy.copy(top_left_cell.border) if top_left_cell.border and not self._is_default_style(top_left_cell.border, self.default_border) else None,
            'alignment': copy.copy(top_left_cell.alignment) if top_left_cell.alignment and not self._is_default_style(top_left_cell.alignment, self.default_alignment) else None,
            'number_format': top_left_cell.number_format,
        }

    def capture_header(self, end_row: int):
        """
        Captures the state of the header section.
        """
        # Determine the actual start row of the header by finding the first row with content
        header_start_row = 1
        for r_idx in range(1, end_row + 1):
            if any(self._has_content_or_style(self.worksheet.cell(row=r_idx, column=c_idx))
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

    def capture_footer(self, data_end_row: int, max_possible_footer_row: int):
        """
        Captures the state of the footer section.
        The footer is assumed to start after the data_end_row.
        """
        footer_start_row = data_end_row + 1
        # Find the actual first row with content after data_end_row
        for r_idx in range(data_end_row + 1, max_possible_footer_row + 1):
            if any(self._has_content_or_style(self.worksheet.cell(row=r_idx, column=c_idx))
                   for c_idx in range(1, self.max_col + 1)):
                footer_start_row = r_idx
                break
        else:
            # No content found after data_end_row, so no footer to capture
            return

        self.template_footer_start_row = footer_start_row

        # Find the true max row with content within the expected footer region
        true_max_row_with_content_in_footer = 0
        for r_idx in range(max_possible_footer_row, footer_start_row - 1, -1): # Iterate from max_possible_footer_row down to footer_start_row
            if any(self._has_content_or_style(self.worksheet.cell(row=r_idx, column=c_idx))
                   for c_idx in range(1, self.max_col + 1)):
                true_max_row_with_content_in_footer = r_idx
                break

        footer_end_row = true_max_row_with_content_in_footer
        self.template_footer_end_row = footer_end_row # Store the end row of the footer

        for r_idx in range(footer_start_row, footer_end_row + 1):
            row_data = []
            for c_idx in range(1, self.max_col + 1):
                row_data.append(self._get_cell_info(self.worksheet, r_idx, c_idx))
            self.footer_state.append(row_data)
            self.row_heights[r_idx] = self.worksheet.row_dimensions[r_idx].height

        # Capture merged cells within the footer range
        for merged_cell_range in self.worksheet.merged_cells.ranges:
            min_col, min_row, max_col, max_row = merged_cell_range.bounds
            if footer_start_row <= min_row <= footer_end_row and footer_start_row <= max_row <= footer_end_row:
                self.footer_merged_cells.append(str(merged_cell_range))

        # Capture column widths
        for c_idx in range(1, self.max_col + 1):
            self.column_widths[c_idx] = self.worksheet.column_dimensions[get_column_letter(c_idx)].width

    def restore_state(self, target_worksheet: Worksheet, data_start_row: int, data_table_end_row: int):
        """
        Restores the captured state to a new worksheet.
        """
        # Restore header merged cells without offset
        for merged_cell_range_str in self.header_merged_cells:
            target_worksheet.merge_cells(merged_cell_range_str)

        # Calculate the offset for footer rows and merged cells
        # Accurately find the last row with content in target_worksheet
        # Use data_table_end_row to determine where the footer should start
        footer_start_row_in_new_sheet = data_table_end_row + 1
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
                
                # Apply value and styles to the top-left cell
                cell.value = cell_info['value']
                if cell_info['font']: cell.font = cell_info['font']
                if cell_info['fill']: cell.fill = cell_info['fill']
                if cell_info['border']: cell.border = cell_info['border']
                if cell_info['alignment']: cell.alignment = cell_info['alignment']
                if cell_info['number_format']: cell.number_format = cell_info['number_format']
            
            # Fix row height lookup
            original_footer_row_idx = self.template_footer_start_row + r_offset
            target_worksheet.row_dimensions[r_idx].height = self.row_heights.get(original_footer_row_idx, None)

        # Restore column widths
        for c_idx, width in self.column_widths.items():
            target_worksheet.column_dimensions[get_column_letter(c_idx)].width = width
