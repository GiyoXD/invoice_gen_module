from typing import Any, Dict, List, Optional, Tuple, Union
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell
import traceback

from invoice_generator.data.data_preparer import prepare_data_rows, parse_mapping_rules
from invoice_generator.utils.layout import apply_column_widths
from invoice_generator.styling.style_applier import apply_row_heights
from invoice_generator.utils.layout import fill_static_row, apply_row_merges, merge_contiguous_cells_by_id, apply_explicit_data_cell_merges_by_id
from invoice_generator.styling.style_applier import apply_cell_style
# FooterBuilder is now called by LayoutBuilder (proper Director pattern)
from invoice_generator.styling.style_config import THIN_BORDER, NO_BORDER, CENTER_ALIGNMENT, LEFT_ALIGNMENT, BOLD_FONT


# --- Constants for Number Formats ---
FORMAT_GENERAL = 'General'
FORMAT_TEXT = '@'
FORMAT_NUMBER_COMMA_SEPARATED1 = '#,##0'
FORMAT_NUMBER_COMMA_SEPARATED2 = '#,##0.00'

from invoice_generator.styling.models import StylingConfigModel
from .bundle_accessor import BundleAccessor

class DataTableBuilderStyler:
    """
    Builds and styles data table sections based on pre-resolved data.
    
    This class is a "dumb" builder. Its only job is to take prepared data
    and write it to the worksheet. It does not contain any data-sourcing
    or mapping logic.
    """
    
    def __init__(
        self,
        worksheet: Worksheet,
        header_info: Dict[str, Any],
        resolved_data: Dict[str, Any],
        sheet_styling_config: Optional[StylingConfigModel] = None
    ):
        """
        Initialize the builder with resolved data.
        
        Args:
            worksheet: The worksheet to write to.
            header_info: Header information with column maps.
            resolved_data: The data prepared by TableDataResolver.
            sheet_styling_config: The styling configuration for the sheet.
        """
        self.worksheet = worksheet
        self.header_info = header_info
        self.resolved_data = resolved_data
        self.sheet_styling_config = sheet_styling_config

        # Extract commonly used values
        self.data_rows = resolved_data.get('data_rows', [])
        self.static_info = resolved_data.get('static_info', {})
        self.formula_rules = resolved_data.get('formula_rules', {})
        self.pallet_counts = resolved_data.get('pallet_counts', [])
        self.dynamic_desc_used = resolved_data.get('dynamic_desc_used', False)
        
        self.col_id_map = header_info.get('column_id_map', {})
        self.idx_to_id_map = {v: k for k, v in self.col_id_map.items()}

    def build(self) -> Tuple[bool, int, int, int, int]:
        if not self.header_info or 'second_row_index' not in self.header_info:
            print("Error: Invalid header_info provided to DataTableBuilderStyler.")
            return False, -1, -1, -1, 0

        num_columns = self.header_info.get('num_columns', 0)
        data_writing_start_row = self.header_info.get('second_row_index', 0) + 1
        
        actual_rows_to_process = len(self.data_rows)
        
        data_start_row = data_writing_start_row
        data_end_row = data_start_row + actual_rows_to_process - 1 if actual_rows_to_process > 0 else data_start_row - 1
        
        footer_row_final = data_end_row + 1

        # --- Fill Data Rows Loop ---
        try:
            data_row_indices_written = []
            for i in range(actual_rows_to_process):
                current_row_idx = data_start_row + i
                data_row_indices_written.append(current_row_idx)
                
                row_data = self.data_rows[i]

                # Write data
                for col_idx, value in row_data.items():
                    cell = self.worksheet.cell(row=current_row_idx, column=col_idx)
                    if not isinstance(cell, MergedCell):
                        # Check if value is a formula dict
                        if isinstance(value, dict) and value.get('type') == 'formula':
                            # Convert formula dict to Excel formula string
                            formula_str = self._build_formula_string(value, current_row_idx)
                            cell.value = formula_str
                        else:
                            cell.value = value
                    # apply_cell_style(...) # Styling logic can be added here

            # --- Merging and other logic can be added here if needed ---

        except Exception as fill_data_err:
            print(f"Error during data filling loop: {fill_data_err}\n{traceback.format_exc()}")
            return False, -1, -1, -1, 0

        local_chunk_pallets = sum(int(p) for p in self.pallet_counts if p is not None and str(p).isdigit())

        return True, footer_row_final, data_start_row, data_end_row, local_chunk_pallets
    
    def _build_formula_string(self, formula_dict: Dict[str, Any], row_num: int) -> str:
        """
        Convert a formula dict to an Excel formula string.
        
        Args:
            formula_dict: Dict with 'template' and 'inputs' keys
            row_num: Current row number
        
        Returns:
            Excel formula string (e.g., "=B5*C5")
        """
        template = formula_dict.get('template', '')
        inputs = formula_dict.get('inputs', [])
        
        # Replace placeholders like {col_ref_0}, {col_ref_1}, etc.
        formula = template
        for i, input_id in enumerate(inputs):
            col_idx = self.col_id_map.get(input_id)
            if col_idx:
                col_letter = get_column_letter(col_idx)
                formula = formula.replace(f'{{col_ref_{i}}}', col_letter)
        
        # Replace {row} with actual row number
        formula = formula.replace('{row}', str(row_num))
        
        # Ensure formula starts with =
        if not formula.startswith('='):
            formula = '=' + formula
        
        return formula

