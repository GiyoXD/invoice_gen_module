
from typing import Any, Dict, List, Optional, Tuple
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Alignment, Font, Side, Border
from openpyxl.utils import get_column_letter

from ..utils.layout import unmerge_row


from ..styling.models import StylingConfigModel

from ..styling.style_applier import apply_cell_style

class FooterBuilder:
    def __init__(
        self,
        worksheet: Worksheet,
        footer_row_num: int,
        header_info: Dict[str, Any],
        sum_ranges: List[Tuple[int, int]],
        footer_config: Dict[str, Any],
        pallet_count: int,
        override_total_text: Optional[str] = None,
        DAF_mode: bool = False,
        sheet_styling_config: Optional[StylingConfigModel] = None,
        all_tables_data: Optional[Dict[str, Any]] = None,
        table_keys: Optional[List[str]] = None,
        mapping_rules: Optional[Dict[str, Any]] = None,
        sheet_name: Optional[str] = None,
        is_last_table: bool = False,
        dynamic_desc_used: bool = False,
    ):
        self.worksheet = worksheet
        self.footer_row_num = footer_row_num
        self.header_info = header_info
        self.sum_ranges = sum_ranges
        self.footer_config = footer_config
        self.pallet_count = pallet_count
        self.override_total_text = override_total_text
        self.DAF_mode = DAF_mode
        self.sheet_styling_config = sheet_styling_config
        self.all_tables_data = all_tables_data
        self.table_keys = table_keys
        self.mapping_rules = mapping_rules
        self.sheet_name = sheet_name
        self.is_last_table = is_last_table
        self.dynamic_desc_used = dynamic_desc_used

    def _apply_footer_cell_style(self, cell, col_id):
        context = {
            "col_id": col_id,
            "col_idx": cell.column,
            "is_footer": True
        }
        apply_cell_style(cell, self.sheet_styling_config, context)

    def build(self) -> int:
        if not self.footer_config or self.footer_row_num <= 0:
            return -1

        try:
            current_footer_row = self.footer_row_num
            
            footer_type = self.footer_config.get("type", "regular")

            if footer_type == "regular":
                self._build_regular_footer(current_footer_row)
            elif footer_type == "grand_total":
                self._build_grand_total_footer(current_footer_row)

            current_footer_row += 1

            # Handle add-ons
            add_ons = self.footer_config.get("add_ons", [])
            if "summary" in add_ons:
                current_footer_row = self._build_summary_add_on(current_footer_row)

            from ..styling.style_applier import apply_row_heights
            apply_row_heights(
                worksheet=self.worksheet,
                sheet_styling_config=self.sheet_styling_config,
                footer_row_index=current_footer_row
            )

            return current_footer_row

        except Exception as e:
            print(f"ERROR: An error occurred during footer generation on row {self.footer_row_num}: {e}")
            return -1

    def _build_regular_footer(self, current_footer_row: int):
        num_columns = self.header_info.get('num_columns', 1)
        column_map_by_id = self.header_info.get('column_id_map', {})

        unmerge_row(self.worksheet, current_footer_row, num_columns)

        total_text = self.override_total_text if self.override_total_text is not None else self.footer_config.get("total_text", "TOTAL:")
        total_text_col_id = self.footer_config.get("total_text_column_id")
        
        total_text_col_idx = None
        if total_text_col_id is not None:
            if isinstance(total_text_col_id, int):
                total_text_col_idx = total_text_col_id + 1
            elif isinstance(total_text_col_id, str):
                try:
                    raw_index = int(total_text_col_id)
                    total_text_col_idx = raw_index + 1
                except ValueError:
                    total_text_col_idx = column_map_by_id.get(total_text_col_id)
        
        if total_text_col_idx:
            self.worksheet.cell(row=current_footer_row, column=total_text_col_idx, value=total_text)

        pallet_col_id = self.footer_config.get("pallet_count_column_id")
        
        pallet_col_idx = None
        if pallet_col_id is not None and self.pallet_count > 0:
            if isinstance(pallet_col_id, int):
                pallet_col_idx = pallet_col_id + 1
            elif isinstance(pallet_col_id, str):
                try:
                    raw_index = int(pallet_col_id)
                    pallet_col_idx = raw_index + 1
                except ValueError:
                    pallet_col_idx = column_map_by_id.get(pallet_col_id)
        
        if pallet_col_idx:
            pallet_text = f"{self.pallet_count} PALLET{'S' if self.pallet_count != 1 else ''}"
            self.worksheet.cell(row=current_footer_row, column=pallet_col_idx, value=pallet_text)

        sum_column_ids = self.footer_config.get("sum_column_ids", [])
        if self.sum_ranges:
            for col_id in sum_column_ids:
                col_idx = column_map_by_id.get(col_id)
                if col_idx:
                    col_letter = get_column_letter(col_idx)
                    sum_parts = [f"{col_letter}{start}:{col_letter}{end}" for start, end in self.sum_ranges]
                    formula = f"=SUM({','.join(sum_parts)})"
                    self.worksheet.cell(row=current_footer_row, column=col_idx, value=formula)
        
        idx_to_id_map = {v: k for k, v in column_map_by_id.items()}
        for c_idx in range(1, num_columns + 1):
            cell = self.worksheet.cell(row=current_footer_row, column=c_idx)
            col_id = idx_to_id_map.get(c_idx)
            self._apply_footer_cell_style(cell, col_id)

        merge_rules = self.footer_config.get("merge_rules", [])
        for rule in merge_rules:
            start_column_id = rule.get("start_column_id")
            colspan = rule.get("colspan")
            
            resolved_start_col = None
            
            if start_column_id is not None:
                if isinstance(start_column_id, int):
                    resolved_start_col = start_column_id + 1
                elif isinstance(start_column_id, str):
                    try:
                        raw_index = int(start_column_id)
                        resolved_start_col = raw_index + 1
                    except ValueError:
                        resolved_start_col = column_map_by_id.get(start_column_id)
            
            if resolved_start_col and colspan:
                end_col = min(resolved_start_col + colspan - 1, num_columns)
                self.worksheet.merge_cells(start_row=current_footer_row, start_column=resolved_start_col, end_row=current_footer_row, end_column=end_col)

    def _build_grand_total_footer(self, current_footer_row: int):
        num_columns = self.header_info.get('num_columns', 1)
        column_map_by_id = self.header_info.get('column_id_map', {})

        total_text = self.override_total_text if self.override_total_text is not None else "TOTAL OF:"
        total_text_col_id = self.footer_config.get("total_text_column_id")
        
        total_text_col_idx = None
        if total_text_col_id is not None:
            if isinstance(total_text_col_id, int):
                total_text_col_idx = total_text_col_id + 1
            elif isinstance(total_text_col_id, str):
                try:
                    raw_index = int(total_text_col_id)
                    total_text_col_idx = raw_index + 1
                except ValueError:
                    total_text_col_idx = column_map_by_id.get(total_text_col_id)
        
        if total_text_col_idx:
            self.worksheet.cell(row=current_footer_row, column=total_text_col_idx, value=total_text)

        pallet_col_id = self.footer_config.get("pallet_count_column_id")
        
        pallet_col_idx = None
        if pallet_col_id is not None and self.pallet_count > 0:
            if isinstance(pallet_col_id, int):
                pallet_col_idx = pallet_col_id + 1
            elif isinstance(pallet_col_id, str):
                try:
                    raw_index = int(pallet_col_id)
                    pallet_col_idx = raw_index + 1
                except ValueError:
                    pallet_col_idx = column_map_by_id.get(pallet_col_id)
        
        if pallet_col_idx:
            pallet_text = f"{self.pallet_count} PALLET{'S' if self.pallet_count != 1 else ''}"
            self.worksheet.cell(row=current_footer_row, column=pallet_col_idx, value=pallet_text)

        sum_column_ids = self.footer_config.get("sum_column_ids", [])
        if self.sum_ranges:
            for col_id in sum_column_ids:
                col_idx = column_map_by_id.get(col_id)
                if col_idx:
                    col_letter = get_column_letter(col_idx)
                    sum_parts = [f"{col_letter}{start}:{col_letter}{end}" for start, end in self.sum_ranges]
                    formula = f"=SUM({','.join(sum_parts)})"
                    self.worksheet.cell(row=current_footer_row, column=col_idx, value=formula)

        idx_to_id_map = {v: k for k, v in column_map_by_id.items()}
        for c_idx in range(1, num_columns + 1):
            cell = self.worksheet.cell(row=current_footer_row, column=c_idx)
            col_id = idx_to_id_map.get(c_idx)
            self._apply_footer_cell_style(cell, col_id)

        merge_rules = self.footer_config.get("merge_rules", [])
        for rule in merge_rules:
            start_column_id = rule.get("start_column_id")
            colspan = rule.get("colspan")
            
            resolved_start_col = None
            
            if start_column_id is not None:
                if isinstance(start_column_id, int):
                    resolved_start_col = start_column_id + 1
                elif isinstance(start_column_id, str):
                    try:
                        raw_index = int(start_column_id)
                        resolved_start_col = raw_index + 1
                    except ValueError:
                        resolved_start_col = column_map_by_id.get(start_column_id)
            
            if resolved_start_col and colspan:
                end_col = min(resolved_start_col + colspan - 1, num_columns)
                self.worksheet.merge_cells(start_row=current_footer_row, start_column=resolved_start_col, end_row=current_footer_row, end_column=end_col)

    def _build_summary_add_on(self, current_footer_row: int) -> int:
        from ..utils.writing import write_summary_rows # NEW IMPORT
        if self.DAF_mode and self.dynamic_desc_used and self.sheet_name == "Packing list" and self.is_last_table and self.all_tables_data and self.table_keys and self.mapping_rules:
            return write_summary_rows(
                worksheet=self.worksheet,
                start_row=current_footer_row,
                header_info=self.header_info,
                all_tables_data=self.all_tables_data,
                table_keys=self.table_keys,
                footer_config=self.footer_config,
                mapping_rules=self.mapping_rules,
                styling_config=self.sheet_styling_config,
                DAF_mode=self.DAF_mode,
                grand_total_pallets=self.pallet_count
            )
        return current_footer_row
