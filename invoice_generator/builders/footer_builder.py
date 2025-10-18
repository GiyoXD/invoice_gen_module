
from typing import Any, Dict, List, Optional, Tuple
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Alignment, Font, Side, Border
from openpyxl.utils import get_column_letter

from ..utils.layout import unmerge_row

from ..styling.models import StylingConfigModel

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
        grand_total_flag: bool = False,
        sheet_styling_config: Optional[StylingConfigModel] = None
    ):
        self.worksheet = worksheet
        self.footer_row_num = footer_row_num
        self.header_info = header_info
        self.sum_ranges = sum_ranges
        self.footer_config = footer_config
        self.pallet_count = pallet_count
        self.override_total_text = override_total_text
        self.DAF_mode = DAF_mode
        self.grand_total_flag = grand_total_flag
        self.sheet_styling_config = sheet_styling_config

    def build(self) -> int:
        if not self.footer_config or self.footer_row_num <= 0:
            return -1

        try:
            num_columns = self.header_info.get('num_columns', 1)
            column_map_by_id = self.header_info.get('column_id_map', {})

            style_config = self.footer_config.get('style', {})
            font_config = style_config.get('font', {'bold': True})
            align_config = style_config.get('alignment', {'horizontal': 'center', 'vertical': 'center'})
            border_config = style_config.get('border', {'apply': True})
            
            number_format_config = self.footer_config.get("number_formats", {})

            font_to_apply = Font(**font_config)
            align_to_apply = Alignment(**align_config)
            border_to_apply = None
            if border_config.get('apply'):
                side = Side(border_style=border_config.get('style', 'thin'), color=border_config.get('color', '000000'))
                border_to_apply = Border(left=side, right=side, top=side, bottom=side)

            unmerge_row(self.worksheet, self.footer_row_num, num_columns)

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
                cell = self.worksheet.cell(row=self.footer_row_num, column=total_text_col_idx, value=total_text)
                cell.font = font_to_apply
                cell.alignment = align_to_apply

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
                cell = self.worksheet.cell(row=self.footer_row_num, column=pallet_col_idx, value=pallet_text)
                cell.font = font_to_apply
                cell.alignment = align_to_apply

            sum_column_ids = self.footer_config.get("sum_column_ids", [])
            if self.sum_ranges:
                for col_id in sum_column_ids:
                    col_idx = column_map_by_id.get(col_id)
                    if col_idx:
                        col_letter = get_column_letter(col_idx)
                        sum_parts = [f"{col_letter}{start}:{col_letter}{end}" for start, end in self.sum_ranges]
                        formula = f"=SUM({','.join(sum_parts)})"
                        cell = self.worksheet.cell(row=self.footer_row_num, column=col_idx, value=formula)
                        cell.font = font_to_apply
                        cell.alignment = align_to_apply
                        
                        number_format_str = number_format_config.get(col_id)
                        if number_format_str and self.DAF_mode and col_id not in ['col_pcs', 'col_qty_pcs']:
                            cell.number_format = "##,00.00"
                        elif number_format_str:
                            cell.number_format = number_format_str["number_format"]
            
            if not self.grand_total_flag:
                for c_idx in range(1, num_columns + 1):
                    cell = self.worksheet.cell(row=self.footer_row_num, column=c_idx)
                    if cell.font != font_to_apply: cell.font = font_to_apply
                    if cell.alignment != align_to_apply: cell.alignment = align_to_apply
                    if border_to_apply:
                        cell.border = border_to_apply

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
                    self.worksheet.merge_cells(start_row=self.footer_row_num, start_column=resolved_start_col, end_row=self.footer_row_num, end_column=end_col)
            
            if self.sheet_styling_config:
                idx_to_id_map = {v: k for k, v in column_map_by_id.items()}
                for c_idx in range(1, num_columns + 1):
                    cell = self.worksheet.cell(row=self.footer_row_num, column=c_idx)
                    column_id = idx_to_id_map.get(c_idx)
                    if column_id and column_id in self.sheet_styling_config.columnIdStyles:
                        col_style = self.sheet_styling_config.columnIdStyles[column_id]
                        if col_style.alignment:
                            cell.alignment = Alignment(**col_style.alignment.model_dump(exclude_none=True))
                        if col_style.font:
                            cell.font = Font(**col_style.font.model_dump(exclude_none=True))

            return self.footer_row_num

        except Exception as e:
            print(f"ERROR: An error occurred during footer generation on row {self.footer_row_num}: {e}")
            return -1
