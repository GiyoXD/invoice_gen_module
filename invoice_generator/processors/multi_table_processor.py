from typing import Any, Dict, List, Tuple
from openpyxl.worksheet.worksheet import Worksheet

from ..utils.writing import write_header
from ..builders.footer_builder import FooterBuilder
from ..utils.layout import apply_row_heights, unmerge_block, safe_unmerge_block
from ..utils import merge_utils
from ..styling import style_applier as style_utils
from .base_processor import BaseProcessor

class MultiTableProcessor(BaseProcessor):
    def process(self) -> bool:
        write_pointer_row = self.sheet_config.get("start_row", 1)
        all_data_ranges: List[Tuple[int, int]] = []
        all_header_infos: List[Dict] = []
        all_footer_rows: List[int] = []
        grand_total_pallets = 0

        header_to_write = self.sheet_config.get("header_to_write", [])
        footer_config = self.sheet_config.get("footer_configurations", {})
        styling_config = self.styling_config
        mappings = self.sheet_config.get("mappings", {})
        data_map = mappings.get("data_map", {})
        static_col_values = mappings.get("initial_static", {}).get("values", [])
        
        raw_data = self.invoice_data.get('processed_tables_data', {})
        table_keys = sorted(raw_data.keys())
        num_tables = len(table_keys)
        last_data_end_row = -1

        for i, table_key in enumerate(table_keys):
            table_data = raw_data[table_key]
            num_data_rows = len(table_data.get('net', []))
            
            is_last_table = (i == num_tables - 1) # Calculate is_last_table

            header_info = write_header(self.worksheet, write_pointer_row, header_to_write, self.styling_config)
            all_header_infos.append(header_info)
            write_pointer_row = header_info.get('second_row_index', write_pointer_row) + 1
            col_map = header_info.get('column_id_map', {})
            num_columns = header_info.get('num_columns', 1)
            static_col_idx = col_map.get(mappings.get("initial_static", {}).get("column_header_id"))
            idx_to_id_map = {v: k for k, v in col_map.items()}

            data_start_row = write_pointer_row
            
            keys_to_convert_to_numeric = {'net', 'amount', 'price'}

            data_end_row = data_start_row + num_data_rows - 1
            safe_unmerge_block(self.worksheet, data_start_row, data_end_row, num_columns)

            for r_idx in range(num_data_rows):
                current_row = write_pointer_row + r_idx
                if static_col_idx and r_idx < len(static_col_values):
                    self.worksheet.cell(row=current_row, column=static_col_idx).value = static_col_values[r_idx]
                
                for data_key, mapping_info in data_map.items():
                    if col_idx := col_map.get(mapping_info.get("id")):
                        if data_key in table_data:
                            value = table_data[data_key][r_idx]
                            
                            if data_key in keys_to_convert_to_numeric and isinstance(value, str):
                                try:
                                    numeric_value = float(value.replace(',', ''))
                                    self.worksheet.cell(row=current_row, column=col_idx).value = numeric_value
                                except (ValueError, TypeError):
                                    self.worksheet.cell(row=current_row, column=col_idx).value = value
                            else:
                                self.worksheet.cell(row=current_row, column=col_idx).value = value

                for c_idx in range(1, num_columns + 1):
                    cell = self.worksheet.cell(row=current_row, column=c_idx)
                    style_context = {
                        "col_id": idx_to_id_map.get(c_idx), "col_idx": c_idx,
                        "static_col_idx": static_col_idx, "row_index": r_idx,
                        "num_data_rows": num_data_rows
                    }
                    style_utils.apply_cell_style(cell, self.styling_config, style_context)
            
            write_pointer_row += num_data_rows
            all_data_ranges.append((data_start_row, write_pointer_row - 1))

            data_end_row = write_pointer_row - 1
            last_data_end_row = data_end_row
            vertical_merge_ids = mappings.get("vertical_merge_on_id", [])
            if vertical_merge_ids:
                print(f"Applying vertical merges for table '{table_key}'...")
                for col_id_to_merge in vertical_merge_ids:
                    if col_idx := col_map.get(col_id_to_merge):
                        merge_utils.merge_vertical_cells_in_range(
                            worksheet=self.worksheet,
                            scan_col=col_idx,
                            start_row=data_start_row,
                            end_row=data_end_row
                        )

            pre_footer_config = footer_config.get("pre_footer_row")
            if pre_footer_config and isinstance(pre_footer_config, dict):
                cells_to_write = pre_footer_config.get("cells", [])
                for cell_data in cells_to_write:
                    value = cell_data.get("value")
                    if col_idx := col_map.get(col_id):
                        self.worksheet.cell(row=write_pointer_row, column=col_idx).value = value
                
                for c_idx in range(1, num_columns + 1):
                    cell = self.worksheet.cell(row=write_pointer_row, column=c_idx)
                    col_id = idx_to_id_map.get(c_idx)
                    style_context = {
                        "col_id": col_id, "col_idx": c_idx,
                        "static_col_idx": static_col_idx, "is_pre_footer": True
                    }
                    style_utils.apply_cell_style(cell, self.styling_config, style_context)
                
                pre_footer_merges = pre_footer_config.get("merge_rules")
                merge_utils.apply_row_merges(self.worksheet, write_pointer_row, num_columns, pre_footer_merges)
                
                if data_row_height := self.styling_config.row_heights.get("data_default"):
                    self.worksheet.row_dimensions[write_pointer_row].height = data_row_height
                
                write_pointer_row += 1

            pallet_count = len(table_data.get('pallet_count', []))
            grand_total_pallets += pallet_count
            
            print(f"DEBUG: MultiTableProcessor - self.args.DAF: {self.args.DAF}") # NEW DEBUG PRINT
            footer_builder = FooterBuilder(
                worksheet=self.worksheet,
                footer_row_num=write_pointer_row,
                header_info=header_info,
                sum_ranges=[(data_start_row, write_pointer_row - 1)],
                footer_config={**footer_config, "type": "regular"},
                pallet_count=pallet_count,
                sheet_styling_config=self.styling_config,
                all_tables_data=self.invoice_data.get('processed_tables_data', {}),
                table_keys=table_keys,
                mapping_rules=mappings,
                sheet_name=self.sheet_name,
                DAF_mode=self.args.DAF,
                is_last_table=is_last_table,
            )
            write_pointer_row = footer_builder.build()
            
            all_footer_rows.append(write_pointer_row)
            # write_pointer_row += 1 # This is now handled by footer_builder.build()

            if i < num_tables - 1:
                write_pointer_row += 1

        if num_tables > 1:
            last_header_info = all_header_infos[-1]
            num_columns = last_header_info.get('num_columns', 1)
            
            # For the grand total footer, it is always the last table in the context of the sheet
            is_last_table_grand_total = True 

            grand_total_footer_builder = FooterBuilder(
                worksheet=self.worksheet,
                footer_row_num=write_pointer_row,
                header_info=last_header_info,
                sum_ranges=all_data_ranges,
                footer_config={**footer_config, "type": "grand_total", "add_ons": ["summary"] if self.args.DAF and self.sheet_name == "Packing list" else []},
                pallet_count=grand_total_pallets,
                override_total_text="TOTAL OF:",
                sheet_styling_config=self.styling_config,
                all_tables_data=self.invoice_data.get('processed_tables_data', {}),
                table_keys=table_keys,
                mapping_rules=mappings,
                sheet_name=self.sheet_name,
                DAF_mode=self.args.DAF,
                is_last_table=is_last_table_grand_total,
            )
            write_pointer_row = grand_total_footer_builder.build()
            
            all_footer_rows.append(write_pointer_row)

        style_utils.apply_row_heights(self.worksheet, self.styling_config, all_header_infos, all_data_ranges, all_footer_rows)
        
        self.run_text_replacement()

        return write_pointer_row, last_data_end_row
