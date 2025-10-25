# invoice_generator/processors/multi_table_processor.py
import sys
from .base_processor import SheetProcessor
from .. import invoice_utils
import traceback
from openpyxl.utils import get_column_letter

class MultiTableProcessor(SheetProcessor):
    """
    Processes a worksheet that contains multiple, repeating blocks of tables,
    such as a packing list.
    """

    def process(self) -> bool:
        """
        Executes the logic for processing a multi-table sheet.
        """
        print(f"Processing sheet '{self.sheet_name}' as multi-table (write header mode).")
        
        all_tables_data = self.invoice_data.get('processed_tables_data', {})
        if not all_tables_data or not isinstance(all_tables_data, dict):
            print(f"Warning: 'processed_tables_data' not found/valid. Skipping '{self.sheet_name}'.")
            return True # Not a failure, just nothing to do

        header_to_write = self.sheet_config.get('header_to_write')
        start_row = self.sheet_config.get('start_row')
        if not start_row or not header_to_write:
            print(f"Error: Config for multi-table '{self.sheet_name}' missing 'start_row' or 'header_to_write'. Skipping.")
            return False

        table_keys = sorted(all_tables_data.keys(), key=lambda x: int(x) if str(x).isdigit() else float('inf'))
        print(f"Found table keys in data: {table_keys}")
        num_tables = len(table_keys)
        last_table_header_info = None

        # --- Pre-calculate and insert rows ---
        success, _ = self._pre_calculate_and_insert_rows(table_keys, all_tables_data, header_to_write)
        if not success:
            return False

        write_pointer_row = start_row
        grand_total_pallets_for_summary_row = 0
        all_data_ranges = []

        # --- Main loop to write data ---
        for i, table_key in enumerate(table_keys):
            print(f"\nProcessing table key: '{table_key}' ({i+1}/{num_tables})")
            table_data_to_fill = all_tables_data.get(str(table_key))
            if not table_data_to_fill or not isinstance(table_data_to_fill, dict):
                print(f"Warning: No/invalid data for table key '{table_key}'. Skipping.")
                continue

            # Write header for the current table
            written_header_info = invoice_utils.write_header(
                self.worksheet, write_pointer_row, header_to_write, self.sheet_config.get("styling")
            )
            if not written_header_info:
                print(f"Error writing header for table '{table_key}'. Skipping sheet.")
                return False
            last_table_header_info = written_header_info

            num_header_rows, _ = invoice_utils.calculate_header_dimensions(header_to_write)
            write_pointer_row += num_header_rows

            temp_header_info = written_header_info.copy()
            temp_header_info['first_row_index'] = write_pointer_row - num_header_rows
            temp_header_info['second_row_index'] = temp_header_info['first_row_index'] + 1

            fill_success, next_row_after_chunk, data_start, data_end, table_pallets = invoice_utils.fill_invoice_data(
                worksheet=self.worksheet,
                sheet_name=self.sheet_name,
                sheet_config=self.sheet_config,
                all_sheet_configs=self.data_mapping_config,
                data_source=table_data_to_fill,
                data_source_type='processed_tables',
                header_info=temp_header_info,
                mapping_rules=self.sheet_config.get('mappings', {}),
                sheet_styling_config=self.sheet_config.get("styling"),
                add_blank_after_header=self.sheet_config.get("add_blank_after_header", False),
                static_content_after_header=self.sheet_config.get("static_content_after_header", {}),
                add_blank_before_footer=self.sheet_config.get("add_blank_before_footer", False),
                static_content_before_footer=self.sheet_config.get("static_content_before_footer", {}),
                merge_rules_after_header=self.sheet_config.get("merge_rules_after_header", {}),
                merge_rules_before_footer=self.sheet_config.get("merge_rules_before_footer", {}),
                merge_rules_footer=self.sheet_config.get("merge_rules_footer", {}),
                footer_info=None, max_rows_to_fill=None,
                grand_total_pallets=self.final_grand_total_pallets,
                custom_flag=self.args.custom,
                data_cell_merging_rules=self.sheet_config.get("data_cell_merging_rule", None),
                DAF_mode=self.args.DAF,
            )

            if not fill_success:
                print(f"Error filling data/footer for table '{table_key}'. Stopping.")
                return False

            grand_total_pallets_for_summary_row += table_pallets
            if data_start > 0 and data_end >= data_start:
                all_data_ranges.append((data_start, data_end))

            write_pointer_row = next_row_after_chunk

            is_last_table = (i == num_tables - 1)
            if not is_last_table:
                # Handle spacer row
                num_cols_spacer = 1
                try:
                    invoice_utils.unmerge_row(self.worksheet, write_pointer_row, num_cols_spacer)
                    self.worksheet.merge_cells(start_row=write_pointer_row, start_column=1, end_row=write_pointer_row, end_column=num_cols_spacer)
                except Exception as merge_err:
                    print(f"Warning: Failed to write/merge spacer row {write_pointer_row}: {merge_err}")
                write_pointer_row += 1
        
        # --- Post-loop processing ---
        if num_tables > 1:
            write_pointer_row = self._write_grand_total_row(write_pointer_row, last_table_header_info, all_data_ranges, grand_total_pallets_for_summary_row)

        if self.sheet_config.get("summary", False) and last_table_header_info and self.args.DAF:
            write_pointer_row = invoice_utils.write_summary_rows(
                worksheet=self.worksheet,
                start_row=write_pointer_row,
                header_info=last_table_header_info,
                all_tables_data=all_tables_data,
                table_keys=table_keys,
                footer_config=self.sheet_config.get("footer_configurations", {}),
                mapping_rules=self.sheet_config.get('mappings', {}),
                styling_config=self.sheet_config.get("styling"),
                DAF_mode=self.args.DAF
            )

        if last_table_header_info:
            invoice_utils.apply_column_widths(
                self.worksheet,
                self.sheet_config.get("styling"),
                last_table_header_info.get('column_map')
            )
        
        final_row_spacing = self.sheet_config.get('row_spacing', 0)
        if final_row_spacing > 0 and num_tables > 0:
            # Just advance pointer, rows are already inserted
            write_pointer_row += final_row_spacing

        return True

    def _pre_calculate_and_insert_rows(self, table_keys, all_tables_data, header_to_write):
        total_rows_to_insert = 0
        num_tables = len(table_keys)
        add_blank_after_hdr_flag = self.sheet_config.get("add_blank_after_header", False)
        add_blank_before_ftr_flag = self.sheet_config.get("add_blank_before_footer", False)
        final_row_spacing = self.sheet_config.get('row_spacing', 0)
        summary_flag = self.sheet_config.get("summary", False)

        for i, table_key in enumerate(table_keys):
            table_data_to_fill = all_tables_data.get(str(table_key))
            if not table_data_to_fill or not isinstance(table_data_to_fill, dict):
                continue

            num_header_rows, _ = invoice_utils.calculate_header_dimensions(header_to_write)
            total_rows_to_insert += num_header_rows
            if add_blank_after_hdr_flag:
                total_rows_to_insert += 1
            
            max_len = max((len(v) for v in table_data_to_fill.values() if isinstance(v, list)), default=0)
            total_rows_to_insert += max_len

            if add_blank_before_ftr_flag:
                total_rows_to_insert += 1
            
            total_rows_to_insert += 1 # Footer

            if i < num_tables - 1:
                total_rows_to_insert += 1 # Spacer

        if num_tables > 1:
            total_rows_to_insert += 1 # Grand Total Row

        if summary_flag and num_tables > 0:
            total_rows_to_insert += 2 # Summary Flag Rows

        if final_row_spacing > 0:
            total_rows_to_insert += final_row_spacing

        if total_rows_to_insert > 0:
            try:
                start_row = self.sheet_config.get('start_row')
                self.worksheet.insert_rows(start_row, amount=total_rows_to_insert)
                return True, total_rows_to_insert
            except Exception as e:
                print(f"ERROR: Failed during bulk row insert: {e}")
                traceback.print_exc()
                return False, 0
        
        return True, 0

    def _write_grand_total_row(self, start_row, header_info, sum_ranges, pallet_count):
        """
        Writes the grand total footer row using invoice_utils helper function.
        Processors should not touch builders directly - they call worker functions.
        """
        print(f"\n--- Adding Grand Total Row at index {start_row} ---")
        try:
            footer_config = self.sheet_config.get("footer_configurations", {})
            
            # Use the utility function instead of builder directly
            next_row = invoice_utils.write_grand_total_footer(
                worksheet=self.worksheet,
                footer_row_num=start_row,
                header_info=header_info,
                sum_ranges=sum_ranges,
                footer_config=footer_config,
                pallet_count=pallet_count,
                sheet_config=self.sheet_config,
                DAF_mode=self.args.DAF
            )
            
            return next_row
            
        except Exception as e:
            print(f"--- ERROR in _write_grand_total_row: {e} ---")
            traceback.print_exc()
            self.processing_successful = False
            return start_row + 1
