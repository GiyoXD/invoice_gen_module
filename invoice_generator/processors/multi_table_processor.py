# invoice_generator/processors/multi_table_processor.py
import sys
from .base_processor import SheetProcessor
from .. import invoice_utils
from ..builders.layout_builder import LayoutBuilder
import traceback
from openpyxl.utils import get_column_letter

class MultiTableProcessor(SheetProcessor):
    """
    Processes a worksheet that contains multiple, repeating blocks of tables,
    such as a packing list. Uses LayoutBuilder for each table iteration.
    """

    def process(self) -> bool:
        """
        Executes the logic for processing a multi-table sheet using LayoutBuilder.
        """
        print(f"Processing sheet '{self.sheet_name}' as multi-table/packing list.")
        
        # Get all tables data
        all_tables_data = self.invoice_data.get('processed_tables_data', {})
        if not all_tables_data or not isinstance(all_tables_data, dict):
            print(f"Warning: 'processed_tables_data' not found/valid. Skipping '{self.sheet_name}'.")
            return True  # Not a failure, just nothing to do

        table_keys = sorted(all_tables_data.keys(), key=lambda x: int(x) if str(x).isdigit() else float('inf'))
        print(f"Found {len(table_keys)} tables to process: {table_keys}")
        
        # Get styling configuration
        sheet_styling_config = self.sheet_config.get("styling")
        
        # Track the current row position as we build multiple tables
        current_row = self.sheet_config.get('start_row', 1)
        all_data_ranges = []
        grand_total_pallets = 0
        last_header_info = None
        
        # Process each table using LayoutBuilder
        for i, table_key in enumerate(table_keys):
            is_last_table = (i == len(table_keys) - 1)
            print(f"\n--- Processing table '{table_key}' ({i+1}/{len(table_keys)}) ---")
            
            # Prepare invoice_data for this specific table
            table_invoice_data = {
                'processed_tables_data': {
                    str(table_key): all_tables_data[str(table_key)]
                }
            }
            
            # Use LayoutBuilder for this table iteration
            # For multi-table, we use a modified sheet_config that points to the current row
            table_sheet_config = self.sheet_config.copy()
            table_sheet_config['start_row'] = current_row
            
            layout_builder = LayoutBuilder(
                workbook=self.output_workbook,
                worksheet=self.output_worksheet,
                template_worksheet=self.template_worksheet,
                sheet_name=self.sheet_name,
                sheet_config=table_sheet_config,
                all_sheet_configs=self.data_mapping_config,
                invoice_data=table_invoice_data,
                styling_config=sheet_styling_config,
                args=self.args,
                final_grand_total_pallets=0,  # Per-table, not grand total
                enable_text_replacement=False  # Already done at main level
            )
            
            # Build this table's layout
            success = layout_builder.build()
            
            if not success:
                print(f"Failed to build layout for table '{table_key}'.")
                return False
            
            # Update tracking variables
            last_header_info = layout_builder.header_info
            current_row = layout_builder.next_row_after_footer
            
            # Track pallet count for grand total
            table_data = all_tables_data.get(str(table_key), {})
            pallet_counts = table_data.get('pallet_count', [])
            table_pallets = sum(int(p) for p in pallet_counts if str(p).isdigit())
            grand_total_pallets += table_pallets
            
            print(f"Table '{table_key}' complete. Next row: {current_row}, Pallets: {table_pallets}")
        
        # After all tables, add grand total row if needed
        if len(table_keys) > 1 and last_header_info:
            print(f"\n--- Adding Grand Total Row ---")
            grand_total_row = current_row
            
            # Write grand total using invoice_utils
            invoice_utils.write_grand_total_pallet_summary(
                worksheet=self.output_worksheet,
                start_row=grand_total_row,
                header_info=last_header_info,
                total_pallets=grand_total_pallets,
                sheet_styling_config=self.sheet_config
            )
            
            print(f"Grand Total Row added at row {grand_total_row}: {grand_total_pallets} pallets")
        
        print(f"Successfully processed {len(table_keys)} tables for sheet '{self.sheet_name}'.")
        return True
