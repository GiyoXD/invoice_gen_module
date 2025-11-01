# invoice_generator/processors/multi_table_processor.py
import sys
from .base_processor import SheetProcessor
from ..builders.layout_builder import LayoutBuilder
from ..builders.footer_builder import FooterBuilderStyler
from ..styling.models import StylingConfigModel
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
        template_state_builder = None  # Save from first table for final footer restoration
        dynamic_desc_used = False  # Track if any table used dynamic description (for summary add-on)
        
        # Process each table using LayoutBuilder
        # IMPORTANT: For multi-table, skip template restoration after first table
        # to avoid capturing template state from wrong row positions
        for i, table_key in enumerate(table_keys):
            is_first_table = (i == 0)
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
            # IMPORTANT: Set data_source to the table key so LayoutBuilder can find the data
            table_sheet_config['data_source'] = str(table_key)
            
            # Prepare config bundles for LayoutBuilder
            style_config = {
                'styling_config': sheet_styling_config
            }
            
            context_config = {
                'sheet_name': self.sheet_name,
                'invoice_data': table_invoice_data,
                'all_sheet_configs': self.data_mapping_config,
                'args': self.args,
                'final_grand_total_pallets': 0,  # Per-table, not grand total
                'config_loader': self.config_loader  # For direct bundled config access
            }
            
            layout_config = {
                'sheet_config': table_sheet_config,
                'enable_text_replacement': False,  # Already done at main level
                # For multi-table: Only restore template header/footer for FIRST table
                # Subsequent tables should skip to avoid wrong row capture
                'skip_template_header_restoration': (not is_first_table),
                'skip_template_footer_restoration': True  # Never restore footer mid-document
                # NOTE: HeaderBuilder writes the TABLE column headers (e.g., "Mark & No", "Description")
                # This is DIFFERENT from template header (static content like company info)
                # HeaderBuilder should run for EACH table to write column headers
            }
            
            layout_builder = LayoutBuilder(
                workbook=self.output_workbook,
                worksheet=self.output_worksheet,
                template_worksheet=self.template_worksheet,
                style_config=style_config,
                context_config=context_config,
                layout_config=layout_config
            )
            
            # Build this table's layout
            success = layout_builder.build()
            
            # Save template state builder from first table for final footer restoration
            if is_first_table:
                template_state_builder = layout_builder.template_state_builder
            
            if not success:
                print(f"Failed to build layout for table '{table_key}'.")
                return False
            
            # Update tracking variables
            last_header_info = layout_builder.header_info
            current_row = layout_builder.next_row_after_footer
            
            # Add 1 blank row spacing after each table footer (except the last one)
            if not is_last_table:
                current_row += 1
            
            # Collect data range for grand total sum formulas
            if layout_builder.data_start_row > 0 and layout_builder.data_end_row >= layout_builder.data_start_row:
                all_data_ranges.append((layout_builder.data_start_row, layout_builder.data_end_row))
            
            # Track if dynamic description was used (needed for summary add-on)
            if layout_builder.dynamic_desc_used:
                dynamic_desc_used = True
            
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
            
            # Get styling configuration
            styling_model = sheet_styling_config
            if styling_model and not isinstance(styling_model, StylingConfigModel):
                try:
                    styling_model = StylingConfigModel(**styling_model)
                except Exception as e:
                    print(f"Warning: Could not create StylingConfigModel: {e}")
                    styling_model = None
            
            # Use FooterBuilder to create grand total footer (proper builder pattern)
            footer_config = self.sheet_config.get('footer_configurations', {})
            footer_config_copy = footer_config.copy()
            footer_config_copy["type"] = "grand_total"  # Mark as grand total type
            
            # Add summary add-on if enabled in sheet config
            if self.sheet_config.get("summary", False) and self.args.DAF:
                footer_config_copy["add_ons"] = ["summary"]
            
            # Bundle configs for FooterBuilder
            fb_style_config = {
                'styling_config': styling_model
            }
            
            fb_context_config = {
                'header_info': last_header_info,
                'pallet_count': grand_total_pallets,
                'sheet_name': self.sheet_name,
                'is_last_table': True,
                'dynamic_desc_used': dynamic_desc_used
            }
            
            fb_data_config = {
                'sum_ranges': all_data_ranges,
                'footer_config': footer_config_copy,
                'all_tables_data': all_tables_data,
                'table_keys': table_keys,
                'mapping_rules': self.sheet_config.get('mappings', {}),
                'DAF_mode': self.args.DAF,
                'override_total_text': None
            }
            
            footer_builder = FooterBuilderStyler(
                worksheet=self.output_worksheet,
                footer_row_num=grand_total_row,
                style_config=fb_style_config,
                context_config=fb_context_config,
                data_config=fb_data_config
            )
            next_row = footer_builder.build()
            
            print(f"Grand Total Row added at row {grand_total_row}: {grand_total_pallets} pallets")
            current_row = next_row  # Update current_row for template footer restoration
        
        # Restore template footer at the very end after all tables and grand total
        if template_state_builder:
            print(f"\n--- Restoring Template Footer ---")
            print(f"[MultiTableProcessor] Restoring template footer after row {current_row}")
            template_state_builder.restore_footer_only(
                target_worksheet=self.output_worksheet,
                footer_start_row=current_row
            )
            print(f"[MultiTableProcessor] Template footer restored successfully")
        
        print(f"Successfully processed {len(table_keys)} tables for sheet '{self.sheet_name}'.")
        return True
