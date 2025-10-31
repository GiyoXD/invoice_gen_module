# invoice_generator/processors/single_table_processor.py
import sys
from .base_processor import SheetProcessor
from .. import invoice_utils
from ..utils import text_replace_utils
from ..builders.layout_builder import LayoutBuilder

class SingleTableProcessor(SheetProcessor):
    """
    Processes a worksheet that is configured to have a single main data table.
    This includes writing a header, filling the table, and applying styles.
    """
    def process(self) -> bool:
        """
        Executes the logic for processing a single-table sheet using the builder pattern.
        """
        print(f"Processing sheet '{self.sheet_name}' as single table/aggregation.")
        
        # Get styling configuration
        sheet_styling_config = self.sheet_config.get("styling")
        
        # Prepare three config bundles for LayoutBuilder
        style_config = {
            'styling_config': sheet_styling_config
        }
        
        context_config = {
            'sheet_name': self.sheet_name,
            'invoice_data': self.invoice_data,
            'all_sheet_configs': self.data_mapping_config,
            'args': self.args,
            'final_grand_total_pallets': self.final_grand_total_pallets
        }
        
        layout_config = {
            'sheet_config': self.sheet_config,
            'enable_text_replacement': False  # Already done at main level
        }
        
        # Use LayoutBuilder to orchestrate the entire layout construction
        layout_builder = LayoutBuilder(
            workbook=self.output_workbook,
            worksheet=self.output_worksheet,
            template_worksheet=self.template_worksheet,
            style_config=style_config,
            context_config=context_config,
            layout_config=layout_config
        )
        
        # Build the entire layout (header + table + footer)
        success = layout_builder.build()
        
        if not success:
            print(f"Failed to build layout for sheet '{self.sheet_name}'.")
            return False
            
        print(f"Successfully filled table data/footer for sheet '{self.sheet_name}'.")
        
        # Get results from the builder
        header_info = layout_builder.header_info
        next_row_after_footer = layout_builder.next_row_after_footer
        sheet_inner_mapping_rules_dict = self.sheet_config.get('mappings', {})
        final_row_spacing = self.sheet_config.get('row_spacing', 0)
        
        # Handle weight summary if enabled (this is a post-processing step)
        weight_summary_config = self.sheet_config.get("weight_summary_config", {})
        if weight_summary_config.get("enabled"):
            processed_tables_data = self.invoice_data.get('processed_tables_data', {})
            if processed_tables_data:
                next_row_after_footer = invoice_utils.write_grand_total_weight_summary(
                    worksheet=self.worksheet,
                    start_row=next_row_after_footer,
                    header_info=header_info,
                    processed_tables_data=processed_tables_data,
                    weight_config=weight_summary_config,
                    styling_config=self.sheet_config
                )
            else:
                print("Warning: Weight summary was enabled, but 'processed_tables_data' was not found in the source data.")

        # Apply column widths
        print(f"Applying column widths for sheet '{self.sheet_name}'...")
        invoice_utils.apply_column_widths(
            self.worksheet,
            sheet_styling_config,
            header_info.get('column_map')
        )

        # Insert final spacer rows
        if final_row_spacing >= 1:
            try:
                print(f"Config requests final spacing ({final_row_spacing}). Adding blank row(s) at {next_row_after_footer}.")
                self.worksheet.insert_rows(next_row_after_footer, amount=final_row_spacing)
            except Exception as final_spacer_err:
                print(f"Warning: Failed to insert final spacer rows: {final_spacer_err}")

        # Fill summary fields
        print("Attempting to fill summary fields...")
        summary_data_source = self.invoice_data.get('final_DAF_compounded_result', {})
        if summary_data_source and sheet_inner_mapping_rules_dict:
            for map_key, map_rule in sheet_inner_mapping_rules_dict.items():
                if isinstance(map_rule, dict) and 'marker' in map_rule:
                    target_cell = invoice_utils.find_cell_by_marker(self.worksheet, map_rule['marker'])
                    summary_value = summary_data_source.get(map_key)
                    if target_cell and summary_value is not None:
                        target_cell.value = summary_value

        return True
