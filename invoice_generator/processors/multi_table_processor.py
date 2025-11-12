# invoice_generator/processors/multi_table_processor.py
import sys
import logging
from .base_processor import SheetProcessor
from ..builders.layout_builder import LayoutBuilder
from ..builders.footer_builder import FooterBuilderStyler
from ..styling.models import StylingConfigModel
from ..config.builder_config_resolver import BuilderConfigResolver
import traceback
from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)

class MultiTableProcessor(SheetProcessor):
    """
    Processes a worksheet that contains multiple, repeating blocks of tables,
    such as a packing list. Uses LayoutBuilder for each table iteration.
    """

    def process(self) -> bool:
        """
        Executes the logic for processing a multi-table sheet using LayoutBuilder.
        """
        logger.info(f"Processing sheet '{self.sheet_name}' as multi-table/packing list")
        
        # Get all tables data
        all_tables_data = self.invoice_data.get('processed_tables_data', {})
        if not all_tables_data or not isinstance(all_tables_data, dict):
            logger.warning(f"'processed_tables_data' not found/valid. Skipping '{self.sheet_name}'")
            return True  # Not a failure, just nothing to do

        table_keys = sorted(all_tables_data.keys(), key=lambda x: int(x) if str(x).isdigit() else float('inf'))
        logger.info(f"Found {len(table_keys)} tables to process: {table_keys}")
        
        # Track the current row position as we build multiple tables
        # Start at header_row (where first table's header will be written)
        current_row = self.sheet_config.get('header_row', 1) if self.sheet_config else 1
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
            logger.info(f"Processing table '{table_key}' ({i+1}/{len(table_keys)})")
            
            # Prepare invoice_data for this specific table
            table_invoice_data = {
                'processed_tables_data': {
                    str(table_key): all_tables_data[str(table_key)]
                }
            }
            
            # Use BuilderConfigResolver to prepare bundles for this table
            resolver = BuilderConfigResolver(
                config_loader=self.config_loader,
                sheet_name=self.sheet_name,
                worksheet=self.output_worksheet,
                args=self.args,
                invoice_data=table_invoice_data,
                pallets=0  # Per-table, not grand total
            )
            
            # Get the bundles
            style_config = resolver.get_style_bundle()
            logger.debug(f"Style config: {style_config}")
            context_config = resolver.get_context_bundle(
                invoice_data=table_invoice_data,
                enable_text_replacement=False
            )
            layout_config = resolver.get_layout_bundle()
            
            # Get data bundle to extract header_info
            data_bundle = resolver.get_data_bundle(table_key=str(table_key))
            layout_config['header_info'] = data_bundle.get('header_info', {})
            layout_config['mapping_rules'] = data_bundle.get('mapping_rules', {})
            layout_config['data_source'] = data_bundle.get('data_source')
            layout_config['data_source_type'] = data_bundle.get('data_source_type')
            layout_config['header_row'] = current_row  # Override with current position for this table's header
            # NOTE: header_info from config is just column metadata, NOT styled Excel rows
            # HeaderBuilder still needs to run to write the actual styled header rows
            layout_config['enable_text_replacement'] = False
            # For multi-table: Only restore template header/footer for FIRST table
            layout_config['skip_template_header_restoration'] = (not is_first_table)
            layout_config['skip_template_footer_restoration'] = True  # Never restore footer mid-document
            
            layout_builder = LayoutBuilder(
                self.output_workbook,
                self.output_worksheet,
                self.template_worksheet,
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
                logger.error(f"Failed to build layout for table '{table_key}'")
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
            
            logger.debug(f"Table '{table_key}' complete. Next row: {current_row}, Pallets: {table_pallets}")
            logger.debug(f"  next_row_after_footer: {layout_builder.next_row_after_footer}")
            logger.debug(f"  data_start_row: {layout_builder.data_start_row}")
            logger.debug(f"  data_end_row: {layout_builder.data_end_row}")
        
        # After all tables, add grand total row if needed
        if len(table_keys) > 1 and last_header_info:
            logger.info("Adding Grand Total Row")
            grand_total_row = current_row
            
            # Create resolver for grand total footer (reuse last table's context)
            grand_total_resolver = BuilderConfigResolver(
                config_loader=self.config_loader,
                sheet_name=self.sheet_name,
                worksheet=self.output_worksheet,
                args=self.args,
                invoice_data=self.invoice_data,
                pallets=grand_total_pallets  # Use calculated grand total pallets
            )
            
            # Get bundles from resolver
            gt_style_config = grand_total_resolver.get_style_bundle()
            gt_layout_config = grand_total_resolver.get_layout_bundle()
            
            # Get styling model from style bundle
            styling_model = gt_style_config.get('styling_config')
            if styling_model and not isinstance(styling_model, StylingConfigModel):
                try:
                    styling_model = StylingConfigModel(**styling_model)
                except Exception as e:
                    logger.warning(f"Could not create StylingConfigModel: {e}")
                    styling_model = None
            
            # Get footer config and mappings from layout bundle
            footer_config = gt_layout_config.get('footer', {})
            footer_config_copy = footer_config.copy()
            footer_config_copy["type"] = "grand_total"  # Mark as grand total type
            
            # Add summary add-on if enabled in layout config
            if gt_layout_config.get('content', {}).get("summary", False) and self.args.DAF:
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
                'mapping_rules': gt_layout_config.get('data_flow', {}).get('mappings', {}),
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
            
            logger.debug(f"Grand Total Row added at row {grand_total_row}: {grand_total_pallets} pallets")
            current_row = next_row  # Update current_row for template footer restoration
        
        # Restore template footer at the very end after all tables and grand total
        if template_state_builder:
            logger.debug(f"\n--- Restoring Template Footer ---")
            logger.info(f"[MultiTableProcessor] Restoring template footer after row {current_row}")
            template_state_builder.restore_footer_only(
                target_worksheet=self.output_worksheet,
                footer_start_row=current_row
            )
            logger.info(f"[MultiTableProcessor] Template footer restored successfully")
        
        logger.info(f"Successfully processed {len(table_keys)} tables for sheet '{self.sheet_name}'.")
        return True
