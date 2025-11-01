# invoice_generator/processors/single_table_processor.py
import sys
from .base_processor import SheetProcessor
from ..utils import text_replace_utils
from ..builders.layout_builder import LayoutBuilder
from ..config.builder_config_resolver import BuilderConfigResolver

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
        
        # Use BuilderConfigResolver to prepare bundles cleanly
        resolver = BuilderConfigResolver(
            config_loader=self.config_loader,
            sheet_name=self.sheet_name,
            worksheet=self.output_worksheet,
            args=self.args,
            invoice_data=self.invoice_data,
            pallets=self.final_grand_total_pallets,
            final_grand_total_pallets=self.final_grand_total_pallets  # Context override
        )
        
        # Get the bundles needed for LayoutBuilder
        style_config = resolver.get_style_bundle()
        context_config = resolver.get_context_bundle(
            invoice_data=self.invoice_data,
            enable_text_replacement=False  # Already done at main level
        )
        layout_config = resolver.get_layout_bundle()
        layout_config['enable_text_replacement'] = False
        
        # Get data bundle to extract header_info and mapping_rules
        data_bundle = resolver.get_data_bundle()
        layout_config['header_info'] = data_bundle.get('header_info', {})
        layout_config['mapping_rules'] = data_bundle.get('mapping_rules', {})
        layout_config['data_source'] = data_bundle.get('data_source')
        layout_config['data_source_type'] = data_bundle.get('data_source_type')
        layout_config['skip_header_builder'] = True  # Using pre-constructed header_info from resolver
        
        # Use LayoutBuilder to orchestrate the entire layout construction
        layout_builder = LayoutBuilder(
            self.output_workbook,
            self.output_worksheet,
            self.template_worksheet,
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
        
        # TODO: Re-implement post-processing features using new architecture:
        # - Weight summary (should be a builder add-on)
        # - Column widths (should be handled by styling in builders)
        # - Summary fields (should be part of data mapping)
        
        return True
