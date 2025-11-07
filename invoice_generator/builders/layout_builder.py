from typing import Any, Dict, Optional
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import Workbook

from ..styling.models import StylingConfigModel
from .bundle_accessor import BundleAccessor
from .header_builder import HeaderBuilderStyler
from .data_table_builder import DataTableBuilderStyler
from .footer_builder import FooterBuilderStyler
from .text_replacement_builder import TextReplacementBuilder
from .template_state_builder import TemplateStateBuilder

import logging


class LayoutBuilder(BundleAccessor):
    """
    The Director in the Builder pattern.
    Coordinates BuilderStyler components to construct the complete document layout.
    
    This class orchestrates the building process in the correct sequence:
    1. Template state capture (structure preservation)
    2. Text replacement (if enabled)
    3. Header building + styling (via HeaderBuilderStyler)
    4. Data table building + styling (via DataTableBuilderStyler)
    5. Footer building + styling (via FooterBuilderStyler)
    6. Template footer restoration
    
    Uses pure bundle architecture for zero duplication and infinite extensibility.
    """
    def __init__(
        self,
        workbook: Workbook,
        worksheet: Worksheet,
        template_worksheet: Worksheet,
        style_config: Dict[str, Any],
        context_config: Dict[str, Any],
        layout_config: Dict[str, Any],
    ):
        """
        Initialize LayoutBuilder with pure bundle pattern.
        
        Args:
            workbook: Output workbook (writable)
            worksheet: Output worksheet (writable)
            template_worksheet: Template worksheet (read-only)
            style_config: Bundle containing styling_config
            context_config: Bundle containing sheet_name, invoice_data, all_sheet_configs, args, final_grand_total_pallets, config_loader
            layout_config: Bundle containing sheet_config, enable_text_replacement, skip flags
        """
        # Initialize base class with common bundles
        super().__init__(
            worksheet=worksheet,
            style_config=style_config,
            context_config=context_config,
            layout_config=layout_config  # Pass layout_config to base via kwargs
        )
        
        # Store LayoutBuilder-specific attributes
        self.workbook = workbook
        self.template_worksheet = template_worksheet
        
        # Check if we have a config_loader for direct bundled config access
        self.config_loader = context_config.get('config_loader')
        
        # Initialize output state variables (build results, not configs)
        self.header_info = None
        self.next_row_after_footer = -1
        self.data_start_row = -1
        self.data_end_row = -1
        self.dynamic_desc_used = False
        self.template_state_builder = None
    
    # ========== Properties for Frequently Accessed Config Values ==========
    # Note: sheet_name, all_sheet_configs, args, sheet_styling_config inherited from BundleAccessor
    
    @property
    def sheet_config(self) -> Dict[str, Any]:
        """
        Sheet configuration - only used for legacy configs now.
        Bundled configs read directly from config_loader in the build() method.
        """
        return self.layout_config.get('sheet_config', {})
    
    @property
    def invoice_data(self) -> Dict[str, Any]:
        """Invoice data from context config."""
        return self.context_config.get('invoice_data', {})
    
    @property
    def final_grand_total_pallets(self) -> int:
        """Final grand total pallets from context config."""
        return self.context_config.get('final_grand_total_pallets', 0)
    
    @property
    def enable_text_replacement(self) -> bool:
        """Enable text replacement from layout config."""
        return self.layout_config.get('enable_text_replacement', False)
    
    @property
    def skip_template_header_restoration(self) -> bool:
        """Skip template header restoration from layout config."""
        return self.layout_config.get('skip_template_header_restoration', False)
    
    @property
    def skip_header_builder(self) -> bool:
        """Skip header builder from layout config."""
        return self.layout_config.get('skip_header_builder', False)
    
    @property
    def skip_data_table_builder(self) -> bool:
        """Skip data table builder from layout config."""
        return self.layout_config.get('skip_data_table_builder', False)
    
    @property
    def skip_footer_builder(self) -> bool:
        """Skip footer builder from layout config."""
        return self.layout_config.get('skip_footer_builder', False)
    
    @property
    def skip_template_footer_restoration(self) -> bool:
        """Skip template footer restoration from layout config."""
        return self.layout_config.get('skip_template_footer_restoration', False)

    def build(self) -> bool:
        """
        Orchestrates all builders in the correct sequence.
        Reads template state from template_worksheet, writes to self.worksheet (output).
        This completely avoids merge conflicts since template and output are separate.
        """
        print(f"[LayoutBuilder] Building layout for sheet '{self.sheet_name}'")
        print(f"[LayoutBuilder] Reading from template, writing to output worksheet")
        
        # 1. Text Replacement (if enabled) - Pre-processing
        # Note: This was already done at workbook level, skip here
        if self.enable_text_replacement:
            text_replacer = TextReplacementBuilder(
                workbook=self.workbook,
                invoice_data=self.invoice_data
            )
            if self.args and self.args.DAF:
                text_replacer.build()  # Run both placeholder and DAF replacements
            else:
                text_replacer._replace_placeholders()  # Only placeholders
        
        # 2. Calculate header boundaries for template state capture
        if self.config_loader:
            structure = self.config_loader.get_sheet_structure(self.sheet_name)
            # For multi-table sheets, sheet_config['start_row'] is set by multi_table_processor for each table
            # Prioritize sheet_config['start_row'] if explicitly set (for multi-table), otherwise use bundled config
            start_row = self.sheet_config.get('start_row', structure.get('start_row', 1))
            header_to_write = None  # Will be converted from columns in HeaderBuilder
            num_header_cols = len(structure.get('columns', []))
        else:
            start_row = self.sheet_config.get('start_row', 1)
            header_to_write = self.sheet_config.get('header_to_write')
            num_header_cols = len(header_to_write) if header_to_write else 0
        
        # IMPORTANT: Template boundaries should ALWAYS be based on TEMPLATE's original start_row
        # The header in the template goes from row 1 to (start_row - 1)
        # For multi-table sheets, we use the ORIGINAL sheet_config start_row (from template),
        # not the dynamic start_row that changes for each table
        
        # Get the original start_row from the template (NOT from modified sheet_config)
        # For bundled config, get from config_loader (unmodified)
        # For legacy config, get from all_sheet_configs (unmodified)
        if self.config_loader:
            structure = self.config_loader.get_sheet_structure(self.sheet_name)
            original_start_row = structure.get('start_row', 1)
        else:
            original_start_row = self.all_sheet_configs.get(self.sheet_name, {}).get('start_row', start_row)
        
        template_header_start_row = 1  # Template header always starts at row 1
        template_header_end_row = original_start_row - 1  # Header ends one row before data starts
        
        # Calculate footer_start_row from template (estimate: original_start_row + minimal data rows)
        template_footer_start_row = original_start_row + 2  # Footer in template after header + minimal data
        
        # 3. Template State Capture - Capture from template_worksheet
        print(f"[LayoutBuilder] Capturing template state from template worksheet")
        self.template_state_builder = TemplateStateBuilder(
            worksheet=self.template_worksheet,  # Read from template
            num_header_cols=num_header_cols,
            header_end_row=template_header_end_row,  # Use template position, not output position
            footer_start_row=template_footer_start_row  # Use template position, not output position
        )
        
        # 3b. Restore ONLY header to output worksheet (unless skipped)
        if not self.skip_template_header_restoration:
            print(f"[LayoutBuilder] Restoring header from template to output worksheet")
            self.template_state_builder.restore_header_only(target_worksheet=self.worksheet)
            # After restoring template header, the next available row is after the template header
            next_available_row_after_template = template_header_end_row + 1
        else:
            print(f"[LayoutBuilder] Skipping template header restoration (skip_template_header_restoration=True)")
            # No template header restored, so next available row is the dynamic start_row
            next_available_row_after_template = start_row
        
        # 4. Header Builder - writes TABLE column headers at next available row (unless skipped)
        # NOTE: This is DIFFERENT from template header - these are the column headers like "Mark & No", "Description"
        if not self.skip_header_builder:
            # Use bundled columns if config_loader is available
            if self.config_loader:
                print(f"[LayoutBuilder] Reading from BUNDLED CONFIG (layout_bundle.{self.sheet_name}.structure)")
                structure = self.config_loader.get_sheet_structure(self.sheet_name)
                bundled_columns = structure.get('columns', [])
                print(f"[LayoutBuilder] Found {len(bundled_columns)} bundled columns")
                print(f"[LayoutBuilder] Writing table column headers at row {next_available_row_after_template}")
                # Print first column as sample (safely encoded)
                if bundled_columns:
                    sample = str(bundled_columns[0]).encode('ascii', 'replace').decode('ascii')
                    print(f"[LayoutBuilder] Sample column: {sample}")
                header_builder = HeaderBuilderStyler(
                    worksheet=self.worksheet,
                    start_row=next_available_row_after_template,
                    bundled_columns=bundled_columns,
                    sheet_styling_config=self.sheet_styling_config,
                )
            else:
                # Legacy format
                print(f"[LayoutBuilder] Writing table column headers at row {next_available_row_after_template}")
                header_builder = HeaderBuilderStyler(
                    worksheet=self.worksheet,
                    start_row=next_available_row_after_template,
                    header_layout_config=header_to_write,
                    sheet_styling_config=self.sheet_styling_config,
                )
            
            self.header_info = header_builder.build()

            if not self.header_info or not self.header_info.get('column_map'):
                print(f"Error: Cannot fill data for '{self.sheet_name}' because header_info or column_map is missing.")
                return False
        else:
            print(f"[LayoutBuilder] Skipping header builder (skip_header_builder=True)")
            # When skipping, we must get the pre-constructed header_info from the layout_config
            self.header_info = self.layout_config.get('header_info')
            if not self.header_info:
                print("Error: skip_header_builder is True, but no header_info was provided in layout_config.")
                # Provide a minimal dummy header_info to avoid crashing downstream
                self.header_info = {
                    'column_map': {},
                    'num_columns': 0,
                    'first_row_index': next_available_row_after_template, 
                    'second_row_index': next_available_row_after_template + 1
                }

        # 5. Data Table Builder (writes data rows, returns footer position) (unless skipped)
        if not self.skip_data_table_builder:
            # Get mappings from bundled config if available, else from legacy
            if self.config_loader:
                data_flow = self.config_loader.get_sheet_data_flow(self.sheet_name)
                structure = self.config_loader.get_sheet_structure(self.sheet_name)
                content = self.config_loader.get_sheet_content(self.sheet_name)
                
                # Convert bundled mappings to internal format
                sheet_inner_mapping_rules_dict = self._convert_bundled_mappings(
                    data_flow.get('mappings', {}),
                    structure.get('columns', [])
                )
                
                # Get content settings
                static_info = content.get('static', {})
                add_blank_after_hdr_flag = False  # Not in bundled config yet
                static_content_after_hdr_dict = {}  # Not in bundled config yet
                add_blank_before_ftr_flag = False  # Handled in footer config
                static_content_before_ftr_dict = {}
                if 'before_footer' in static_info:
                    bf = static_info['before_footer']
                    if 'col' in bf:
                        static_content_before_ftr_dict[str(bf['col'])] = bf.get('text', '')
                merge_rules_after_hdr = {}  # Not in bundled config yet
            else:
                # Legacy format
                sheet_inner_mapping_rules_dict = self.sheet_config.get('mappings', {})
                add_blank_after_hdr_flag = self.sheet_config.get("add_blank_after_header", False)
                static_content_after_hdr_dict = self.sheet_config.get("static_content_after_header", {})
                add_blank_before_ftr_flag = self.sheet_config.get("add_blank_before_footer", False)
                static_content_before_ftr_dict = self.sheet_config.get("static_content_before_footer", {})
                merge_rules_after_hdr = self.sheet_config.get("merge_rules_after_header", {})
            # Merge rules - not yet in bundled config, use empty for now
            if self.config_loader:
                merge_rules_before_ftr = {}
                merge_rules_footer = {}
                data_cell_merging_rules = None
                # IMPORTANT: Check sheet_config first for multi-table override (e.g., table key '1', '2')
                # Multi-table processor sets sheet_config['data_source'] to the specific table key
                # For multi-table, data_source is set directly on layout_config. For single table, it's in sheet_config.
                data_source_indicator = self.layout_config.get("data_source") or self.sheet_config.get("data_source") or self.config_loader.get_data_source(self.sheet_name)
            else:
                merge_rules_before_ftr = self.sheet_config.get("merge_rules_before_footer", {})
                merge_rules_footer = self.sheet_config.get("merge_rules_footer", {})
                data_cell_merging_rules = self.sheet_config.get("data_cell_merging_rule", None)
                data_source_indicator = self.sheet_config.get("data_source")

            data_to_fill = None
            data_source_type = None

            if self.args.custom and data_source_indicator == 'aggregation':
                data_to_fill = self.invoice_data.get('custom_aggregation_results')
                data_source_type = 'custom_aggregation'

            if data_to_fill is None:
                if self.args.DAF and self.sheet_name in ["Invoice", "Contract"]:
                    data_source_indicator = 'DAF_aggregation'

                if data_source_indicator == 'DAF_aggregation':
                    data_to_fill = self.invoice_data.get('final_DAF_compounded_result')
                    data_source_type = 'DAF_aggregation'
                elif data_source_indicator == 'aggregation':
                    data_to_fill = self.invoice_data.get('standard_aggregation_results')
                    data_source_type = 'aggregation'
                elif 'processed_tables_data' in self.invoice_data and data_source_indicator in self.invoice_data.get('processed_tables_data', {}):
                    data_to_fill = self.invoice_data['processed_tables_data'].get(data_source_indicator)
                    data_source_type = 'processed_tables'

            if data_to_fill is None:
                print(f"Warning: Data source '{data_source_indicator}' unknown or data empty. Skipping fill.")
                return True

            # Bundle configs for DataTableBuilder
            dtb_style_config = {
                'styling_config': self.sheet_styling_config
            }
            
            dtb_context_config = {
                'sheet_name': self.sheet_name,
                'all_sheet_configs': self.all_sheet_configs,
                'args': self.args,
                'grand_total_pallets': self.final_grand_total_pallets,
                'all_tables_data': None,
                'table_keys': None,
                'is_last_table': True
            }
            
            dtb_layout_config = {
                'sheet_config': self.sheet_config,
                'add_blank_after_header': add_blank_after_hdr_flag,
                'static_content_after_header': static_content_after_hdr_dict,
                'add_blank_before_footer': add_blank_before_ftr_flag,
                'static_content_before_footer': static_content_before_ftr_dict,
                'merge_rules_after_header': merge_rules_after_hdr,
                'merge_rules_before_footer': merge_rules_before_ftr,
                'merge_rules_footer': merge_rules_footer,
                'data_cell_merging_rules': data_cell_merging_rules,
                'max_rows_to_fill': None
            }
            
            dtb_data_config = {
                'data_source': data_to_fill,
                'data_source_type': data_source_type,
                'header_info': self.header_info,
                'mapping_rules': sheet_inner_mapping_rules_dict
            }

            data_table_builder = DataTableBuilderStyler(
                worksheet=self.worksheet,
                style_config=dtb_style_config,
                context_config=dtb_context_config,
                layout_config=dtb_layout_config,
                data_config=dtb_data_config
            )

            fill_success, footer_row_position, data_start_row, data_end_row, local_chunk_pallets = data_table_builder.build()

            # Store data range for multi-table processors to access
            self.data_start_row = data_start_row
            self.data_end_row = data_end_row
            self.dynamic_desc_used = data_table_builder.dynamic_desc_used  # Track for summary add-on

            if not fill_success:
                print(f"Failed to fill table data for sheet '{self.sheet_name}'.")
                return False
        else:
            print(f"[LayoutBuilder] Skipping data table builder (skip_data_table_builder=True)")
            # Provide dummy values for downstream builders
            footer_row_position = start_row + 2  # After header
            data_start_row = 0
            data_end_row = 0
            local_chunk_pallets = 0
            data_source_type = None
        
        # 6. Footer Builder (proper Director pattern - called explicitly by LayoutBuilder) (unless skipped)
        print(f"[LayoutBuilder] Checking FooterBuilder - skip_footer_builder={self.skip_footer_builder}")
        if not self.skip_footer_builder:
            # Prepare footer parameters
            pallet_count = 0
            if data_source_type == "processed_tables":
                pallet_count = local_chunk_pallets
            else:
                pallet_count = self.final_grand_total_pallets

            # Get footer config and sum ranges
            if self.config_loader:
                footer_config = self.config_loader.get_sheet_footer(self.sheet_name)
                print(f"[LayoutBuilder] Footer config from bundled: {footer_config}")
                # Get mappings (already converted if needed above)
                data_flow = self.config_loader.get_sheet_data_flow(self.sheet_name)
                structure = self.config_loader.get_sheet_structure(self.sheet_name)
                sheet_inner_mapping_rules_dict = self._convert_bundled_mappings(
                    data_flow.get('mappings', {}),
                    structure.get('columns', [])
                )
            else:
                footer_config = self.sheet_config.get('footer_configurations', {})
                sheet_inner_mapping_rules_dict = self.sheet_config.get('mappings', {})
            
            data_range_to_sum = []
            if data_start_row > 0 and data_end_row >= data_start_row:
                data_range_to_sum = [(data_start_row, data_end_row)]

            # Bundle configs for FooterBuilder
            fb_style_config = {
                'styling_config': self.sheet_styling_config
            }
            
            fb_context_config = {
                'header_info': self.header_info,
                'pallet_count': pallet_count,
                'sheet_name': self.sheet_name,
                'is_last_table': True,
                'dynamic_desc_used': False  # TODO: Track this if needed
            }
            
            fb_data_config = {
                'sum_ranges': data_range_to_sum,
                'footer_config': footer_config,
                'all_tables_data': None,  # TODO: Pass if multi-table support needed
                'table_keys': None,
                'mapping_rules': sheet_inner_mapping_rules_dict,
                'DAF_mode': data_source_type == "DAF_aggregation",
                'override_total_text': None
            }

            print(f"[LayoutBuilder] Creating FooterBuilder at row {footer_row_position}")
            footer_builder = FooterBuilderStyler(
                worksheet=self.worksheet,
                footer_row_num=footer_row_position,
                style_config=fb_style_config,
                context_config=fb_context_config,
                data_config=fb_data_config
            )
            print(f"[LayoutBuilder] Calling FooterBuilder.build()")
            self.next_row_after_footer = footer_builder.build()
            print(f"[LayoutBuilder] FooterBuilder returned next_row: {self.next_row_after_footer}")
            
            # Apply footer height to all footer rows (including add-ons like grand total)
            if self.next_row_after_footer > footer_row_position:
                # Multiple footer rows were created (e.g., regular footer + grand total)
                for footer_row in range(footer_row_position, self.next_row_after_footer):
                    self._apply_footer_row_height(footer_row)
            else:
                # Single footer row
                self._apply_footer_row_height(footer_row_position)
        else:
            print(f"[LayoutBuilder] Skipping footer builder (skip_footer_builder=True)")
            # No footer, so next row is right after data (or header if no data)
            self.next_row_after_footer = footer_row_position
        
        # 7. Template Footer Restoration (unless skipped)
        # Restore the template footer (static content like "Manufacture:", etc.) AFTER the dynamic footer
        # This places the template footer below the data footer
        if not self.skip_template_footer_restoration:
            write_pointer_row = self.next_row_after_footer  # Next available row after dynamic footer
            
            print(f"[LayoutBuilder] Restoring template footer after row {write_pointer_row}")
            self.template_state_builder.restore_footer_only(
                target_worksheet=self.worksheet,  # Write to output worksheet
                footer_start_row=write_pointer_row
            )
        else:
            print(f"[LayoutBuilder] Skipping template footer restoration (skip_template_footer_restoration=True)")
        
        print(f"[LayoutBuilder] Layout built successfully for sheet '{self.sheet_name}'")
        
        return True
    
    # ========== Helper Methods for Bundled Config Conversion ==========
    
    def _convert_bundled_mappings(self, data_flow_mappings: dict, columns: list) -> dict:
        """
        Convert bundled data_flow.mappings to internal mappings format.
        This is the same logic as in the sheet_config property.
        """
        return self._build_mappings_from_data_flow(data_flow_mappings, columns)
    
    def _build_header_from_structure(self, columns: list) -> list:
        """Convert bundled config columns to header_to_write format."""
        headers = []
        col_index = 0
        
        for col in columns:
            col_id = col.get('id', '')
            header_text = col.get('header', '')
            rowspan = col.get('rowspan', 1)
            colspan = col.get('colspan', 1)
            
            # Handle parent column with children (e.g., Quantity with PCS/SF)
            if 'children' in col:
                # Add parent header
                headers.append({
                    'row': 0,
                    'col': col_index,
                    'text': header_text,
                    'id': col_id,
                    'rowspan': 1,
                    'colspan': len(col['children'])
                })
                
                # Add children headers
                for child in col['children']:
                    headers.append({
                        'row': 1,
                        'col': col_index,
                        'text': child.get('header', ''),
                        'id': child.get('id', ''),
                        'rowspan': 1,
                        'colspan': 1
                    })
                    col_index += 1
            else:
                headers.append({
                    'row': 0,
                    'col': col_index,
                    'text': header_text,
                    'id': col_id,
                    'rowspan': rowspan,
                    'colspan': colspan
                })
                col_index += 1
        
        return headers
    
    def _build_mappings_from_data_flow(self, data_flow_mappings: dict, columns: list) -> dict:
        """Convert bundled config data_flow.mappings to legacy mappings format."""
        # First, extract formats from columns
        column_formats = {}
        for col in columns:
            col_id = col.get('id', '')
            col_format = col.get('format')
            if col_format:
                column_formats[col_id] = col_format
            
            # Handle children columns
            if 'children' in col:
                for child in col['children']:
                    child_id = child.get('id', '')
                    child_format = child.get('format')
                    if child_format:
                        column_formats[child_id] = child_format
        
        # Now build mappings
        mappings = {}
        for field_name, mapping in data_flow_mappings.items():
            column_id = mapping.get('column')
            
            legacy_mapping = {
                'id': column_id
            }
            
            # Handle source_key (old key_index)
            if 'source_key' in mapping:
                legacy_mapping['key_index'] = mapping['source_key']
            
            # Handle source_value (old value_key)
            if 'source_value' in mapping:
                legacy_mapping['value_key'] = mapping['source_value']
            
            # Handle formula
            if 'formula' in mapping:
                legacy_mapping['formula'] = mapping['formula']
            
            # Handle fallback
            if 'fallback' in mapping:
                legacy_mapping['fallback_on_none'] = mapping['fallback']
                legacy_mapping['fallback_on_DAF'] = mapping['fallback']
            
            # Add format from columns if available
            if column_id and column_id in column_formats:
                legacy_mapping['number_format'] = column_formats[column_id]
            
            mappings[field_name] = legacy_mapping
        
        return mappings
