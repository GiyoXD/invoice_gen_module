from typing import Any, Dict, Optional
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import Workbook

from ..styling.models import StylingConfigModel
from .header_builder import HeaderBuilder
from .data_table_builder import DataTableBuilder
from .footer_builder import FooterBuilder
from .text_replacement_builder import TextReplacementBuilder
from .template_state_builder import TemplateStateBuilder

class LayoutBuilder:
    """
    The Director in the Builder pattern.
    Coordinates all builders to construct the complete document layout.
    """
    def __init__(
        self,
        workbook: Workbook,
        worksheet: Worksheet,
        template_worksheet: Worksheet,
        sheet_name: str,
        sheet_config: Dict[str, Any],
        all_sheet_configs: Dict[str, Any],
        invoice_data: Dict[str, Any],
        styling_config: Optional[StylingConfigModel] = None,
        args: Optional[Any] = None,
        final_grand_total_pallets: int = 0,
        enable_text_replacement: bool = False,
        # Optional skip flags for custom processors
        skip_template_header_restoration: bool = False,
        skip_header_builder: bool = False,
        skip_data_table_builder: bool = False,
        skip_footer_builder: bool = False,
        skip_template_footer_restoration: bool = False,
    ):
        self.workbook = workbook  # Output workbook (writable)
        self.worksheet = worksheet  # Output worksheet (writable)
        self.template_worksheet = template_worksheet  # Template worksheet (read-only usage)
        self.sheet_name = sheet_name
        self.sheet_config = sheet_config
        self.all_sheet_configs = all_sheet_configs
        self.invoice_data = invoice_data
        self.styling_config = styling_config
        self.args = args
        self.final_grand_total_pallets = final_grand_total_pallets
        self.enable_text_replacement = enable_text_replacement
        
        # Skip flags for flexible processor customization
        self.skip_template_header_restoration = skip_template_header_restoration
        self.skip_header_builder = skip_header_builder
        self.skip_data_table_builder = skip_data_table_builder
        self.skip_footer_builder = skip_footer_builder
        self.skip_template_footer_restoration = skip_template_footer_restoration
        
        # Store results after build
        self.header_info = None
        self.next_row_after_footer = -1
        self.data_start_row = -1  # Expose data range for multi-table sum calculation
        self.data_end_row = -1    # Expose data range for multi-table sum calculation
        self.dynamic_desc_used = False  # Expose for summary add-on condition
        self.template_state_builder = None

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
        start_row = self.sheet_config.get('start_row', 1)
        header_to_write = self.sheet_config.get('header_to_write')
        num_header_cols = len(header_to_write) if header_to_write else 0
        
        # IMPORTANT: Template boundaries should ALWAYS be based on TEMPLATE's original start_row
        # The header in the template goes from row 1 to (start_row - 1)
        # For multi-table sheets, we use the ORIGINAL sheet_config start_row (from template),
        # not the dynamic start_row that changes for each table
        
        # Get the original start_row from the sheet config (this is the template's data start row)
        # This is where the data starts in the template, so header is everything before it
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
        else:
            print(f"[LayoutBuilder] Skipping template header restoration (skip_template_header_restoration=True)")
        
        # 4. Header Builder - writes header data to NEW worksheet (unless skipped)
        if not self.skip_header_builder:
            # Convert styling_config dict to StylingConfigModel if needed
            styling_model = self.styling_config
            if styling_model and not isinstance(styling_model, StylingConfigModel):
                try:
                    styling_model = StylingConfigModel(**styling_model)
                except Exception as e:
                    print(f"Warning: Could not create StylingConfigModel: {e}")
                    styling_model = None

            header_builder = HeaderBuilder(
                worksheet=self.worksheet,
                start_row=start_row,
                header_layout_config=header_to_write,
                sheet_styling_config=styling_model,
            )
            self.header_info = header_builder.build()

            if not self.header_info or not self.header_info.get('column_map'):
                print(f"Error: Cannot fill data for '{self.sheet_name}' because header_info or column_map is missing.")
                return False
        else:
            print(f"[LayoutBuilder] Skipping header builder (skip_header_builder=True)")
            # Must provide dummy header_info for downstream builders
            self.header_info = {'column_map': {}, 'first_row_index': start_row, 'second_row_index': start_row + 1}
            styling_model = self.styling_config

        # 5. Data Table Builder (writes data rows, returns footer position) (unless skipped)
        if not self.skip_data_table_builder:
            sheet_inner_mapping_rules_dict = self.sheet_config.get('mappings', {})
            add_blank_after_hdr_flag = self.sheet_config.get("add_blank_after_header", False)
            static_content_after_hdr_dict = self.sheet_config.get("static_content_after_header", {})
            add_blank_before_ftr_flag = self.sheet_config.get("add_blank_before_footer", False)
            static_content_before_ftr_dict = self.sheet_config.get("static_content_before_footer", {})
            merge_rules_after_hdr = self.sheet_config.get("merge_rules_after_header", {})
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

            data_table_builder = DataTableBuilder(
                worksheet=self.worksheet,
                sheet_name=self.sheet_name,
                sheet_config=self.sheet_config,
                all_sheet_configs=self.all_sheet_configs,
                data_source=data_to_fill,
                data_source_type=data_source_type,
                header_info=self.header_info,
                mapping_rules=sheet_inner_mapping_rules_dict,
                sheet_styling_config=styling_model,
                add_blank_after_header=add_blank_after_hdr_flag,
                static_content_after_header=static_content_after_hdr_dict,
                add_blank_before_footer=add_blank_before_ftr_flag,
                static_content_before_footer=static_content_before_ftr_dict,
                merge_rules_after_header=merge_rules_after_hdr,
                merge_rules_before_footer=merge_rules_before_ftr,
                merge_rules_footer=merge_rules_footer,
                max_rows_to_fill=None,
                grand_total_pallets=self.final_grand_total_pallets,
                custom_flag=self.args.custom,
                data_cell_merging_rules=data_cell_merging_rules,
                DAF_mode=self.args.DAF,
                all_tables_data=None,
                table_keys=None,
                is_last_table=True,
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
        if not self.skip_footer_builder:
            # Prepare footer parameters
            pallet_count = 0
            if data_source_type == "processed_tables":
                pallet_count = local_chunk_pallets
            else:
                pallet_count = self.final_grand_total_pallets

            # Get footer config and sum ranges
            footer_config = self.sheet_config.get('footer_configurations', {})
            sheet_inner_mapping_rules_dict = self.sheet_config.get('mappings', {})
            data_range_to_sum = []
            if data_start_row > 0 and data_end_row >= data_start_row:
                data_range_to_sum = [(data_start_row, data_end_row)]

            footer_builder = FooterBuilder(
                worksheet=self.worksheet,
                footer_row_num=footer_row_position,
                header_info=self.header_info,
                sum_ranges=data_range_to_sum,
                footer_config=footer_config,
                pallet_count=pallet_count,
                DAF_mode=data_source_type == "DAF_aggregation",
                sheet_styling_config=styling_model,
                all_tables_data=None,  # TODO: Pass if multi-table support needed
                table_keys=None,
                mapping_rules=sheet_inner_mapping_rules_dict,
                sheet_name=self.sheet_name,
                is_last_table=True,
                dynamic_desc_used=False,  # TODO: Track this if needed
            )
            self.next_row_after_footer = footer_builder.build()
            
            # Apply footer height to all footer rows (including add-ons like grand total)
            if self.next_row_after_footer > footer_row_position:
                # Multiple footer rows were created (e.g., regular footer + grand total)
                for footer_row in range(footer_row_position, self.next_row_after_footer):
                    self._apply_footer_row_height(footer_row, styling_model)
            else:
                # Single footer row
                self._apply_footer_row_height(footer_row_position, styling_model)
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
    
    def _apply_footer_row_height(self, footer_row: int, styling_config):
        """Helper method to apply footer height to a single footer row."""
        if not styling_config or not styling_config.rowHeights:
            return
        
        row_heights_cfg = styling_config.rowHeights
        footer_height_config = row_heights_cfg.get("footer")
        match_header_height_flag = row_heights_cfg.get("footer_matches_header_height", True)
        
        # Determine the footer height
        final_footer_height = None
        if match_header_height_flag:
            # Get header height from config
            header_height = row_heights_cfg.get("header")
            if header_height is not None:
                final_footer_height = header_height
        if final_footer_height is None and footer_height_config is not None:
            final_footer_height = footer_height_config
        
        # Apply the height
        if final_footer_height is not None and footer_row > 0:
            try:
                h_val = float(final_footer_height)
                if h_val > 0:
                    self.worksheet.row_dimensions[footer_row].height = h_val
            except (ValueError, TypeError):
                pass
