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
        
        # Store results after build
        self.header_info = None
        self.next_row_after_footer = -1
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
        
        # Calculate header_end_row (typically start_row + 1 for 2-row headers)
        header_end_row = start_row + 1  # Assumption: 2-row header (row 0 and row 1)
        
        # Calculate footer_start_row from template
        # In templates, footer typically starts right after minimal data area
        footer_start_row_template = start_row + 3  # Conservative estimate
        
        # 3. Template State Capture - Capture from template_worksheet
        print(f"[LayoutBuilder] Capturing template state from template worksheet")
        self.template_state_builder = TemplateStateBuilder(
            worksheet=self.template_worksheet,  # Read from template
            num_header_cols=num_header_cols,
            header_end_row=header_end_row,
            footer_start_row=footer_start_row_template
        )
        
        # 3b. Restore ONLY header to output worksheet
        print(f"[LayoutBuilder] Restoring header from template to output worksheet")
        self.template_state_builder.restore_header_only(target_worksheet=self.worksheet)
        
        # 4. Header Builder - writes header data to NEW worksheet
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

        # 5. Data Table Builder (writes data rows, returns footer position)
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

        if not fill_success:
            print(f"Failed to fill table data for sheet '{self.sheet_name}'.")
            return False
        
        # 6. Footer Builder (proper Director pattern - called explicitly by LayoutBuilder)
        # Prepare footer parameters
        pallet_count = 0
        if data_source_type == "processed_tables":
            pallet_count = local_chunk_pallets
        else:
            pallet_count = self.final_grand_total_pallets

        # Get footer config and sum ranges
        footer_config = self.sheet_config.get('footer_configurations', {})
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
        
        # 7. Template Footer Restoration
        # Restore the template footer (static content like "Manufacture:", etc.) AFTER the dynamic footer
        # This places the template footer below the data footer
        write_pointer_row = self.next_row_after_footer  # Next available row after dynamic footer
        
        print(f"[LayoutBuilder] Restoring template footer after row {write_pointer_row}")
        self.template_state_builder.restore_footer_only(
            target_worksheet=self.worksheet,  # Write to output worksheet
            footer_start_row=write_pointer_row
        )
        
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
