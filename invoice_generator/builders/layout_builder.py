from typing import Any, Dict, Optional
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import Workbook

from ..styling.models import StylingConfigModel
from .header_builder import HeaderBuilder
from .data_table_builder import DataTableBuilder
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
        sheet_name: str,
        sheet_config: Dict[str, Any],
        all_sheet_configs: Dict[str, Any],
        invoice_data: Dict[str, Any],
        styling_config: Optional[StylingConfigModel] = None,
        args: Optional[Any] = None,
        final_grand_total_pallets: int = 0,
        enable_text_replacement: bool = False,
    ):
        self.workbook = workbook
        self.worksheet = worksheet
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
        """
        # 1. Text Replacement (if enabled) - Pre-processing
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
        # This needs to be calculated before HeaderBuilder to know template boundaries
        header_end_row = start_row + 1  # Assumption: 2-row header (row 0 and row 1)
        
        # Calculate footer_start_row from template
        # In templates, footer typically starts right after minimal data area
        # For Invoice/Contract: usually 3-4 rows after header
        # We'll use start_row + 2 + 1 as a reasonable estimate (header takes 2 rows, then 1+ data rows)
        footer_start_row_template = start_row + 3  # Conservative estimate
        
        # 3. Template State Capture - Capture BEFORE any modifications
        self.template_state_builder = TemplateStateBuilder(
            worksheet=self.worksheet,
            num_header_cols=num_header_cols,
            header_end_row=header_end_row,
            footer_start_row=footer_start_row_template
        )
        
        # 4. Header Builder
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

        # 5. Data Table Builder (handles data rows + footer internally)
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

        fill_success, self.next_row_after_footer, _, _, _ = data_table_builder.build()

        if not fill_success:
            print(f"Failed to fill table data/footer for sheet '{self.sheet_name}'.")
            return False
        
        # 6. Template State Restoration
        # According to issue #18: BOTH data_start_row and data_table_end_row should be set
        # to write_pointer_row (the row AFTER all dynamically generated content including footer).
        # This places the template's static footer AFTER everything, not overwriting our footer.
        write_pointer_row = self.next_row_after_footer - 1  # Last row with content
        
        # Restore template state (merges, heights, etc.)
        self.template_state_builder.restore_state(
            target_worksheet=self.worksheet,
            data_start_row=write_pointer_row,
            data_table_end_row=write_pointer_row  # Same as data_start_row per issue #18
        )
        
        return True
