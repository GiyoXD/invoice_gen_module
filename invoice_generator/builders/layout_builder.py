from typing import Any, Dict, Optional
from openpyxl.worksheet.worksheet import Worksheet

from ..styling.models import StylingConfigModel
from .header_builder import HeaderBuilder
from .table_builder import TableBuilder

class LayoutBuilder:
    def __init__(
        self,
        worksheet: Worksheet,
        sheet_name: str,
        sheet_config: Dict[str, Any],
        all_sheet_configs: Dict[str, Any],
        invoice_data: Dict[str, Any],
        styling_config: Optional[StylingConfigModel] = None,
        args: Optional[Any] = None,
        final_grand_total_pallets: int = 0,
    ):
        self.worksheet = worksheet
        self.sheet_name = sheet_name
        self.sheet_config = sheet_config
        self.all_sheet_configs = all_sheet_configs
        self.invoice_data = invoice_data
        self.styling_config = styling_config
        self.args = args
        self.final_grand_total_pallets = final_grand_total_pallets

    def build(self) -> bool:
        start_row = self.sheet_config.get('start_row', 1)
        header_to_write = self.sheet_config.get('header_to_write')

        header_builder = HeaderBuilder(
            worksheet=self.worksheet,
            start_row=start_row,
            header_layout_config=header_to_write,
            sheet_styling_config=self.styling_config,
        )
        header_info = header_builder.build()

        if not header_info or not header_info.get('column_map'):
            print(f"Error: Cannot fill data for '{self.sheet_name}' because header_info or column_map is missing.")
            return False

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

        table_builder = TableBuilder(
            worksheet=self.worksheet,
            sheet_name=self.sheet_name,
            sheet_config=self.sheet_config,
            all_sheet_configs=self.all_sheet_configs,
            data_source=data_to_fill,
            data_source_type=data_source_type,
            header_info=header_info,
            mapping_rules=sheet_inner_mapping_rules_dict,
            sheet_styling_config=self.styling_config,
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

        fill_success, _, _, _, _ = table_builder.build()

        if not fill_success:
            print(f"Failed to fill table data/footer for sheet '{self.sheet_name}'.")
            return False
            
        return True
