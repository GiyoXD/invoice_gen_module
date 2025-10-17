print(f"Loading module: {__file__}")
from typing import Any, Dict
from invoice_generator.utils.writing import write_header
from invoice_generator.builders.table_builder import TableBuilder

class SingleTableProcessor:
    def __init__(self, workbook: Any, worksheet: Any, sheet_name: str, sheet_config: Dict[str, Any], data_mapping_config: Dict[str, Any], data_source_indicator: str, invoice_data: Dict[str, Any], args: Any, final_grand_total_pallets: int):
        self.workbook = workbook
        self.worksheet = worksheet
        self.sheet_name = sheet_name
        self.sheet_config = sheet_config
        self.data_mapping_config = data_mapping_config
        self.data_source_indicator = data_source_indicator
        self.invoice_data = invoice_data
        self.args = args
        self.final_grand_total_pallets = final_grand_total_pallets

    def process(self) -> bool:
        # Write the header based on the layout in the config file
        # print(f

        # --- Get flags and rules from the sheet's configuration ---
        sheet_inner_mapping_rules_dict = self.sheet_config.get('mappings', {})
        sheet_styling_config = self.sheet_config.get("styling")

        add_blank_after_hdr_flag = self.sheet_config.get("add_blank_after_header", False)
        static_content_after_hdr_dict = self.sheet_config.get("static_content_after_header", {})
        add_blank_before_ftr_flag = self.sheet_config.get("add_blank_before_footer", False)
        static_content_before_ftr_dict = self.sheet_config.get("static_content_before_footer", {})
        merge_rules_after_hdr = self.sheet_config.get("merge_rules_after_header", {})
        merge_rules_before_ftr = self.sheet_config.get("merge_rules_before_footer", {})
        merge_rules_footer = self.sheet_config.get("merge_rules_footer", {})
        data_cell_merging_rules = self.sheet_config.get("data_cell_merging_rule", None)

        # --- Get Data Source ---
        data_to_fill = None
        data_source_type = None
        print(f"Retrieving data source for '{self.sheet_name}' using indicator: '{self.data_source_indicator}'")

        # Logic to select the correct data source based on flags and config
        if self.args.custom and self.data_source_indicator == 'aggregation':
            data_to_fill = self.invoice_data.get('custom_aggregation_results')
            data_source_type = 'custom_aggregation'

        if data_to_fill is None:
            if self.args.DAF and self.sheet_name in ["Invoice", "Contract"]:
                self.data_source_indicator = 'DAF_aggregation'

            if self.data_source_indicator == 'DAF_aggregation':
                data_to_fill = self.invoice_data.get('final_DAF_compounded_result')
                data_source_type = 'DAF_aggregation'
            elif self.data_source_indicator == 'aggregation':
                data_to_fill = self.invoice_data.get('standard_aggregation_results')
                data_source_type = 'aggregation'
            elif 'processed_tables_data' in self.invoice_data and self.data_source_indicator in self.invoice_data.get('processed_tables_data', {}):
                data_to_fill = self.invoice_data['processed_tables_data'].get(self.data_source_indicator)
                data_source_type = 'processed_tables'

        if data_to_fill is None:
            print(f"Warning: Data source '{self.data_source_indicator}' unknown or data empty. Skipping fill.")
            return True

        start_row = self.sheet_config.get('start_row', 1)
        header_to_write = self.sheet_config.get('header_to_write')
        header_info = write_header(self.worksheet, start_row, header_to_write, sheet_styling_config)

        if not header_info or not header_info.get('column_map'):
            print(f"DEBUG: header_info after write_header: {header_info}")
            print(f"Error: Cannot fill data for '{self.sheet_name}' because header_info or column_map is missing.")
            return False

        # Instantiate TableBuilder
        table_builder = TableBuilder(
            worksheet=self.worksheet,
            sheet_name=self.sheet_name,
            sheet_config=self.sheet_config,
            all_sheet_configs=self.data_mapping_config,
            data_source=data_to_fill,
            data_source_type=data_source_type,
            header_info=header_info,
            mapping_rules=sheet_inner_mapping_rules_dict,
            sheet_styling_config=sheet_styling_config,
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
        )

        print(f"DEBUG: start_row: {start_row}")
        print(f"DEBUG: header_to_write: {header_to_write}")
        print(f"DEBUG: sheet_styling_config: {sheet_styling_config}")

        # Call build method
        fill_success, next_row_after_footer, _, _, _ = table_builder.build()

        if not fill_success:
            print(f"Failed to fill table data/footer for sheet '{self.sheet_name}'.")
            return False
        print(f"Successfully filled table data/footer for sheet '{self.sheet_name}'.")

        # Placeholder for other post-table processing (e.g., weight summary, column widths, final spacers)
        return True
