from typing import Any, Dict
from ..builders.text_replacement_builder import TextReplacementBuilder

class BaseProcessor:
    def __init__(self, workbook: Any, worksheet: Any, sheet_name: str, sheet_config: Dict[str, Any], data_mapping_config: Dict[str, Any], data_source_indicator: str, invoice_data: Dict[str, Any], args: Any, final_grand_total_pallets: int, styling_config: Dict[str, Any]):
        self.workbook = workbook
        self.worksheet = worksheet
        self.sheet_name = sheet_name
        self.sheet_config = sheet_config
        self.data_mapping_config = data_mapping_config
        self.data_source_indicator = data_source_indicator
        self.invoice_data = invoice_data
        self.args = args
        self.final_grand_total_pallets = final_grand_total_pallets
        self.styling_config = styling_config

        self.dynamic_desc_used = False
        if self.data_source_indicator == 'processed_tables_multi':
            raw_data = self.invoice_data.get('processed_tables_data', {})
            table_keys = sorted(raw_data.keys())
            for table_key in table_keys:
                table_data = raw_data[table_key]
                descriptions = table_data.get("description", [])
                if any(descriptions):
                    self.dynamic_desc_used = True
                    break
        elif self.data_source_indicator == 'aggregation':
            aggregation_data = self.invoice_data.get('standard_aggregation_results', {})
            for key_tuple in aggregation_data.keys():
                # The key is a tuple, and the description is the 4th element
                if len(key_tuple) > 3 and key_tuple[3]:
                    self.dynamic_desc_used = True
                    break
        print(f"DEBUG: dynamic_desc_used set to {self.dynamic_desc_used}")

    def run_text_replacement(self):
        """
        Runs the text replacement builder.
        """
        text_replacement_builder = TextReplacementBuilder(self.workbook, self.invoice_data)
        text_replacement_builder.build()

    def process(self) -> bool:
        raise NotImplementedError("The process method must be implemented by a subclass.")
