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

    def run_text_replacement(self):
        """
        Runs the text replacement builder.
        """
        text_replacement_builder = TextReplacementBuilder(self.workbook, self.invoice_data)
        text_replacement_builder.build()

    def process(self) -> bool:
        raise NotImplementedError("The process method must be implemented by a subclass.")
