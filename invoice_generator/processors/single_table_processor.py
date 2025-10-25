print(f"Loading module: {__file__}")
from typing import Any, Dict
from ..builders.header_builder import HeaderBuilder
from ..builders.layout_builder import LayoutBuilder
from .base_processor import BaseProcessor

class SingleTableProcessor(BaseProcessor):
    def process(self) -> bool:
        layout_builder = LayoutBuilder(
            worksheet=self.worksheet,
            sheet_name=self.sheet_name,
            sheet_config=self.sheet_config,
            all_sheet_configs=self.data_mapping_config,
            invoice_data=self.invoice_data,
            styling_config=self.styling_config,
            args=self.args,
            final_grand_total_pallets=self.final_grand_total_pallets,
        )

        if not layout_builder.build():
            print(f"Failed to build layout for sheet '{self.sheet_name}'.")
            return False

        self.run_text_replacement()

        return True
