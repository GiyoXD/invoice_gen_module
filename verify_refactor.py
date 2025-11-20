import sys
import os
import logging
from openpyxl import Workbook
from invoice_generator.builders.data_table_builder import DataTableBuilderStyler
from invoice_generator.builders.footer_builder import FooterBuilderStyler
from invoice_generator.styling.models import FooterData, StylingConfigModel

# Setup logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

def test_refactor():
    wb = Workbook()
    ws = wb.active
    
    # Mock Header Info
    header_info = {
        'second_row_index': 1,
        'num_columns': 5,
        'column_id_map': {
            'col_desc': 2,
            'col_qty': 3,
            'col_net_weight': 4,
            'col_gross_weight': 5
        }
    }
    
    # Mock Data
    resolved_data = {
        'data_source': [
            {2: 'Item 1', 3: 10, 4: 100.5, 5: 110.0},
            {2: 'Item 2', 3: 20, 4: 200.5, 5: 220.0}
        ],
        'data_source_type': 'list',
        'header_info': header_info,
        'mapping_rules': {}
    }
    
    # 1. Test DataTableBuilder
    print("\n--- Testing DataTableBuilder ---")
    dtb = DataTableBuilderStyler(
        worksheet=ws,
        header_info=header_info,
        resolved_data=resolved_data
    )
    
    footer_data = dtb.build()
    
    if isinstance(footer_data, FooterData):
        print("SUCCESS: DataTableBuilder returned FooterData")
        print(f"  Total Pallets: {footer_data.total_pallets}")
        print(f"  Weight Summary: {footer_data.weight_summary}")
        print(f"  Data Range: {footer_data.data_start_row}-{footer_data.data_end_row}")
    else:
        print("FAILURE: DataTableBuilder did not return FooterData")
        return

    # 2. Test FooterBuilder
    print("\n--- Testing FooterBuilder ---")
    
    footer_config = {
        "type": "regular",
        "total_text_column_id": "col_desc",
        "pallet_count_column_id": "col_qty",
        "add_ons": {
            "weight_summary": {
                "enabled": True,
                "label_col_id": "col_desc",
                "value_col_id": "col_net_weight"
            }
        }
    }
    
    # Mock Bundled Style Config
    style_config = {
        'styling_config': {
            'columns': {
                'col_desc': {'font': {'name': 'Arial', 'size': 10}, 'alignment': {'horizontal': 'left'}},
                'col_qty': {'font': {'name': 'Arial', 'size': 10}, 'alignment': {'horizontal': 'center'}},
                'col_net_weight': {'font': {'name': 'Arial', 'size': 10}, 'alignment': {'horizontal': 'right'}},
                'col_gross_weight': {'font': {'name': 'Arial', 'size': 10}, 'alignment': {'horizontal': 'right'}},
            },
            'row_contexts': {
                'footer': {'row_height': 20},
                'header': {'row_height': 25}
            }
        }
    }
    
    fb = FooterBuilderStyler(
        worksheet=ws,
        footer_data=footer_data,
        style_config=style_config,
        context_config={'header_info': header_info},
        data_config={'footer_config': footer_config}
    )
    
    next_row = fb.build()
    print(f"SUCCESS: FooterBuilder completed. Next row: {next_row}")
    
    # Verify Output
    print("\n--- Verifying Output ---")
    wb.save("test_refactor_output.xlsx")
    print("Output saved to test_refactor_output.xlsx")

if __name__ == "__main__":
    test_refactor()
