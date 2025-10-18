# invoice_generator/generate_invoice.py
# Main script to orchestrate invoice generation using a processor-based pattern.

import os
import json
import pickle
import argparse
import shutil
import openpyxl
import traceback
import sys
import time
from pathlib import Path
from typing import Optional, Dict, Any, List
import ast
import re

from .utils.layout import calculate_header_dimensions

# --- Import utility functions from the new structure ---
# from . import invoice_utils
from .utils import merge_utils
from . import text_replace_utils
from .processors.single_table_processor import SingleTableProcessor
from .processors.multi_table_processor import MultiTableProcessor
from .builders.template_state_builder import TemplateStateBuilder

from .config.loader import load_config, load_styling_config

# --- Helper Functions (derive_paths, load_data) ---
# These functions remain largely the same but are part of the new script.
def derive_paths(input_data_path_str: str, template_dir_str: str, config_dir_str: str) -> Optional[Dict[str, Path]]:
    """
    Derives template and config file paths based on the input data filename.
    """
    print(f"Deriving paths from input: {input_data_path_str}")
    try:
        input_data_path = Path(input_data_path_str).resolve()
        template_dir = Path(template_dir_str).resolve()
        config_dir = Path(config_dir_str).resolve()

        if not input_data_path.is_file(): print(f"Error: Input data file not found: {input_data_path}"); return None
        if not template_dir.is_dir(): print(f"Error: Template directory not found: {template_dir}"); return None
        if not config_dir.is_dir(): print(f"Error: Config directory not found: {config_dir}"); return None

        base_name = input_data_path.stem
        template_name_part = base_name
        suffixes_to_remove = ['_data', '_input', '_pkl']
        prefixes_to_remove = ['data_']

        for suffix in suffixes_to_remove:
            if base_name.lower().endswith(suffix):
                template_name_part = base_name[:-len(suffix)]
                break
        else:
            for prefix in prefixes_to_remove:
                if base_name.lower().startswith(prefix):
                    template_name_part = base_name[len(prefix):]
                    break

        if not template_name_part:
            print(f"Error: Could not derive template name part from: '{base_name}'")
            return None
        print(f"Derived initial template name part: '{template_name_part}'")

        exact_template_filename = f"{template_name_part}.xlsx"
        exact_config_filename = f"{template_name_part}_config.json"
        exact_template_path = template_dir / exact_template_filename
        exact_config_path = config_dir / exact_config_filename
        print(f"Checking for exact match: Template='{exact_template_path}', Config='{exact_config_path}'")

        if exact_template_path.is_file() and exact_config_path.is_file():
            print("Found exact match for template and config.")
            return {"data": input_data_path, "template": exact_template_path, "config": exact_config_path}
        else:
            print("Exact match not found. Attempting prefix matching...")
            prefix_match = re.match(r'^([a-zA-Z]+[-_]?[a-zA-Z]*)', template_name_part)
            if prefix_match:
                prefix = prefix_match.group(1)
                print(f"Extracted prefix: '{prefix}'")
                prefix_template_filename = f"{prefix}.xlsx"
                prefix_config_filename = f"{prefix}_config.json"
                prefix_template_path = template_dir / prefix_template_filename
                prefix_config_path = config_dir / prefix_config_filename
                print(f"Checking for prefix match: Template='{prefix_template_path}', Config='{prefix_config_path}'")

                if prefix_template_path.is_file() and prefix_config_path.is_file():
                    print("Found prefix match for template and config.")
                    return {"data": input_data_path, "template": prefix_template_path, "config": prefix_config_path}
                else:
                    print("Prefix match not found.")
            else:
                print("Could not extract a letter-based prefix.")

            print(f"Error: Could not find matching template/config files using exact ('{template_name_part}') or prefix methods.")
            if not exact_template_path.is_file(): print(f"Error: Template file not found: {exact_template_path}")
            if not exact_config_path.is_file(): print(f"Error: Configuration file not found: {exact_config_path}")
            return None

    except Exception as e:
        print(f"Error deriving file paths: {e}")
        traceback.print_exc()
        return None

from .styling.models import StylingConfigModel



def load_data(data_path: Path) -> Optional[Dict[str, Any]]:
    """ Loads and parses the input data file. Supports .json and .pkl. """
    print(f"Loading data from: {data_path}")
    invoice_data = None; file_suffix = data_path.suffix.lower()
    try:
        if file_suffix == '.json':
            with open(data_path, 'r', encoding='utf-8') as f: invoice_data = json.load(f)
        elif file_suffix == '.pkl':
            with open(data_path, 'rb') as f: invoice_data = pickle.load(f)
        else:
            print(f"Error: Unsupported data file extension: '{file_suffix}'."); return None
        
        # Key conversion logic remains the same
        for key_to_convert in ["standard_aggregation_results", "custom_aggregation_results"]:
            raw_data = invoice_data.get(key_to_convert)
            if isinstance(raw_data, dict):
                processed_data = {}
                decimal_pattern = re.compile(r"Decimal\('(-?\d*\.?\d+)'\)")
                for key_str, value_dict in raw_data.items():
                    try:
                        processed_key_str = decimal_pattern.sub(r"'\1'", key_str)
                        key_tuple = ast.literal_eval(processed_key_str)
                        processed_data[key_tuple] = value_dict
                    except Exception as e:
                        print(f"Warning: Could not convert key '{key_str}': {e}")
                invoice_data[key_to_convert] = processed_data
        
        return invoice_data
    except Exception as e:
        print(f"Error loading data file {data_path}: {e}"); traceback.print_exc(); return None

def main():
    """Main function to orchestrate invoice generation."""
    start_time = time.time()
    
    parser = argparse.ArgumentParser(description="Generate Invoice from Template and Data using configuration files.")
    parser.add_argument("input_data_file", help="Path to the input data file (.json or .pkl).")
    parser.add_argument("-o", "--output", default="result.xlsx", help="Path for the output Excel file.")
    parser.add_argument("-t", "--templatedir", default="./TEMPLATE", help="Directory containing template Excel files.")
    parser.add_argument("-c", "--configdir", default="./configs", help="Directory containing configuration JSON files.")
    parser.add_argument("--DAF", action="store_true", help="Enable DAF-specific processing.")
    parser.add_argument("--custom", action="store_true", help="Enable custom processing logic.")
    args = parser.parse_args()

    print("--- Starting Invoice Generation (Refactored) ---")
    print(f"🕒 Started at: {time.strftime('%H:%M:%S', time.localtime(start_time))}")

    paths = derive_paths(args.input_data_file, args.templatedir, args.configdir)
    if not paths: sys.exit(1)

    config = load_config(paths['config'])
    invoice_data = load_data(paths['data'])
    if not config or not invoice_data: sys.exit(1)

    output_path = Path(args.output).resolve()
    template_workbook = None
    output_workbook = None
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        template_workbook = openpyxl.load_workbook(paths['template'])
        output_workbook = openpyxl.Workbook()
        # Remove the default sheet created by openpyxl.Workbook()
        if "Sheet" in output_workbook.sheetnames: output_workbook.remove(output_workbook["Sheet"])
    except Exception as e:
        print(f"Error preparing workbooks: {e}"); sys.exit(1)
    print(f"Template workbook loaded and new output workbook created.")

    print("\n4. Processing workbook...")
    processing_successful = True

    try:
        sheets_to_process_config = config.get('sheets_to_process', [])
        sheets_to_process = [s for s in sheets_to_process_config if s in template_workbook.sheetnames]

        if not sheets_to_process:
            print("Error: No valid sheets to process."); sys.exit(1)

        # Global pallet calculation remains the same
        final_grand_total_pallets = 0
        processed_tables_data_for_calc = invoice_data.get('processed_tables_data', {})
        if isinstance(processed_tables_data_for_calc, dict):
            # Simplified calculation
            final_grand_total_pallets = sum(int(c) for t in processed_tables_data_for_calc.values() for c in t.get("pallet_count", []) if str(c).isdigit())
        print(f"DEBUG: Globally calculated final grand total pallets: {final_grand_total_pallets}")

        sheet_data_source = config.get('sheet_data_map', {})
        data_mapping_config = config.get('data_mapping', {})

        # --- REFACTORED Main Processing Loop ---
        for sheet_name in sheets_to_process:
            print(f"\n--- Processing Sheet: '{sheet_name}' ---")
            if sheet_name not in template_workbook.sheetnames:
                print(f"Warning: Sheet '{sheet_name}' not found in template. Skipping.")
                continue
            
            template_worksheet = template_workbook[sheet_name]
            output_worksheet = output_workbook.create_sheet(title=sheet_name)

            sheet_config = data_mapping_config.get(sheet_name, {})
            data_source_indicator = sheet_data_source.get(sheet_name)

            if not sheet_config or not data_source_indicator:
                print(f"Warning: No config for sheet '{sheet_name}'. Skipping.")
                continue

            # --- Unified "Finish Table First" Orchestration ---
            processor = None
            styling_config = load_styling_config(sheet_config)

            if data_source_indicator == "processed_tables_multi":
                processor = MultiTableProcessor(
                    output_workbook, output_worksheet, sheet_name, sheet_config, data_mapping_config,
                    data_source_indicator, invoice_data, args, final_grand_total_pallets, styling_config
                )
            else:
                processor = SingleTableProcessor(
                    output_workbook, output_worksheet, sheet_name, sheet_config, data_mapping_config,
                    data_source_indicator, invoice_data, args, final_grand_total_pallets, styling_config
                )

            if processor:
                template_state_builder = TemplateStateBuilder(template_worksheet)
                header_end_row = sheet_config.get('start_row', 1) - 1
                template_state_builder.capture_header(header_end_row)

                if processor.process():
                    # Dynamically find the footer by scanning within the table's column width
                    _, num_header_cols = calculate_header_dimensions(sheet_config.get('header_to_write', []))
                    max_col_to_check = max(num_header_cols, template_worksheet.max_column)

                    footer_start_row_in_template = -1
                    for row in range(sheet_config.get('start_row', 1), template_worksheet.max_row + 1):
                        if any(template_worksheet.cell(row=row, column=col).value for col in range(1, max_col_to_check + 1)):
                            footer_start_row_in_template = row
                            break
                    
                    data_end_row_in_template = footer_start_row_in_template - 1 if footer_start_row_in_template != -1 else template_worksheet.max_row
                    template_state_builder.capture_footer(data_end_row_in_template)
                    template_state_builder.restore_state(output_worksheet, sheet_config.get('start_row', 1))
                else:
                    processing_successful = False
            else:
                print(f"Warning: No suitable processor found for sheet '{sheet_name}'. Skipping.")

            if not processing_successful:
                print(f"--- ERROR occurred while processing sheet '{sheet_name}'. Halting. ---")
                break

        # --- End of Loop ---

        if args.DAF:
            text_replace_utils.run_DAF_specific_replacement_task(workbook=output_workbook)
        
        text_replace_utils.run_invoice_header_replacement_task(output_workbook, invoice_data)

        print("\n--------------------------------")
        if processing_successful:
            print("5. Saving final workbook...")
            output_workbook.save(output_path)
            print(f"--- Workbook saved successfully: '{output_path}' ---")

    except Exception as e:
        print(f"\n--- UNHANDLED ERROR: {e} ---"); traceback.print_exc()
    finally:
        if template_workbook: template_workbook.close()
        if output_workbook:
            try: output_workbook.close(); print("Output workbook closed.")
            except Exception: pass

    total_time = time.time() - start_time
    print("\n--- Invoice Generation Finished ---")
    print(f"🕒 Total Time: {total_time:.2f} seconds")
    print(f"🏁 Completed at: {time.strftime('%H:%M:%S', time.localtime())}")

if __name__ == "__main__":
    # To run this script directly, you might need to adjust Python's path
    # to recognize the 'invoice_generator' package, e.g., by running from the parent directory:
    # python -m invoice_generator.generate_invoice ...
    main()
