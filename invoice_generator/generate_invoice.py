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

# --- Import utility functions from the new structure ---
from . import invoice_utils
from . import merge_utils
from . import text_replace_utils
from .processors.single_table_processor import SingleTableProcessor
from .processors.multi_table_processor import MultiTableProcessor

# --- Helper Functions (derive_paths, load_config, load_data) ---
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

def load_config(config_path: Path) -> Optional[Dict[str, Any]]:
    """Loads and parses the JSON configuration file."""
    print(f"Loading configuration from: {config_path}")
    try:
        with open(config_path, 'r', encoding='utf-8') as f: config_data = json.load(f)
        print("Configuration loaded successfully.")
        return config_data
    except Exception as e:
        print(f"Error loading configuration file {config_path}: {e}"); traceback.print_exc(); return None

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
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(paths['template'], output_path)
    except Exception as e:
        print(f"Error copying template: {e}"); sys.exit(1)
    print(f"Template copied successfully to {output_path}")

    print("\n4. Processing workbook...")
    workbook = None
    processing_successful = True

    try:
        workbook = openpyxl.load_workbook(output_path)

        sheets_to_process_config = config.get('sheets_to_process', [])
        sheets_to_process = [s for s in sheets_to_process_config if s in workbook.sheetnames]

        if not sheets_to_process:
            print("Error: No valid sheets to process."); sys.exit(1)

        if args.DAF:
            text_replace_utils.run_DAF_specific_replacement_task(workbook=workbook)
        
        text_replace_utils.run_invoice_header_replacement_task(workbook, invoice_data)

        original_merges = merge_utils.store_original_merges(workbook, sheets_to_process)
        sheet_data_source = config.get('sheet_data_map', {})
        data_mapping_config = config.get('data_mapping', {})

        # Global pallet calculation remains the same
        final_grand_total_pallets = 0
        processed_tables_data_for_calc = invoice_data.get('processed_tables_data', {})
        if isinstance(processed_tables_data_for_calc, dict):
            # Simplified calculation
            final_grand_total_pallets = sum(int(c) for t in processed_tables_data_for_calc.values() for c in t.get("pallet_count", []) if str(c).isdigit())
        print(f"DEBUG: Globally calculated final grand total pallets: {final_grand_total_pallets}")

        # --- REFACTORED Main Processing Loop ---
        for sheet_name in sheets_to_process:
            print(f"\n--- Processing Sheet: '{sheet_name}' ---")
            if sheet_name not in workbook.sheetnames:
                print(f"Warning: Sheet '{sheet_name}' not found. Skipping.")
                continue
            
            worksheet = workbook[sheet_name]
            sheet_config = data_mapping_config.get(sheet_name, {})
            data_source_indicator = sheet_data_source.get(sheet_name)

            if not sheet_config or not data_source_indicator:
                print(f"Warning: No config for sheet '{sheet_name}'. Skipping.")
                continue

            # --- Processor Factory ---
            processor = None
            if data_source_indicator == "processed_tables_multi":
                processor = MultiTableProcessor(
                    workbook, worksheet, sheet_name, sheet_config, data_mapping_config, 
                    data_source_indicator, invoice_data, args, final_grand_total_pallets
                )
            else: # Default to single table processor
                processor = SingleTableProcessor(
                    workbook, worksheet, sheet_name, sheet_config, data_mapping_config, 
                    data_source_indicator, invoice_data, args, final_grand_total_pallets
                )
            
            # --- Execute Processing ---
            if processor:
                processing_successful = processor.process()
                if not processing_successful:
                    print(f"--- ERROR occurred while processing sheet '{sheet_name}'. Halting. ---")
                    break # Stop on first error
            else:
                print(f"Warning: No suitable processor found for sheet '{sheet_name}'. Skipping.")

        # --- End of Loop ---

        merge_utils.find_and_restore_merges_heuristic(workbook, original_merges, sheets_to_process)

        print("\n--------------------------------")
        if processing_successful:
            print("5. Saving final workbook...")
            workbook.save(output_path)
            print(f"--- Workbook saved successfully: '{output_path}' ---")
        else:
            print("--- Processing completed with errors. Saving workbook (may be incomplete). ---")
            workbook.save(output_path)

    except Exception as e:
        print(f"\n--- UNHANDLED ERROR: {e} ---"); traceback.print_exc()
    finally:
        if workbook:
            try: workbook.close(); print("Workbook closed.")
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
