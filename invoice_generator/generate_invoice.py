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
from .builders.text_replacement_builder import TextReplacementBuilder
from .builders.workbook_builder import WorkbookBuilder
from .processors.single_table_processor import SingleTableProcessor
from .processors.multi_table_processor import MultiTableProcessor
from .config.loader import BundledConfigLoader

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
        
        # Check for bundled config in subdirectory structure
        exact_bundled_config_subdir = config_dir / f"{template_name_part}_config" / f"{template_name_part}_config.json"
        exact_bundled_config_v2 = config_dir / f"{template_name_part}_bundled_v2.json"
        
        print(f"Checking for exact match: Template='{exact_template_path}', Config='{exact_config_path}'")

        # Check for bundled config first (preferred) - try subdirectory structure first
        if exact_template_path.is_file() and exact_bundled_config_subdir.is_file():
            print(f"Found exact match for template and bundled config (subdir): {exact_bundled_config_subdir}")
            return {"data": input_data_path, "template": exact_template_path, "config": exact_bundled_config_subdir}
        elif exact_template_path.is_file() and exact_bundled_config_v2.is_file():
            print(f"Found exact match for template and bundled config (v2): {exact_bundled_config_v2}")
            return {"data": input_data_path, "template": exact_template_path, "config": exact_bundled_config_v2}
        elif exact_template_path.is_file() and exact_config_path.is_file():
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
                
                # Check for bundled config in subdirectory structure
                prefix_bundled_config_subdir = config_dir / f"{prefix}_config" / f"{prefix}_config.json"
                prefix_bundled_config_v2 = config_dir / f"{prefix}_bundled_v2.json"
                print(f"Checking for prefix match: Template='{prefix_template_path}', Config='{prefix_config_path}'")

                # Check bundled config first - try subdirectory structure first
                if prefix_template_path.is_file() and prefix_bundled_config_subdir.is_file():
                    print(f"Found prefix match for template and bundled config (subdir): {prefix_bundled_config_subdir}")
                    return {"data": input_data_path, "template": prefix_template_path, "config": prefix_bundled_config_subdir}
                elif prefix_template_path.is_file() and prefix_bundled_config_v2.is_file():
                    print(f"Found prefix match for template and bundled config (v2): {prefix_bundled_config_v2}")
                    return {"data": input_data_path, "template": prefix_template_path, "config": prefix_bundled_config_v2}
                elif prefix_template_path.is_file() and prefix_config_path.is_file():
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
        
        # Check if this is a bundled config (has _meta and config_version)
        if '_meta' in config_data and 'config_version' in config_data['_meta']:
            print(f"Detected bundled config version: {config_data['_meta']['config_version']}")
            config_data['_is_bundled'] = True
        else:
            print("Detected legacy config format")
            config_data['_is_bundled'] = False
        
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
    print(f"Started at: {time.strftime('%H:%M:%S', time.localtime(start_time))}")

    paths = derive_paths(args.input_data_file, args.templatedir, args.configdir)
    if not paths: sys.exit(1)

    config = load_config(paths['config'])
    invoice_data = load_data(paths['data'])
    if not config or not invoice_data: sys.exit(1)

    output_path = Path(args.output).resolve()
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        print(f"Error creating output directory: {e}"); sys.exit(1)
    
    # NOTE: We no longer copy the template file directly.
    # Instead, we'll load it as read-only and create a new workbook separately.
    print("Template will be loaded as read-only for state capture")

    print("\n4. Loading template and creating new workbook...")
    template_workbook = None
    output_workbook = None
    processing_successful = True

    try:
        # Step 1: Load template workbook as read_only=False
        # - read_only=False: Allows access to merged_cells for state capture
        # - data_only=False (default): Preserves formulas and text with '=' signs
        print(f"Loading template from: {paths['template']}")
        template_workbook = openpyxl.load_workbook(
            paths['template'], 
            read_only=False
        )
        print(f"Template loaded successfully (read_only=False)")
        print(f"   Template sheets: {template_workbook.sheetnames}")
        
        # Step 2: Collect all sheet names from template
        template_sheet_names = template_workbook.sheetnames
        print(f"Found {len(template_sheet_names)} sheets in template")
        
        # Step 3: Create WorkbookBuilder with template sheet names
        print("Creating new output workbook using WorkbookBuilder...")
        workbook_builder = WorkbookBuilder(sheet_names=template_sheet_names)
        
        # Step 4: Build the new clean workbook
        output_workbook = workbook_builder.build()
        print(f"New output workbook created with {len(output_workbook.sheetnames)} sheets")
        
        # Step 5: Store references to both workbooks
        # - template_workbook: Used for reading template state (READ-ONLY usage)
        # - output_workbook: Used for writing final output (WRITABLE)
        workbook = output_workbook  # Keep 'workbook' name for compatibility with rest of code
        print("Both template (read) and output (write) workbooks ready")

        # Detect config type and setup appropriate accessors
        is_bundled = config.get('_is_bundled', False)
        config_loader = None
        
        if is_bundled:
            print("\n--- Using Bundled Config Format (Direct) ---")
            config_loader = BundledConfigLoader(config)
            sheets_to_process_config = config_loader.get_sheets_to_process()
            sheet_data_map = config_loader.get_sheet_data_map()
            
            # NO CONVERSION - Pass config_loader to processors directly
            data_mapping_config = {}  # Empty, processors will use config_loader instead
            
            print(f"Customer: {config_loader.customer}")
            print(f"Sheets to process: {sheets_to_process_config}")
        else:
            print("\n--- Using Legacy Config Format ---")
            sheets_to_process_config = config.get('sheets_to_process', [])
            sheet_data_map = config.get('sheet_data_map', {})
            data_mapping_config = config.get('data_mapping', {})
        
        sheets_to_process = [s for s in sheets_to_process_config if s in workbook.sheetnames]

        if not sheets_to_process:
            print("Error: No valid sheets to process."); sys.exit(1)

        # Use TextReplacementBuilder instead of old utils
        if args.DAF:
            text_replacer = TextReplacementBuilder(workbook=workbook, invoice_data=invoice_data)
            text_replacer.build()
        else:
            # Still run header replacement for non-DAF mode
            text_replacer = TextReplacementBuilder(workbook=workbook, invoice_data=invoice_data)
            text_replacer._replace_placeholders()  # Only run placeholder replacement

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
            
            # Get both template and output worksheets
            template_worksheet = template_workbook[sheet_name]
            output_worksheet = workbook[sheet_name]
            
            # Get sheet config - from loader if bundled, else from legacy dict
            if config_loader:
                sheet_config = {}  # Processor will use config_loader directly
                data_source_indicator = config_loader.get_data_source(sheet_name)
            else:
                sheet_config = data_mapping_config.get(sheet_name, {})
                data_source_indicator = sheet_data_map.get(sheet_name)

            # Validate - if not bundled, need sheet_config; always need data_source
            if not data_source_indicator:
                print(f"Warning: No data source specified for sheet '{sheet_name}'. Skipping.")
                continue
            if not config_loader and not sheet_config:
                print(f"Warning: No config for sheet '{sheet_name}'. Skipping.")
                continue

            # --- Processor Factory ---
            processor = None
            if data_source_indicator in ["processed_tables_multi", "processed_tables"]:
                processor = MultiTableProcessor(
                    template_workbook=template_workbook,
                    output_workbook=output_workbook,
                    template_worksheet=template_worksheet,
                    output_worksheet=output_worksheet,
                    sheet_name=sheet_name,
                    sheet_config=sheet_config,
                    data_mapping_config=data_mapping_config,
                    data_source_indicator=data_source_indicator,
                    invoice_data=invoice_data,
                    cli_args=args,
                    final_grand_total_pallets=final_grand_total_pallets,
                    config_loader=config_loader
                )
            else: # Default to single table processor
                processor = SingleTableProcessor(
                    template_workbook=template_workbook,
                    output_workbook=output_workbook,
                    template_worksheet=template_worksheet,
                    output_worksheet=output_worksheet,
                    sheet_name=sheet_name,
                    sheet_config=sheet_config,
                    data_mapping_config=data_mapping_config,
                    data_source_indicator=data_source_indicator,
                    invoice_data=invoice_data,
                    cli_args=args,
                    final_grand_total_pallets=final_grand_total_pallets,
                    config_loader=config_loader
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

        print("\n--------------------------------")
        if processing_successful:
            print("5. Saving final workbook...")
            output_workbook.save(output_path)
            print(f"--- Workbook saved successfully: '{output_path}' ---")
        else:
            print("--- Processing completed with errors. Saving workbook (may be incomplete). ---")
            output_workbook.save(output_path)

    except Exception as e:
        print(f"\n--- UNHANDLED ERROR: {e} ---"); traceback.print_exc()
    finally:
        # Close both workbooks
        if template_workbook:
            try: template_workbook.close(); print("Template workbook closed.")
            except Exception: pass
        if output_workbook:
            try: output_workbook.close(); print("Output workbook closed.")
            except Exception: pass

    total_time = time.time() - start_time
    print("\n--- Invoice Generation Finished ---")
    print(f"Total Time: {total_time:.2f} seconds")
    print(f"Completed at: {time.strftime('%H:%M:%S', time.localtime())}")

if __name__ == "__main__":
    # To run this script directly, you might need to adjust Python's path
    # to recognize the 'invoice_generator' package, e.g., by running from the parent directory:
    # python -m invoice_generator.generate_invoice ...
    main()
