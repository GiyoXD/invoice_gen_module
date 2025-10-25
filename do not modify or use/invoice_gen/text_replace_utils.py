import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell import Cell
from typing import List, Dict, Optional, Any
import re
import datetime

# The python-dateutil library is required for advanced date parsing.
# Install it using: pip install python-dateutil
from dateutil.parser import parse, ParserError


# ==============================================================================
# SECTION 1: CORE HELPER FUNCTIONS (WITH UPGRADED DATE HANDLING)
# ==============================================================================

def excel_number_to_datetime(excel_num: Any) -> Optional[datetime.datetime]:
    """Converts an Excel date number to a Python datetime object."""
    try:
        excel_num = float(excel_num)
        # Excel's 1900 leap year bug needs to be accounted for.
        if excel_num > 59:
            excel_num -= 1
        delta = datetime.timedelta(days=excel_num - 1)
        return datetime.datetime(1900, 1, 1) + delta
    except (ValueError, TypeError):
        return None

def format_cell_as_date_smarter(cell: Cell, value: Any):
    """
    Intelligently parses a value (string, number, or datetime) into a
    datetime object and formats the cell accordingly.
    
    This function REPLACES the old `is_date_string` and `format_cell_as_date`.
    """
    parsed_date = None

    # 1. If it's already a datetime object
    if isinstance(value, (datetime.datetime, datetime.date)):
        parsed_date = value

    # 2. If it's a string, try to parse it
    elif isinstance(value, str):
        if not value.strip(): # Skip empty strings
            pass
        else:
            try:
                # dateutil.parser handles many formats like "2025-05-11T00:00:00"
                # and "10/11/2025". 'dayfirst=True' correctly handles DD/MM/YYYY.
                parsed_date = parse(value, dayfirst=True)
            except (ParserError, ValueError):
                # The string is not a recognizable date, so it will be treated as text.
                pass

    # 3. If it's a number, it could be an Excel serial date
    elif isinstance(value, (int, float)):
        if value >= 1: # Plausible Excel dates are positive numbers.
            parsed_date = excel_number_to_datetime(value)

    # 4. If we successfully found and parsed a date, format the cell
    if parsed_date:
        cell.value = parsed_date
        # This number format tells Excel to display the date as dd/mm/yyyy
        cell.number_format = "dd/mm/yyyy"
    else:
        # If no date could be parsed, just set the cell to the original value
        cell.value = value

def _get_nested_data(data_dict: Dict[str, Any], path: List[Any]) -> Optional[Any]:
    """Safely retrieves a value from a nested structure of dictionaries and lists."""
    current_level = data_dict
    for key in path:
        if isinstance(current_level, dict) and key in current_level:
            current_level = current_level[key]
        elif isinstance(current_level, list):
            try:
                index = int(key)
                if 0 <= index < len(current_level):
                    current_level = current_level[index]
                else: return None
            except (ValueError, TypeError): return None
        else:
            return None
    return current_level

# ==============================================================================
# SECTION 2: THE ONE REPLACEMENT ENGINE (UPDATED TO USE SMARTER DATE FUNCTION)
# ==============================================================================

def find_and_replace(
    workbook: openpyxl.Workbook,
    rules: List[Dict[str, Any]],
    limit_rows: int,
    limit_cols: int,
    invoice_data: Optional[Dict[str, Any]] = None
):
    """
    A two-pass engine that handles 'exact', 'substring', and formula-based replacements.
    Pass 1: Locates all placeholders and performs simple value replacements.
    Pass 2: Uses the locations found in Pass 1 to build and apply formulas.
    """
    print(f"\n--- Starting Find and Replace on sheets (Searching Range up to row {limit_rows}, col {limit_cols}) ---")
    
    # NEW: A dictionary to store the cell coordinates of each placeholder.
    placeholder_locations: Dict[str, str] = {}
    
    # NEW: Separate rules for formulas vs. simple replacements.
    simple_rules = [r for r in rules if "formula_template" not in r]
    formula_rules = [r for r in rules if "formula_template" in r]

    for sheet in workbook.worksheets:
        if sheet.sheet_state != 'visible':
            continue

        # --- PASS 1: Find all placeholder locations and apply simple replacements ---
        print("  PASS 1: Locating placeholders and applying simple value replacements...")
        for row in sheet.iter_rows(max_row=limit_rows, max_col=limit_cols):
            for cell in row:
                if not isinstance(cell.value, str) or not cell.value:
                    continue

                # First, find and store the location of ANY placeholder
                for rule in rules:
                    if rule.get("find") == cell.value.strip():
                        placeholder_locations[rule["find"]] = cell.coordinate
                        break # Found the placeholder, no need to check other rules for this cell

                # Second, apply SIMPLE replacement rules
                for rule in simple_rules:
                    text_to_find = rule.get("find")
                    if not text_to_find:
                        continue
                    
                    match_mode = rule.get("match_mode", "substring")
                    is_match = (match_mode == 'exact' and cell.value.strip() == text_to_find) or \
                               (match_mode == 'substring' and text_to_find in cell.value)

                    if is_match:
                        replacement_content = None
                        if "data_path" in rule:
                            if not invoice_data: continue
                            replacement_content = _get_nested_data(invoice_data, rule["data_path"])
                        elif "replace" in rule:
                            replacement_content = rule["replace"]

                        if replacement_content is not None:
                            print(f"    -> Applying rule for '{text_to_find}' at {cell.coordinate}...")
                            if rule.get("is_date", False):
                                format_cell_as_date_smarter(cell, replacement_content)
                            elif match_mode == 'exact':
                                cell.value = replacement_content
                            elif match_mode == 'substring':
                                cell.value = cell.value.replace(str(text_to_find), str(replacement_content))
                        break

        # --- PASS 2: Build and apply formula-based replacements ---
        print("  PASS 2: Building and applying formula replacements...")
        if not formula_rules:
            print("    -> No formula rules to apply.")
        
        for rule in formula_rules:
            formula_template = rule["formula_template"]
            target_placeholder = rule["find"]
            
            # Find the cell where the formula should go
            target_cell_coord = placeholder_locations.get(target_placeholder)
            if not target_cell_coord:
                print(f"    -> WARNING: Could not find cell for formula placeholder '{target_placeholder}'. Skipping.")
                continue

            # Find all dependent placeholders (e.g., {[[NET]]}) in the template
            dependent_placeholders = re.findall(r'(\{\[\[.*?\]\]\})', formula_template)
            
            final_formula_str = formula_template
            all_deps_found = True
            
            for dep_placeholder in dependent_placeholders:
                # Strip the curly braces to get the actual placeholder key (e.g., [[NET]])
                dep_key = dep_placeholder.strip('{}')
                # Get the cell address for the dependency
                dep_coord = placeholder_locations.get(dep_key)
                
                if dep_coord:
                    # Replace the variable in the template with the real cell address
                    final_formula_str = final_formula_str.replace(dep_placeholder, dep_coord)
                else:
                    print(f"    -> ERROR: Could not find location for dependency '{dep_key}' needed by formula for '{target_placeholder}'.")
                    all_deps_found = False
                    break # Stop processing this formula if a dependency is missing
            
            if all_deps_found:
                # Prepend '=' to make it a valid Excel formula
                final_formula_str = f"={final_formula_str}"
                print(f"    -> SUCCESS: Placing formula '{final_formula_str}' in cell {target_cell_coord}.")
                sheet[target_cell_coord].value = final_formula_str


# ==============================================================================
# SECTION 3: TASK-RUNNER FUNCTIONS (No changes needed here)
# ==============================================================================

def run_invoice_header_replacement_task(workbook: openpyxl.Workbook, invoice_data: Dict[str, Any]):
    """Defines and runs the data-driven header replacement task."""
    print("\n--- Running Invoice Header Replacement Task (within A1:N14) ---")
    header_rules = [
        {"find": "JFINV", "data_path": ["processed_tables_data", "1", "inv_no", 0], "match_mode": "exact"},
        # This rule will now correctly handle any date format coming from your data
        {"find": "JFTIME", "data_path": ["processed_tables_data", "1", "inv_date", 0], "is_date": True, "match_mode": "exact"},
        {"find": "JFREF", "data_path": ["processed_tables_data", "1", "inv_ref", 0], "match_mode": "exact"},
        {"find": "[[CUSTOMER_NAME]]", "data_path": ["customer_info", "name"], "match_mode": "exact"},
        {"find": "[[CUSTOMER_ADDRESS]]", "data_path": ["customer_info", "address"], "match_mode": "exact"}
    ]
    find_and_replace(
        workbook=workbook,
        rules=header_rules,
        limit_rows=14,
        limit_cols=14,
        invoice_data=invoice_data
    )
    print("--- Finished Invoice Header Replacement Task ---")

def run_DAF_specific_replacement_task(workbook: openpyxl.Workbook):
    """Defines and runs the hardcoded, DAF-specific replacement task."""
    print("\n--- Running DAF-Specific Replacement Task (within 50x16 grid) ---")
    DAF_rules = [
        {"find": "BINH PHUOC", "replace": "BAVET", "match_mode": "exact"},
        {"find": "BAVET, SVAY RIENG", "replace": "BAVET", "match_mode": "exact"},
        {"find": "BAVET,SVAY RIENG", "replace": "BAVET", "match_mode": "exact"},
        {"find": "BAVET, SVAYRIENG", "replace": "BAVET", "match_mode": "exact"},
        {"find": "BINH DUONG", "replace": "BAVET", "match_mode": "exact"},
        {"find": "FCA  BAVET,SVAYRIENG", "replace": "DAF BAVET", "match_mode": "exact"},
        {"find": "FCA: BAVET,SVAYRIENG", "replace": "DAF: BAVET", "match_mode": "exact"},
        {"find": "DAF  BAVET,SVAYRIENG", "replace": "DAF BAVET", "match_mode": "exact"},
        {"find": "DAF: BAVET,SVAYRIENG", "replace": "DAF: BAVET", "match_mode": "exact"},
        {"find": "SVAY RIENG", "replace": "BAVET", "match_mode": "exact"},
        {"find": "PORT KLANG", "replace": "BAVET", "match_mode": "exact"},
        {"find": "HCM", "replace": "BAVET", "match_mode": "exact"},
        {"find": "DAP", "replace": "DAF", "match_mode": "substring"},
        {"find": "FCA", "replace": "DAF", "match_mode": "substring"},
        {"find": "CIF", "replace": "DAF", "match_mode": "substring"},
    ]
    find_and_replace(
        workbook=workbook,
        rules=DAF_rules,
        limit_rows=200,
        limit_cols=16
    )
    print("--- Finished DAF-Specific Replacement Task ---")

# ==============================================================================
# EXAMPLE USAGE (for demonstration purposes)
# ==============================================================================

if __name__ == '__main__':
    # Create a dummy workbook and data to simulate a real run
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "Invoice"

    # Set up some placeholder cells to be replaced
    sheet["A1"] = "Invoice No:"
    sheet["B1"] = "JFINV"
    sheet["A2"] = "Invoice Date:"
    sheet["B2"] = "JFTIME" # This will be replaced by the date
    sheet["A3"] = "Terms:"
    sheet["B3"] = "FCA Port"
    sheet["A4"] = "Customer:"
    sheet["B4"] = "[[CUSTOMER_NAME]]"
    sheet["A5"] = "Origin:"
    sheet["B5"] = "BINH DUONG"


    # Simulate the invoice data you would have
    mock_invoice_data = {
        "processed_tables_data": {
            "1": {
                "inv_no": ["INV-12345"],
                # This is the date format you mentioned
                "inv_date": ["2025-05-11T00:00:00"], 
                "inv_ref": ["REF-ABC"]
            }
        },
        "customer_info": {
            "name": "Global Exports Inc.",
            "address": "123 Supply Chain Rd, Commerce City"
        }
    }

    # Run the replacement tasks
    run_invoice_header_replacement_task(wb, mock_invoice_data)
    run_DAF_specific_replacement_task(wb)

    # Save the workbook to see the result
    output_filename = "invoice_output.xlsx"
    wb.save(output_filename)

    print(f"\nProcessing complete. Check the output file: {output_filename}")
    print(f"Cell B2 should now contain the date '11/05/2025' and be formatted as a date in Excel.")
    print(f"Cell B3 should now contain 'DAF Port'.")
    print(f"Cell B5 should now contain 'BAVET'.")