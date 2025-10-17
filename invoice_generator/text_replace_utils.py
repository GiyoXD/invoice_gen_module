import openpyxl
from typing import Dict, Any

# Import the core replacement engine from its new location in the utils package
from .utils.text import find_and_replace

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