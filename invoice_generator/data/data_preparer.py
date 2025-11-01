from typing import Any, Union, Dict, List, Tuple
from decimal import Decimal

def parse_mapping_rules(
    mapping_rules: Dict[str, Any],
    column_id_map: Dict[str, int],
    idx_to_header_map: Dict[int, str]
) -> Dict[str, Any]:
    """
    Parses the mapping rules from a standardized, ID-based configuration.

    This function is refined to handle different mapping structures, such as a
    flat structure for aggregation sheets and a nested 'data_map' for table-based sheets.

    Args:
        mapping_rules: The raw mapping rules dictionary from the sheet's configuration.
        column_id_map: A dictionary mapping column IDs to their 1-based column index.
        idx_to_header_map: A dictionary mapping a column index back to its header text.

    Returns:
        A dictionary containing all the parsed information required for data filling.
    """
    # --- Initialize all return values ---
    parsed_result = {
        "static_value_map": {},
        "initial_static_col1_values": [],
        "dynamic_mapping_rules": {},
        "formula_rules": {},
        "col1_index": -1,
        "num_static_labels": 0,
        "static_column_header_name": None,
        "apply_special_border_rule": False
    }

    # --- Process all rules in a single, intelligent pass ---
    for rule_key, rule_value in mapping_rules.items():
        if not isinstance(rule_value, dict):
            continue # Skip non-dictionary rules

        # --- Handler for nested 'data_map' (used by 'processed_tables_multi') ---
        if rule_key == "data_map":
            # The entire dictionary under "data_map" is our set of dynamic rules.
            parsed_result["dynamic_mapping_rules"].update(rule_value)
            continue

        rule_type = rule_value.get("type")

        # --- Handler for Initial Static Rows ---
        if rule_type == "initial_static_rows":
            static_column_id = rule_value.get("column_header_id")
            target_col_idx = column_id_map.get(static_column_id)

            if target_col_idx:
                parsed_result["static_column_header_name"] = idx_to_header_map.get(target_col_idx)
                parsed_result["col1_index"] = target_col_idx
                parsed_result["initial_static_col1_values"] = rule_value.get("values", [])
                parsed_result["num_static_labels"] = len(parsed_result["initial_static_col1_values"])
                
                header_text = parsed_result["static_column_header_name"]
                parsed_result["apply_special_border_rule"] = header_text and header_text.strip() in ["Mark & Nº", "Mark & N °"]
            else:
                print(f"Warning: Initial static rows column with ID '{static_column_id}' not found.")
            continue

        # For all other rules, get the target column index using the RELIABLE ID
        target_id = rule_value.get("id")
        target_col_idx = column_id_map.get(target_id)

        # --- Handler for Formulas ---
        if rule_type == "formula":
            if target_col_idx:
                parsed_result["formula_rules"][target_col_idx] = {
                    "template": rule_value.get("formula_template"),
                    "input_ids": rule_value.get("inputs", [])
                }
            else:
                print(f"Warning: Could not find target column for formula rule with id '{target_id}'.")

        # --- Handler for Static Values ---
        elif "static_value" in rule_value:
            if target_col_idx:
                parsed_result["static_value_map"][target_col_idx] = rule_value["static_value"]
            else:
                print(f"Warning: Could not find target column for static_value rule with id '{target_id}'.")
        
        # --- Handler for top-level Dynamic Rules (used by 'aggregation') ---
        else:
            # If it's not a special rule, it's a dynamic mapping rule for the aggregation data type.
            parsed_result["dynamic_mapping_rules"][rule_key] = rule_value
            
    return parsed_result

def _to_numeric(value: Any) -> Union[int, float, None, Any]:
    """
    Safely attempts to convert a value to a float or int.
    Handles strings with commas and returns the original value on failure.
    """
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, str):
        try:
            cleaned_val = value.replace(',', '').strip()
            if not cleaned_val:
                return None
            return float(cleaned_val) if '.' in cleaned_val else int(cleaned_val)
        except (ValueError, TypeError):
            return value # Return original string if conversion fails
    if isinstance(value, Decimal):
        return float(value)
    return value # Return original value for other types

def _apply_fallback(
    row_dict: Dict[int, Any],
    target_col_idx: int,
    mapping_rule: Dict[str, Any],
    DAF_mode: bool
):
    """
    Applies a fallback value to the row_dict based on the DAF_mode.
    """
    if DAF_mode:
        fallback_key = "fallback_on_DAF"
    else:
        fallback_key = "fallback_on_none"
        
    row_dict[target_col_idx] = mapping_rule.get(fallback_key) or mapping_rule.get("fallback_on_none")

def prepare_data_rows(
    data_source_type: str,
    data_source: Union[Dict, List],
    dynamic_mapping_rules: Dict[str, Any],
    column_id_map: Dict[str, int],
    idx_to_header_map: Dict[int, str],
    desc_col_idx: int,
    num_static_labels: int,
    static_value_map: Dict[int, Any],
    DAF_mode: bool,
) -> Tuple[List[Dict[int, Any]], List[int], bool, int]:
    """
    Corrected version with typo fix and improved fallback flexibility.
    """
    data_rows_prepared = []
    pallet_counts_for_rows = []
    num_data_rows_from_source = 0
    dynamic_desc_used = False
    
    NUMERIC_IDS = {"col_qty_pcs", "col_qty_sf", "col_unit_price", "col_amount", "col_net", "col_gross", "col_cbm"}

    # --- Handler for DAF Aggregation (Uses new fallback logic) ---
    if data_source_type == 'DAF_aggregation':
        DAF_data = data_source or {}
        num_data_rows_from_source = len(DAF_data)
        id_to_data_key_map = {"col_po": "combined_po", "col_item": "combined_item", "col_desc": "combined_description", "col_qty_sf": "total_sqft", "col_amount": "total_amount"}
        price_col_idx = column_id_map.get("col_unit_price")
        
        for row_key in sorted(DAF_data.keys()):
            row_value_dict = DAF_data.get(row_key, {})
            row_dict = {}
            for col_id, data_key in id_to_data_key_map.items():
                target_col_idx = column_id_map.get(col_id)
                if not target_col_idx: continue

                data_value = row_value_dict.get(data_key)
                is_empty = data_value is None or (isinstance(data_value, str) and not data_value.strip())

                if not is_empty:
                    row_dict[target_col_idx] = _to_numeric(data_value)
                    if col_id == "col_desc":
                        dynamic_desc_used = True
                else:
                    mapping_rule_for_id = {}
                    for rule in dynamic_mapping_rules.values():
                        if rule.get("id") == col_id:
                            mapping_rule_for_id = rule
                            break
                    _apply_fallback(row_dict, target_col_idx, mapping_rule_for_id, DAF_mode)

            if price_col_idx:
                row_dict[price_col_idx] = {"type": "formula", "template": "{col_ref_1}{row}/{col_ref_0}{row}", "inputs": ["col_qty_sf", "col_amount"]}
            data_rows_prepared.append(row_dict)

    # --- Handler for Custom Aggregation (DAF Check and Fallback Added) ---
    elif data_source_type == 'custom_aggregation':
        custom_data = data_source or {}
        num_data_rows_from_source = len(custom_data)
        price_col_idx = column_id_map.get("col_unit_price")
        desc_col_idx_local = column_id_map.get("col_desc")

        for key_tuple, value_dict in custom_data.items():
            if not isinstance(key_tuple, tuple) or len(key_tuple) < 4: continue
            
            row_dict = {}
            # Directly map known values first
            row_dict[column_id_map.get("col_po")] = key_tuple[0]
            row_dict[column_id_map.get("col_item")] = key_tuple[1]
            row_dict[column_id_map.get("col_qty_sf")] = _to_numeric(value_dict.get("sqft_sum"))
            row_dict[column_id_map.get("col_amount")] = _to_numeric(value_dict.get("amount_sum"))

            if desc_col_idx_local:
                desc_value = key_tuple[3]
                if desc_value:
                    row_dict[desc_col_idx_local] = desc_value
                    dynamic_desc_used = True
            
            # Apply fallbacks for any unmapped columns based on DAF_mode
            for header, mapping_rule in dynamic_mapping_rules.items():
                target_id = mapping_rule.get("id")
                target_col_idx = column_id_map.get(target_id)
                if not target_col_idx or target_col_idx in row_dict:
                    continue

                _apply_fallback(row_dict, target_col_idx, mapping_rule, DAF_mode)

            if price_col_idx:
                row_dict[price_col_idx] = {"type": "formula", "template": "{col_ref_1}{row}/{col_ref_0}{row}", "inputs": ["col_qty_sf", "col_amount"]}
            
            data_rows_prepared.append({k: v for k, v in row_dict.items() if k is not None})

    # --- Unified Handler for Aggregation & Processed Tables (TYPO FIXED) ---
    else:
        normalized_data = []
        if data_source_type == 'aggregation':
            aggregation_data = data_source or {}
            num_data_rows_from_source = len(aggregation_data)
            for key_tuple, value_dict in aggregation_data.items():
                normalized_data.append({'key_tuple': key_tuple, 'value_dict': value_dict})

        elif data_source_type in ['processed_tables', 'processed_tables_multi']:
            table_data = data_source or {}
            if isinstance(table_data, dict):
                max_len = max((len(v) for v in table_data.values() if isinstance(v, list)), default=0)
                num_data_rows_from_source = max_len
                raw_pallet_counts = table_data.get("pallet_count", [])
                pallet_counts_for_rows = raw_pallet_counts[:max_len] + [0] * (max_len - len(raw_pallet_counts)) if isinstance(raw_pallet_counts, list) else [0] * max_len
                for i in range(max_len):
                    normalized_data.append({'table_row_index': i, 'table_data': table_data})
        
        for item in normalized_data:
            row_dict = {}
            for header, mapping_rule in dynamic_mapping_rules.items():
                target_id = mapping_rule.get("id")
                target_col_idx = column_id_map.get(target_id)
                if not target_col_idx: continue

                data_value = None
                if 'key_tuple' in item:
                    key_tuple, value_dict = item['key_tuple'], item['value_dict']
                    if 'key_index' in mapping_rule and mapping_rule['key_index'] < len(key_tuple):
                        data_value = key_tuple[mapping_rule['key_index']]
                    elif 'value_key' in mapping_rule:
                        data_value = value_dict.get(mapping_rule['value_key'])
                elif 'table_row_index' in item:
                    i, table_data = item['table_row_index'], item['table_data']
                    source_list = table_data.get(header, [])
                    if i < len(source_list):
                        data_value = source_list[i]

                is_empty = data_value is None or (isinstance(data_value, str) and not data_value.strip())
                
                if not is_empty:
                    if target_id in NUMERIC_IDS: data_value = _to_numeric(data_value)
                    row_dict[target_col_idx] = data_value
                    if target_id == 'col_desc':
                        dynamic_desc_used = True
                else:
                    _apply_fallback(row_dict, target_col_idx, mapping_rule, DAF_mode)
            
            if data_source_type == 'aggregation':
                amount_col_idx = column_id_map.get("col_amount")
                if amount_col_idx:
                    row_dict[amount_col_idx] = {"type": "formula", "template": "{col_ref_1}{row}*{col_ref_0}{row}", "inputs": ["col_qty_sf", "col_unit_price"]}
            
            data_rows_prepared.append(row_dict)

    # --- Final Processing Steps (Unchanged) ---
    if static_value_map:
        for row_data in data_rows_prepared:
            for col_idx, static_val in static_value_map.items():
                if col_idx not in row_data:
                    row_data[col_idx] = static_val
    
    if num_static_labels > len(data_rows_prepared):
        data_rows_prepared.extend([{}] * (num_static_labels - len(data_rows_prepared)))
    
    return data_rows_prepared, pallet_counts_for_rows, dynamic_desc_used, num_data_rows_from_source