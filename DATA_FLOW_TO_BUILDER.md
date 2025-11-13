# Data Flow to DataTableBuilder

## What Format DataTableBuilder Accepts

DataTableBuilder expects `resolved_data` dict with this structure:

```python
resolved_data = {
    'data_rows': [
        {1: "PO123", 2: "ITEM-001", 3: "LEATHER", 4: 100.5, 5: {"type": "formula", ...}},
        {1: "PO124", 2: "ITEM-002", 3: "SUEDE", 4: 200.0, 5: {...}},
        # ... more rows
    ],  # List of dicts where KEY = column_index, VALUE = cell_value
    
    'pallet_counts': [5, 3, 7, ...],  # List of ints (one per row)
    
    'dynamic_desc_used': True,  # Boolean
    
    'num_data_rows': 15,  # Int
    
    'static_info': {
        'col1_index': 1,
        'num_static_labels': 3,
        'initial_static_col1_values': ["VENDOR#:", "Des: LEATHER", "MADE IN CAMBODIA"],
        'static_column_header_name': 'col_static',
        'apply_special_border_rule': False
    },
    
    'formula_rules': {...},  # Dict of formula definitions
    
    'static_content': {  # NEW: Added by our fix
        'col_static': ["VENDOR#:", "Des: LEATHER", "MADE IN CAMBODIA"],
        'before_footer': {...}
    }
}
```

### Key Format Details:

**`data_rows`** - This is the critical format DataTableBuilder uses:
```python
# It's a LIST of DICT where:
# - Each dict represents ONE ROW
# - Keys are COLUMN INDICES (integers like 1, 2, 3...)
# - Values are CELL VALUES (strings, numbers, or formula dicts)

Example:
[
    {1: "PO123", 2: "ITEM-001", 3: "LEATHER", 4: 100.5},  # Row 1
    {1: "PO123", 2: "ITEM-002", 3: "SUEDE", 4: 200.0},    # Row 2
]
```

**Formula Format**:
```python
{
    "type": "formula",
    "template": "{col_ref_0}{row}*{col_ref_1}{row}",
    "inputs": ["col_qty_sf", "col_unit_price"]
}
```

## Complete Data Flow Pipeline

```
┌─────────────────────────────────────────────────────────────────┐
│ 1. JSON Config File (JF_config.json)                           │
├─────────────────────────────────────────────────────────────────┤
│   layout_bundle:                                                │
│     Invoice:                                                    │
│       content:                                                  │
│         static:                                                 │
│           col_static: ["VENDOR#:", "Des: LEATHER", ...]         │
│                                                                 │
│   data_bundle:                                                  │
│     Invoice:                                                    │
│       data_flow:                                                │
│         mappings:                                               │
│           po: {column: "col_po", source_key: 0}                 │
│           item: {column: "col_item", source_key: 1}             │
│           ...                                                   │
└─────────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────────┐
│ 2. BundledConfigLoader.get_layout_config() / get_data_config() │
├─────────────────────────────────────────────────────────────────┤
│   Returns raw config sections as dicts                          │
└─────────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────────┐
│ 3. BuilderConfigResolver.get_layout_bundle()                    │
├─────────────────────────────────────────────────────────────────┤
│   Extracts static_content from layout_config.content.static:    │
│   {                                                             │
│     'static_content': {                                         │
│       'col_static': ["VENDOR#:", "Des: LEATHER", ...]           │
│     }                                                           │
│   }                                                             │
└─────────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────────┐
│ 4. BuilderConfigResolver.get_table_data_resolver()              │
├─────────────────────────────────────────────────────────────────┤
│   Creates TableDataAdapter with:                                │
│   - data_config (mapping rules, data source)                    │
│   - layout_config (static_content)                              │
└─────────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────────┐
│ 5. TableDataAdapter.__init__()                                  │
├─────────────────────────────────────────────────────────────────┤
│   Stores:                                                       │
│   - self.mapping_rules (from data_config)                       │
│   - self.static_content (from layout_config)                    │
│   - self.data_source (raw invoice data)                         │
└─────────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────────┐
│ 6. TableDataAdapter.resolve()                                   │
├─────────────────────────────────────────────────────────────────┤
│   Calls parse_mapping_rules() to convert config to rules        │
│   Then calls prepare_data_rows() to transform data              │
└─────────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────────┐
│ 7. parse_mapping_rules() in data_preparer.py                    │
├─────────────────────────────────────────────────────────────────┤
│   INPUT: mapping_rules dict                                     │
│   {                                                             │
│     "po": {"column": "col_po", "source_key": 0},                │
│     "item": {"column": "col_item", "source_key": 1}             │
│   }                                                             │
│                                                                 │
│   OUTPUT: parsed rules dict                                     │
│   {                                                             │
│     'dynamic_mapping_rules': {...},                             │
│     'static_value_map': {...},                                  │
│     'formula_rules': {...}                                      │
│   }                                                             │
└─────────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────────┐
│ 8. prepare_data_rows() in data_preparer.py                      │
├─────────────────────────────────────────────────────────────────┤
│   INPUT: Raw invoice data (aggregation dict)                    │
│   {                                                             │
│     ('PO123', 'ITEM-001', 10.5, 'LEATHER'): {                   │
│       'sqft_sum': 100.5,                                        │
│       'amount_sum': 1050.0                                      │
│     },                                                          │
│     ...                                                         │
│   }                                                             │
│                                                                 │
│   PROCESS:                                                      │
│   - Loops through each data entry                              │
│   - For each entry, creates row_dict                            │
│   - Maps source_key indices to column_id_map indices            │
│   - Converts: ('PO123', 'ITEM-001', ...) → {1: 'PO123', 2: ...}│
│                                                                 │
│   OUTPUT: data_rows (list of row dicts)                         │
│   [                                                             │
│     {1: "PO123", 2: "ITEM-001", 3: "LEATHER", 4: 100.5, ...},   │
│     {1: "PO123", 2: "ITEM-002", 3: "SUEDE", 4: 200.0, ...},     │
│   ]                                                             │
│   ^^^^^^^ THIS IS THE FORMAT DATABUILDER USES ^^^^^^^           │
└─────────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────────┐
│ 9. TableDataAdapter.resolve() returns                           │
├─────────────────────────────────────────────────────────────────┤
│   {                                                             │
│     'data_rows': [row_dicts...],  ← From prepare_data_rows()   │
│     'pallet_counts': [...],                                     │
│     'static_content': {...}       ← Passed through from __init__│
│   }                                                             │
└─────────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────────┐
│ 10. DataTableBuilder.__init__(resolved_data)                    │
├─────────────────────────────────────────────────────────────────┤
│   Extracts:                                                     │
│   - self.data_rows = resolved_data['data_rows']                 │
│     (List of dicts: {col_idx: value, ...})                      │
│   - self.static_col_values = resolved_data['static_content']... │
│     (List: ["VENDOR#:", "Des: LEATHER", ...])                   │
└─────────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────────┐
│ 11. DataTableBuilder.build()                                    │
├─────────────────────────────────────────────────────────────────┤
│   ROW-BY-ROW WRITING:                                           │
│   for i in range(num_rows):                                     │
│       row_data = self.data_rows[i]  # Get dict for this row     │
│                                                                 │
│       # Write static content                                    │
│       if self.static_col_idx and self.static_col_values:        │
│           static_value = self.static_col_values[i % len(...)]   │
│           worksheet.cell(row=i, col=static_col_idx).value = ... │
│                                                                 │
│       # Write data columns                                      │
│       for col_idx, value in row_data.items():                   │
│           worksheet.cell(row=i, column=col_idx).value = value   │
│                                                                 │
│   OUTPUT: Excel cells filled with data!                         │
└─────────────────────────────────────────────────────────────────┘
```

## Key Conversion Point

**THE KEY CONVERTER IS `prepare_data_rows()` in `data_preparer.py`**

It transforms:
```python
# FROM: Invoice data (tuple keys with aggregated values)
{
    ('PO123', 'ITEM-001', 10.5, 'LEATHER'): {'sqft_sum': 100.5, 'amount_sum': 1050.0}
}

# TO: Row-based structure (column index keys)
[
    {1: 'PO123', 2: 'ITEM-001', 3: 'LEATHER', 4: 100.5, 5: 1050.0}
]
```

Using the mapping rules:
```python
mapping_rules = {
    "po": {"column": "col_po", "source_key": 0},      # tuple[0] → col_po
    "item": {"column": "col_item", "source_key": 1},  # tuple[1] → col_item
    # ...
}
```

And column_id_map:
```python
column_id_map = {
    'col_po': 1,      # col_po is Excel column 1 (A)
    'col_item': 2,    # col_item is Excel column 2 (B)
    'col_desc': 3,    # col_desc is Excel column 3 (C)
    # ...
}
```

## Summary

**DataTableBuilder accepts**:
- `data_rows`: List[Dict[int, Any]] - Row-oriented, column indices as keys
- `static_col_values`: List[str] - Static content to repeat
- `static_col_idx`: int - Which column gets the static content

**Conversion happens in**: `data_preparer.prepare_data_rows()`
- Takes raw data + mapping rules
- Returns row-oriented list of dicts with column indices as keys
