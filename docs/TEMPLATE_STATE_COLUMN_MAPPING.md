# Template State Builder: Column Mapping Feature

## Overview

The **Column Mapping** feature allows `TemplateStateBuilder` to intelligently shift template content when columns are filtered or removed during invoice generation.

## The Problem

When generating invoices with conditional column logic (e.g., DAF mode, custom mode), some template columns may be removed from the output:

```
Template:  [A] [B] [C] [D] [E] [F] [G]  (7 columns)
                    ↓ Filter (remove C and F)
Output:    [A] [B] [D] [E] [G]          (5 columns)
```

**Without column mapping**, template content restoration fails:
- Template content from column D (index 4) goes to output column D (index 4) ❌
- But column D in output is actually template column E! ❌
- **Result**: Template content appears in wrong columns or gets lost

**With column mapping**, content is shifted correctly:
- Template column D (index 4) → Output column C (index 3) ✅
- Template column E (index 5) → Output column D (index 4) ✅
- Template column G (index 7) → Output column E (index 5) ✅

## Usage

### Step 1: Create TemplateStateBuilder

```python
from invoice_generator.builders.template_state_builder import TemplateStateBuilder

template_state = TemplateStateBuilder(
    worksheet=template_ws,
    num_header_cols=7,
    header_end_row=21,
    footer_start_row=50,
    debug=True
)
```

### Step 2: Define Column Mapping

Create a mapping dictionary that specifies where each template column should go in the output:

```python
column_mapping = {
    1: 1,      # Template col 1 → Output col 1 (kept)
    2: 2,      # Template col 2 → Output col 2 (kept)
    3: None,   # Template col 3 → Removed (content skipped)
    4: 3,      # Template col 4 → Output col 3 (shifted left)
    5: 4,      # Template col 5 → Output col 4 (shifted left)
    6: None,   # Template col 6 → Removed (content skipped)
    7: 5       # Template col 7 → Output col 5 (shifted left)
}
```

**Key**: Template column index (1-based)  
**Value**: Output column index (1-based), or `None` if removed

### Step 3: Apply Mapping

```python
template_state.set_column_mapping(column_mapping)
```

### Step 4: Restore Template Content

When restoring, the mapping is automatically applied:

```python
# Restore header with column shifts
template_state.restore_header_only(
    target_worksheet=output_ws,
    actual_num_cols=5  # Output has 5 columns (not 7)
)

# Restore footer with column shifts
template_state.restore_footer_only(
    target_worksheet=output_ws,
    footer_start_row=60,
    actual_num_cols=5
)
```

## Automatic Mapping Generation

For convenience, you can auto-generate mappings from filtered column lists:

```python
def generate_column_mapping(template_columns, filtered_columns):
    """
    Generate mapping from column ID lists.
    
    Args:
        template_columns: ['col_static', 'col_po', 'col_item', ...]
        filtered_columns: ['col_static', 'col_po', 'col_desc', ...]
    
    Returns:
        {1: 1, 2: 2, 3: None, 4: 3, ...}
    """
    mapping = {}
    output_index = 1
    
    for template_index, col_id in enumerate(template_columns, start=1):
        if col_id in filtered_columns:
            mapping[template_index] = output_index
            output_index += 1
        else:
            mapping[template_index] = None
    
    return mapping

# Usage
template_cols = ['col_static', 'col_po', 'col_item', 'col_desc', 
                 'col_qty_sf', 'col_unit_price', 'col_amount']
filtered_cols = ['col_static', 'col_po', 'col_desc', 
                 'col_qty_sf', 'col_amount']

mapping = generate_column_mapping(template_cols, filtered_cols)
template_state.set_column_mapping(mapping)
```

## What Gets Shifted

When column mapping is applied, the following template elements are shifted:

1. **Cell Values**: Content is written to the mapped column
2. **Cell Formatting**: Fonts, fills, borders, alignment follow the content
3. **Merged Cells**: Merge ranges are adjusted to mapped column positions

### Example: Merged Cell Shifting

```
Template merge: B5:D5 (columns 2-4)
Mapping: {2: 2, 3: None, 4: 3}
Result: B5:C5 (columns 2-3, shifted left)
```

### Skipped Content

When a column is mapped to `None`:
- Cell values are NOT written (skipped)
- Cell formatting is NOT applied
- Merged cells spanning only removed columns are skipped
- Merged cells partially in removed columns are skipped (to avoid errors)

## Debug Logging

Enable debug logging to see column shifts in action:

```python
template_state = TemplateStateBuilder(..., debug=True)
template_state.set_column_mapping(mapping)
```

Example debug output:
```
Column mapping set: 5 columns mapped
  Active: {1: 1, 2: 2, 4: 3, 5: 4, 7: 5}
  Skipped template columns: [3, 6]
Shifted column 4 -> 3 at row 5 (value: 'Description')
Merged (shifted): D5:F5 -> C5:D5
```

## Integration with Layout Builder

The column mapping should be set **after** `TemplateStateBuilder` is created but **before** restoration methods are called:

```python
# In LayoutBuilder.build() or similar
self.template_state_builder = TemplateStateBuilder(...)

# Calculate column mapping from filtered columns
mapping = generate_column_mapping(
    template_columns=all_columns_from_config,
    filtered_columns=active_columns_after_filter
)

# Apply mapping
self.template_state_builder.set_column_mapping(mapping)

# Restore (mapping will be applied automatically)
self.template_state_builder.restore_header_only(...)
```

## Best Practices

### ✅ DO:
- Create mapping AFTER `TemplateStateBuilder` initialization
- Set mapping BEFORE calling restoration methods
- Use auto-generation helper for consistency
- Enable debug logging when troubleshooting column shifts
- Map all template columns (even if kept at same position)

### ❌ DON'T:
- Don't set mapping after restoration (too late!)
- Don't mix 0-based and 1-based indexing (always use 1-based)
- Don't forget to update `actual_num_cols` parameter to match output
- Don't skip setting mapping if columns were filtered (content will be wrong)

## Common Use Cases

### 1. DAF Mode (Shipping Term)
```python
# DAF mode removes unit price column
template_cols = ['static', 'po', 'item', 'desc', 'qty', 'price', 'amount']
daf_cols = ['static', 'po', 'item', 'desc', 'qty', 'amount']  # No price

mapping = generate_column_mapping(template_cols, daf_cols)
# Result: {1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: None, 7: 6}
```

### 2. Custom Mode
```python
# Custom mode shows different columns
template_cols = ['static', 'po', 'item', 'desc', 'qty', 'price', 'amount']
custom_cols = ['static', 'item', 'qty', 'amount']  # Minimal columns

mapping = generate_column_mapping(template_cols, custom_cols)
# Result: {1: 1, 2: None, 3: 2, 4: None, 5: 3, 6: None, 7: 4}
```

### 3. Multi-Table with Different Columns
```python
# Table 1: Full columns
table1_mapping = {1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 7}

# Table 2: Reduced columns
table2_mapping = {1: 1, 2: 2, 3: None, 4: 3, 5: 4, 6: None, 7: 5}

# Use different state builders for each table
state1.set_column_mapping(table1_mapping)
state2.set_column_mapping(table2_mapping)
```

## See Also

- **Example**: `docs/examples/template_state_column_mapping_example.py`
- **Builder Docs**: `docs/builders/template_state_builder_documentation.md`
- **Layout Builder**: `docs/builders/layout_builder_documentation.md`
