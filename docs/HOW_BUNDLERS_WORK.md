# How Bundlers Work Together - Complete Flow

## Overview

The bundle system uses a **3-tier cascade architecture** to eliminate parameter explosion and provide clean separation of concerns:

```
JSON Config → BundledConfigLoader → BuilderConfigResolver → Builders
```

---

## Complete Data Flow Diagram

```
┌─────────────────────────────────────────────────────────────────────┐
│                    1. CONFIGURATION LOADING                         │
│                                                                     │
│   JF_config.json (v2.1 format)                                     │
│   ├── _meta: {config_version, customer}                            │
│   ├── processing: {sheets, data_sources}                           │
│   ├── styling_bundle: {Invoice: {...}, Packing list: {...}}        │
│   ├── layout_bundle: {Invoice: {...}, Packing list: {...}}         │
│   └── data_bundle: {Invoice: {...}, Packing list: {...}}           │
│                                                                     │
│                          ↓ (load)                                   │
│                                                                     │
│   BundledConfigLoader (config/loader.py)                           │
│   • Loads JSON into memory                                         │
│   • Provides access methods:                                       │
│     - get_sheet_layout(sheet_name)                                 │
│     - get_sheet_styling(sheet_name)                                │
│     - get_data_source(sheet_name)                                  │
│   • No transformation, just clean access                           │
└─────────────────────────────────────────────────────────────────────┘
                                 │
                                 ↓
┌─────────────────────────────────────────────────────────────────────┐
│                    2. INVOICE GENERATION START                      │
│                                                                     │
│   generate_invoice.py                                              │
│   • Loads JF.json (invoice data)                                   │
│   • Creates BundledConfigLoader(config_data)                       │
│   • Creates WorkbookBuilder → output workbook                      │
│   • Loads template workbook (for state capture)                    │
│   • Determines processor type:                                     │
│     - 'aggregation' → SingleTableProcessor                         │
│     - 'processed_tables_multi' → MultiTableProcessor               │
└─────────────────────────────────────────────────────────────────────┘
                                 │
                                 ↓
┌─────────────────────────────────────────────────────────────────────┐
│                    3. PROCESSOR INITIALIZATION                      │
│                                                                     │
│   For "Packing list" sheet (processed_tables_multi):               │
│                                                                     │
│   MultiTableProcessor(                                             │
│       config_loader=config_loader,                                 │
│       sheet_name="Packing list",                                   │
│       template_worksheet=template["Packing list"],                 │
│       output_worksheet=output["Packing list"],                     │
│       output_workbook=output_workbook,                             │
│       invoice_data=invoice_data,                                   │
│       args=args,                                                   │
│       final_grand_total_pallets=31                                 │
│   )                                                                │
│                                                                     │
│   Processor extracts table_keys from config:                       │
│   • table_keys = ['1', '2']                                        │
│   • all_tables_data = invoice_data['processed_tables_data']       │
└─────────────────────────────────────────────────────────────────────┘
                                 │
                                 ↓
┌─────────────────────────────────────────────────────────────────────┐
│            4. TEMPLATE STATE CAPTURE (Once, Reused)                │
│                                                                     │
│   MultiTableProcessor.process():                                   │
│                                                                     │
│   # OPTIMIZATION: Capture template ONCE before loop                │
│   template_state_builder = TemplateStateBuilder(                   │
│       template_worksheet=template_worksheet,                       │
│       header_end_row=22,  # From config                            │
│       footer_start_row=21, # From config                           │
│       max_row=template_worksheet.max_row                           │
│   )                                                                │
│   • Captures header styling (rows 1-22)                            │
│   • Captures footer styling (rows 21+)                             │
│   • Stores for reuse by all table builders                         │
└─────────────────────────────────────────────────────────────────────┘
                                 │
                                 ↓
┌─────────────────────────────────────────────────────────────────────┐
│                    5. TABLE LOOP - Table '1'                        │
│                                                                     │
│   BuilderConfigResolver created for table '1':                     │
│                                                                     │
│   resolver = BuilderConfigResolver(                                │
│       config_loader=config_loader,                                 │
│       sheet_name="Packing list",                                   │
│       worksheet=output_worksheet,                                  │
│       args=args,                                                   │
│       invoice_data=table_invoice_data,  # Just table '1' data      │
│       pallets=20  # Table '1' has 20 pallets                       │
│   )                                                                │
│                                                                     │
│   ┌─────────────────────────────────────────────────────────────┐  │
│   │ RESOLVER CREATES 3 BUNDLES:                                 │  │
│   │                                                             │  │
│   │ style_config = {                                            │  │
│   │     'styling_config': StylingConfigModel(...)               │  │
│   │ }                                                           │  │
│   │                                                             │  │
│   │ context_config = {                                          │  │
│   │     'sheet_name': 'Packing list',                           │  │
│   │     'args': args,                                           │  │
│   │     'pallets': 20,                                          │  │
│   │     'invoice_data': table_invoice_data,                     │  │
│   │     'all_sheet_configs': {...}                              │  │
│   │ }                                                           │  │
│   │                                                             │  │
│   │ layout_config = {                                           │  │
│   │     'sheet_config': {                                       │  │
│   │         'structure': {'header_row': 21},  # Start at 21    │  │
│   │         'blanks': {...},                                    │  │
│   │         'merge_rules': {...}                                │  │
│   │     },                                                      │  │
│   │     'skip_template_header_restoration': False, # First table│  │
│   │     'skip_template_footer_restoration': True,  # Skip until │  │
│   │     'resolved_data': {...}  # From TableDataResolver        │  │
│   │ }                                                           │  │
│   └─────────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────────┘
                                 │
                                 ↓
┌─────────────────────────────────────────────────────────────────────┐
│                    6. LAYOUT BUILDER (DIRECTOR)                     │
│                                                                     │
│   LayoutBuilder(                                                   │
│       output_workbook,                                             │
│       output_worksheet,                                            │
│       template_worksheet,                                          │
│       style_config=style_config,      # BUNDLE 1                   │
│       context_config=context_config,  # BUNDLE 2                   │
│       layout_config=layout_config,    # BUNDLE 3                   │
│       template_state_builder=template_state_builder # Pre-captured │
│   )                                                                │
│                                                                     │
│   LayoutBuilder.build():                                           │
│   ┌─────────────────────────────────────────────────────────────┐  │
│   │ STEP 1: Restore Template Header (if not skipped)           │  │
│   │ • Uses pre-captured template_state_builder                  │  │
│   │ • Restores rows 1-22 to output worksheet                    │  │
│   ├─────────────────────────────────────────────────────────────┤  │
│   │ STEP 2: Build Header Section                               │  │
│   │ • Creates HeaderBuilder(worksheet, style, layout)           │  │
│   │ • Writes styled header rows at row 21-22                    │  │
│   ├─────────────────────────────────────────────────────────────┤  │
│   │ STEP 3: Build Data Table                                   │  │
│   │ • Creates DataTableBuilder(                                 │  │
│   │       worksheet,                                            │  │
│   │       style_config,     ← CASCADE                           │  │
│   │       context_config,   ← CASCADE                           │  │
│   │       layout_config,    ← CASCADE                           │  │
│   │       data_config       ← CREATE                            │  │
│   │   )                                                         │  │
│   │ • Writes 30 data rows (rows 23-52)                          │  │
│   ├─────────────────────────────────────────────────────────────┤  │
│   │ STEP 4: Build Footer Section                               │  │
│   │ • Creates FooterBuilder(                                    │  │
│   │       worksheet,                                            │  │
│   │       style_config,     ← CASCADE                           │  │
│   │       context_config,   ← CASCADE                           │  │
│   │       data_config,      ← CREATE                            │  │
│   │       footer_row_num=53                                     │  │
│   │   )                                                         │  │
│   │ • Writes footer rows 53-55 (Total, CBM, Pallets)           │  │
│   ├─────────────────────────────────────────────────────────────┤  │
│   │ STEP 5: Skip Template Footer (skip=True for mid-tables)    │  │
│   └─────────────────────────────────────────────────────────────┘  │
│                                                                     │
│   Returns: next_row_after_footer = 56                              │
└─────────────────────────────────────────────────────────────────────┘
                                 │
                                 ↓
┌─────────────────────────────────────────────────────────────────────┐
│                    7. TABLE LOOP - Table '2'                        │
│                                                                     │
│   current_row = 56  # From table '1' next_row_after_footer         │
│                                                                     │
│   resolver = BuilderConfigResolver(                                │
│       ...same setup...                                             │
│       invoice_data=table_2_data,  # Just table '2' data            │
│       pallets=11  # Table '2' has 11 pallets                       │
│   )                                                                │
│                                                                     │
│   layout_config['sheet_config']['structure']['header_row'] = 56    │
│   layout_config['skip_template_header_restoration'] = True         │
│   layout_config['skip_template_footer_restoration'] = True         │
│                                                                     │
│   LayoutBuilder(..., template_state_builder=same_instance)         │
│   • REUSES pre-captured template state                             │
│   • Skips header restoration (already done in table 1)             │
│   • Builds header at row 56-57                                     │
│   • Builds data at rows 58-71 (14 rows)                            │
│   • Builds footer at rows 72-73                                    │
│   • Skips template footer restoration                              │
│                                                                     │
│   Returns: next_row_after_footer = 74                              │
└─────────────────────────────────────────────────────────────────────┘
                                 │
                                 ↓
┌─────────────────────────────────────────────────────────────────────┐
│                    8. GRAND TOTAL & TEMPLATE FOOTER                │
│                                                                     │
│   MultiTableProcessor:                                             │
│   • Writes Grand Total row at 74 (31 pallets)                      │
│   • Restores template footer at row 75+                            │
│   • Uses same template_state_builder.restore_footer_only()         │
└─────────────────────────────────────────────────────────────────────┘
                                 │
                                 ↓
┌─────────────────────────────────────────────────────────────────────┐
│                    9. FINAL OUTPUT                                  │
│                                                                     │
│   Output Excel Structure:                                          │
│   ├── Rows 1-20: Template header (logo, company info)              │
│   ├── Rows 21-55: Table 1 (header + 30 data rows + footer)         │
│   ├── Rows 56-73: Table 2 (header + 14 data rows + footer)         │
│   ├── Row 74: Grand Total (31 pallets)                             │
│   └── Rows 75+: Template footer (terms, signatures)                │
└─────────────────────────────────────────────────────────────────────┘
```

---

## Key Design Patterns

### 1. Bundle Cascade Pattern

**Problem:** Passing 20+ parameters to each builder

**Solution:** Group into 4 semantic bundles:

```python
# ❌ OLD WAY (20+ parameters)
DataTableBuilder(
    worksheet, sheet_name, config, styling, data_source, 
    mappings, header_info, args, pallets, start_row, 
    merge_rules, blanks, static_content, ...  # 20+ params
)

# ✅ NEW WAY (4 bundles)
DataTableBuilder(
    worksheet,
    style_config,    # All styling
    context_config,  # Runtime context
    layout_config,   # Layout controls
    data_config      # Data sources
)
```

**Benefits:**
- **Tiny constructors:** 5 lines instead of 50
- **No maintenance:** Add new config → automatically available
- **Consistent interface:** Same bundles across all builders

### 2. Lazy Access via Properties

**Problem:** Extracting all bundle values in constructor is wasteful

**Solution:** Access on-demand via `@property`:

```python
class DataTableBuilder:
    def __init__(self, worksheet, style_config, context_config, layout_config, data_config):
        # Store bundles (NO extraction!)
        self.worksheet = worksheet
        self.style_config = style_config
        self.context_config = context_config
        self.layout_config = layout_config
        self.data_config = data_config
    
    # Access frequently-used values via properties
    @property
    def sheet_name(self):
        return self.context_config.get('sheet_name', '')
    
    @property
    def data_source(self):
        return self.data_config.get('data_source')
    
    # Access rare values directly
    def build(self):
        if self.layout_config.get('rare_flag'):  # Direct access
            # ...
```

**Benefits:**
- Clean access syntax: `self.sheet_name`
- Only extract what's used
- No stale extracted values

### 3. Template State Reuse

**Problem:** Multi-table processing captured template state N times (inefficient)

**Solution:** Capture once, reuse for all tables:

```python
# Before loop: Capture ONCE
template_state_builder = TemplateStateBuilder(
    template_worksheet=template_ws,
    header_end_row=22,
    footer_start_row=21,
    max_row=template_ws.max_row
)

# In loop: REUSE for each table
for table_key in ['1', '2']:
    layout_builder = LayoutBuilder(
        ...,
        template_state_builder=template_state_builder  # Pre-captured!
    )
```

**Benefits:**
- 50% faster multi-table processing
- Consistent template capture
- Cleaner code

### 4. Config Key Separation

**Critical:** Different config keys serve different purposes:

```python
# ❌ WRONG: header_row refers to TEMPLATE position (row 1)
layout_config['header_row'] = 56  # Wrong key!

# ✅ CORRECT: structure.header_row controls TABLE write position
layout_config['sheet_config']['structure']['header_row'] = 56  # Correct!
```

**Key Mapping:**
- `header_row`: Template header end position (usually row 1-22)
- `structure.header_row`: Where to write table header (row 21, 56, etc.)
- `footer_start_row`: Template footer start position

---

## Bundle Contents Reference

### `style_config` Bundle
```python
{
    'styling_config': StylingConfigModel(
        row_height=25,
        font_name='Arial',
        colors={'header': 'CCCCCC', 'data': 'FFFFFF'},
        borders={'style': 'thin'},
        ...
    )
}
```
**Used by:** All builders for consistent styling

---

### `context_config` Bundle
```python
{
    'sheet_name': 'Packing list',
    'args': argparse.Namespace(...),
    'pallets': 20,  # Table-specific pallet count
    'invoice_data': {...},  # Table-specific data
    'all_sheet_configs': {...},  # Cross-sheet references
    'final_grand_total_pallets': 31,  # Grand total
    'header_info': {...},  # Passed to child builders
    'pallet_count': 20,  # Passed to child builders
    'is_last_table': False,  # Multi-table context
    'dynamic_desc_used': False  # Description tracking
}
```
**Used by:** LayoutBuilder, DataTableBuilder, FooterBuilder

---

### `layout_config` Bundle
```python
{
    'sheet_config': {
        'structure': {
            'header_row': 21,  # CRITICAL: Where to write table
            'columns': {...}
        },
        'blanks': {
            'add_blank_after_header': True,
            'add_blank_before_footer': False
        },
        'merge_rules': {...},
        'skip_data_table_builder': False,
        'skip_header_builder': False,
        'skip_footer_builder': False,
        'skip_template_header_restoration': False,
        'skip_template_footer_restoration': True
    },
    'resolved_data': {...},  # From TableDataResolver
    'enable_text_replacement': False
}
```
**Used by:** LayoutBuilder, DataTableBuilder

---

### `data_config` Bundle
```python
{
    'data_source': {  # Actual data from JSON
        'po': '12345',
        'item': 'Widget',
        ...
    },
    'data_source_type': 'processed_tables_multi',
    'header_info': {  # Column metadata
        'po': {'col': 1, 'width': 15},
        'item': {'col': 2, 'width': 30}
    },
    'mapping_rules': {  # JSON key → Excel column
        'po': 'po',
        'item': 'item_description'
    },
    'table_key': '1',  # Multi-table identifier
    'footer_config': {...},
    'sum_ranges': [(23, 52)]  # For footer calculations
}
```
**Used by:** DataTableBuilder, FooterBuilder

---

## Resolver Flow Example

### Step-by-Step: Processing Table '2'

```python
# 1. Create resolver with table-specific context
resolver = BuilderConfigResolver(
    config_loader=config_loader,
    sheet_name='Packing list',
    worksheet=output_worksheet,
    args=args,
    invoice_data={'processed_tables_data': {'2': table_2_data}},
    pallets=11  # Table 2 has 11 pallets
)

# 2. Get style bundle (same for all tables)
style_config = resolver.get_style_bundle()
# Returns: {'styling_config': StylingConfigModel(...)}

# 3. Get context bundle (table-specific)
context_config = resolver.get_context_bundle(
    invoice_data=table_2_invoice_data,
    enable_text_replacement=False
)
# Returns: {
#     'sheet_name': 'Packing list',
#     'args': args,
#     'pallets': 11,  # Table 2 specific
#     'invoice_data': table_2_invoice_data,
#     ...
# }

# 4. Get layout bundle (structure config)
layout_config = resolver.get_layout_bundle()
# Returns: {
#     'sheet_config': {
#         'structure': {'header_row': 21},  # Will override
#         'blanks': {...},
#         'merge_rules': {...}
#     }
# }

# 5. Override table position for Table 2
layout_config['sheet_config']['structure']['header_row'] = 56  # After Table 1

# 6. Set skip flags for multi-table
layout_config['skip_template_header_restoration'] = True  # Already done
layout_config['skip_template_footer_restoration'] = True  # Wait for end

# 7. Get resolved data via TableDataResolver
table_data_resolver = resolver.get_table_data_resolver(table_key='2')
resolved_data = table_data_resolver.resolve()
# Returns: {
#     'data_rows': [14 processed rows],
#     'sum_ranges': [(58, 71)],
#     'header_info': {...},
#     'dynamic_desc_used': False
# }
layout_config['resolved_data'] = resolved_data

# 8. Create LayoutBuilder with bundles
layout_builder = LayoutBuilder(
    output_workbook,
    output_worksheet,
    template_worksheet,
    style_config=style_config,      # Bundle 1
    context_config=context_config,  # Bundle 2
    layout_config=layout_config,    # Bundle 3
    template_state_builder=pre_captured_state  # Optimization!
)

# 9. Build table
layout_builder.build()
# Result:
# - Header at rows 56-57
# - Data at rows 58-71 (14 rows)
# - Footer at rows 72-73
# - next_row_after_footer = 74
```

---

## Common Issues & Solutions

### Issue 1: Tables Overlapping
**Symptom:** Table 2 overwrites Table 1

**Root Cause:**
```python
# ❌ WRONG KEY
layout_config['header_row'] = current_row  # Sets template position!
```

**Fix:**
```python
# ✅ CORRECT KEY
layout_config['sheet_config']['structure']['header_row'] = current_row
```

---

### Issue 2: Template Footer Too Early
**Symptom:** Footer appears after first table instead of at end

**Root Cause:**
```python
# ❌ WRONG: Not skipping footer restoration
layout_config['skip_template_footer_restoration'] = False  # Restores too early!
```

**Fix:**
```python
# ✅ CORRECT: Skip until last table
for i, table_key in enumerate(table_keys):
    is_last_table = (i == len(table_keys) - 1)
    layout_config['skip_template_footer_restoration'] = not is_last_table
```

---

### Issue 3: Multiple Template Captures
**Symptom:** Slow multi-table processing, inconsistent template capture

**Root Cause:**
```python
# ❌ WRONG: Capturing in each iteration
for table_key in table_keys:
    template_state_builder = TemplateStateBuilder(...)  # Redundant!
```

**Fix:**
```python
# ✅ CORRECT: Capture once before loop
template_state_builder = TemplateStateBuilder(...)  # Once!

for table_key in table_keys:
    layout_builder = LayoutBuilder(
        ...,
        template_state_builder=template_state_builder  # Reuse!
    )
```

---

## Testing the Bundle Flow

### Verify Bundle Contents
```python
# In resolver
logger.debug(f"Style config keys: {list(style_config.keys())}")
logger.debug(f"Context config: {context_config}")
logger.debug(f"Layout structure: {layout_config.get('sheet_config', {}).get('structure')}")

# In builder
logger.debug(f"[{self.__class__.__name__}] sheet_name={self.sheet_name}")
logger.debug(f"[{self.__class__.__name__}] table_header_row={self.table_header_row}")
```

### Verify Table Positioning
```bash
python -m invoice_generator.generate_invoice JF.json --debug 2>&1 | 
    Select-String "table_header_row|DataTableBuilder completed"
```

Expected output:
```
table_header_row=21  # Table 1
DataTableBuilder completed: 30 data rows written (rows 23-52)
table_header_row=56  # Table 2
DataTableBuilder completed: 14 data rows written (rows 58-71)
```

---

## Summary

The bundle system provides:

1. **Clean Architecture:** 3-tier loader → resolver → builder
2. **Zero Maintenance:** New configs automatically available
3. **Consistent Interface:** 4 bundles across all builders
4. **Performance:** Template state reuse, lazy property access
5. **Type Safety:** Bundles group related configs logically

**Critical Files:**
- `config/loader.py` - Config loading
- `config/builder_config_resolver.py` - Bundle preparation
- `builders/layout_builder.py` - Director pattern
- `builders/data_table_builder.py` - Pure bundle implementation
- `builders/footer_builder.py` - Pure bundle implementation
- `processors/multi_table_processor.py` - Multi-table orchestration
