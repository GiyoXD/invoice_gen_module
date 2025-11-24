# Copilot Instructions: Invoice Generator Module

## Project Overview

This is an **Excel invoice generation system** that uses the **Builder pattern** to construct complex, multi-sheet invoices from JSON configuration files and template workbooks. The codebase is undergoing an active refactoring from monolithic scripts to a clean, bundle-based architecture.

**Core Purpose**: Generate Excel invoices by combining template structures with dynamic data, applying complex styling, and handling multi-table layouts (e.g., Invoice sheets, Packing lists).

## Architecture: Bundle Cascade Pattern

### The Big Picture

The system uses a **3-tier resolver pattern** to eliminate parameter explosion:

```
BundledConfigLoader → BuilderConfigResolver → TableDataResolver → Builders
```

**Critical Concept**: Builders accept **4 configuration bundles** instead of 20+ individual parameters:

1. **`style_config`** - All styling (fonts, borders, fills)
2. **`context_config`** - Runtime data (invoice_data, sheet_name, args)
3. **`layout_config`** - Layout controls (skip flags, merge rules, blanks)
4. **`data_config`** - Data sources and mappings

### Bundle Storage Pattern

**✅ DO THIS** (Modern approach - use properties):
```python
class Builder:
    def __init__(self, worksheet, style_config, context_config, layout_config, data_config):
        # Store bundles ONLY - no extraction!
        self.worksheet = worksheet
        self.style_config = style_config
        self.context_config = context_config
        self.layout_config = layout_config
        self.data_config = data_config
    
    @property
    def sheet_name(self):
        """Access bundle values via properties"""
        return self.context_config.get('sheet_name', '')
```

**❌ DON'T DO THIS** (Legacy approach - avoid):
```python
def __init__(self, worksheet, sheet_name, config, data, styling, ...):  # 20+ params
    self.worksheet = worksheet
    self.sheet_name = sheet_name
    # ... 20+ more assignments
```

## Key Components

### Config Layer (`invoice_generator/config/`)

- **`BundledConfigLoader`**: Loads JSON config files (v2.1 format with `_meta`, `processing`, `styling_bundle`, `layout_bundle`, `data_bundle` sections)
- **`BuilderConfigResolver`**: Bridges loader → builders. Use `get_layout_bundles_with_data()` for LayoutBuilder
- **`TableDataResolver`**: Prepares table data. Always use `get_table_data_resolver()` from BuilderConfigResolver

### Builder Layer (`invoice_generator/builders/`)

**Director**:
- **`LayoutBuilder`**: Orchestrates header, table, footer construction. Cascades bundles to child builders

**Specialized Builders**:
- **`DataTableBuilder`**: Writes table rows (24 params → 4 bundles after refactor)
- **`FooterBuilder`**: Creates footer/summary sections (15 params → 3 bundles + footer_row_num)
- **`HeaderBuilder`**: Writes headers (not yet refactored to bundles)
- **`TemplateStateBuilder`**: Captures/restores template structure (no bundles - uses 4 direct params)
- **`WorkbookBuilder`**: Creates empty workbooks with named sheets

### Processor Layer (`invoice_generator/processors/`)

- **`SheetProcessor`** (ABC): Base class
- **`SingleTableProcessor`**: Standard invoices (one table per sheet)
- **`MultiTableProcessor`**: Packing lists (multiple table chunks per sheet)

## Essential Patterns

### 1. Using BuilderConfigResolver (ALWAYS use this)

```python
from invoice_generator.config.builder_config_resolver import BuilderConfigResolver

resolver = BuilderConfigResolver(
    config_loader=config_loader,
    sheet_name='Invoice',
    worksheet=worksheet,
    args=args,
    invoice_data=invoice_data,
    pallets=31
)

# For LayoutBuilder
style_config, context_config, layout_config = resolver.get_layout_bundles_with_data()

# For multi-table scenarios with table_key
style_config, context_config, layout_config = resolver.get_layout_bundles_with_data(table_key='1')

# For data preparation
table_resolver = resolver.get_table_data_resolver()
resolved_data = table_resolver.resolve()
```

### 2. Multi-Table Processing Pattern

```python
# Table '1'
resolver_t1 = BuilderConfigResolver(config_loader, 'Packing list', worksheet, args, invoice_data, pallets=20)
style_cfg, ctx_cfg, layout_cfg = resolver_t1.get_layout_bundles_with_data(table_key='1')

builder_t1 = LayoutBuilder(workbook, worksheet, template, style_cfg, ctx_cfg, layout_cfg)
builder_t1.build()
current_row = builder_t1.next_row_after_footer

# Table '2' - position after table 1
resolver_t2 = BuilderConfigResolver(config_loader, 'Packing list', worksheet, args, invoice_data, pallets=11)
style_cfg, ctx_cfg, layout_cfg = resolver_t2.get_layout_bundles_with_data(table_key='2')
layout_cfg['sheet_config']['header_row'] = current_row + 1  # Position second table

builder_t2 = LayoutBuilder(
    workbook, worksheet, template, style_cfg, ctx_cfg, layout_cfg,
    skip_template_header_restoration=True,  # Only restore once
    skip_template_footer_restoration=False  # Restore after last table
)
builder_t2.build()
```

### 3. Template Separation Architecture

**Critical**: Templates are READ-ONLY. Output workbooks are created fresh:

```python
template_workbook = load_workbook(template_path)  # Read-only source
output_workbook = load_workbook(template_path)    # Copy for writing

template_ws = template_workbook['Invoice']  # For state capture
output_ws = output_workbook['Invoice']      # For actual writing

# Capture template structure
state_builder = TemplateStateBuilder(
    template_worksheet=template_ws,
    header_end_row=header_row,
    footer_start_row=footer_row,
    max_row=template_ws.max_row
)

# Restore to output
state_builder.restore_header_only(output_ws)
state_builder.restore_footer_only(output_ws, footer_start_row=write_position)
```

## Project-Specific Conventions

### Logging

**CRITICAL: NEVER use `print()` statements** - Always use the `logging` library instead (per `.instructions.md`):


```python
import logging
logger = logging.getLogger(__name__)

logger.debug(f"Variable state: {value}")           # Diagnostic info
logger.info(f"Completed step X")                  # Normal progress
logger.warning(f"Using fallback for {param}")     # Unexpected but recoverable
logger.error(f"Failed to build {component}")      # Operation failed
logger.critical(f"Missing required config")       # Fatal error
```

### Configuration File Structure

**Only bundled config v2.1 format is supported** (found in `config_bundled/`). Legacy formats are deprecated and will be removed.

```json
{
  "_meta": {"config_version": "2.1", "customer": "CLW"},
  "processing": {
    "sheets": ["Invoice", "Packing list"],
    "data_sources": {
      "Invoice": "aggregation",
      "Packing list": "processed_tables_multi"
    }
  },
  "styling_bundle": {"Invoice": {"styling_config": {...}}},
  "layout_bundle": {"Invoice": {"header_row": 21, "add_blank_before_footer": true, ...}},
  "data_bundle": {"Invoice": {"mappings": {...}, "header_info": {...}}}
}
```

**Supported Data Source Types** (in `processing.data_sources`):
1. **`aggregation`**: Standard invoice data (single aggregated table)
2. **`DAF_aggregation`**: DAF (Delivered At Frontier) shipping term - uses DAF-specific aggregation data
3. **`processed_tables_multi`**: Multi-table layouts (e.g., packing lists with multiple table chunks)

### Skip Flags Pattern

Control builder execution via `layout_config` skip flags:
- `skip_header_builder`: Skip header construction
- `skip_data_table_builder`: Skip table data writing
- `skip_footer_builder`: Skip footer construction
- `skip_template_header_restoration`: Skip restoring template header
- `skip_template_footer_restoration`: Skip restoring template footer

**Multi-table usage**: Set `skip_template_header_restoration=True` for all tables except first, `skip_template_footer_restoration=False` only for last table.

## Development Workflow

### Running Tests

```powershell
# Main test suite
python -m pytest tests/

# Specific test
python tests/test_invoice_generation.py
```

### Generating Invoices

```powershell
python -m invoice_generator.generate_invoice data/CLW.json --output result.xlsx
```

### File Naming Convention

- Templates: `{customer}.xlsx` (e.g., `CLW.xlsx`)
- Configs: `{customer}_config.json` OR `config_bundled/{customer}_config/{customer}_config.json`
- Data: `{customer}.json` or `{customer}_data.json`

## Common Pitfalls

### ❌ DON'T: Extract bundle values in constructor
```python
def __init__(self, style_config, context_config):
    self.sheet_name = context_config.get('sheet_name')  # NO!
    self.invoice_data = context_config.get('invoice_data')  # NO!
```

### ✅ DO: Use properties for frequent access
```python
def __init__(self, style_config, context_config):
    self.context_config = context_config  # Store bundle

@property
def sheet_name(self):
    return self.context_config.get('sheet_name', '')  # Access via property
```

### ❌ DON'T: Call data preparation directly from builders
```python
from invoice_generator.data.data_preparer import prepare_data_rows
rows = prepare_data_rows(...)  # Tight coupling!
```

### ✅ DO: Use TableDataResolver
```python
resolver = builder_resolver.get_table_data_resolver()
result = resolver.resolve()  # Returns prepared data + metadata
```

### ❌ DON'T: Modify template workbook
```python
template_ws.cell(row=5, column=1, value="Modified")  # Corrupts template!
```

### ✅ DO: Write to output workbook
```python
output_ws.cell(row=5, column=1, value="Modified")  # Correct
```

## Important Files to Reference

- **`docs/BUNDLE_CASCADE_PATTERN.md`**: Complete bundle architecture guide
- **`docs/examples/layout_builder_resolver_pattern.py`**: Usage examples
- **`docs/PURE_BUNDLE_REFACTOR_SUMMARY.md`**: Refactoring history and metrics
- **`docs/builders/LOGGING_GUIDE.md`**: Logging level guidelines
- **`docs/CONFIG_AUDIT.md`**: Configuration schema reference

## Active Refactoring Goals

1. **Eliminate legacy parameter passing**: All builders should use bundles
2. **Centralize data resolution**: Use TableDataResolver, not direct imports
3. **Complete v2.1 config migration**: All configs must use bundled format (legacy formats removed)
4. **Document patterns**: Update builder docs when adding bundle support

When adding features, **always use the resolver pattern** and **avoid adding new individual parameters** to builder constructors. All new configurations must follow the v2.1 bundled format structure.
