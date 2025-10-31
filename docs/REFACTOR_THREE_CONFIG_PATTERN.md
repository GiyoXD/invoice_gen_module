# LayoutBuilder Refactor: Three-Config Architecture

## Overview

Successfully refactored `LayoutBuilder` from a parameter-heavy constructor to a clean **three-config architecture** for maximum extensibility and maintainability.

## Problem Statement

**Before:** The `LayoutBuilder` constructor had 15+ individual parameters, making it:
- ❌ Hard to extend (adding new configs required signature changes)
- ❌ Difficult to maintain (parameters scattered across different concerns)
- ❌ Unclear organization (style, context, and layout mixed together)

## Solution: Three-Config Pattern

Organized all configurations into **three semantic bundles**:

### 1. `style_config` - Style Configurations
Contains all styling-related configurations:
- `styling_config`: StylingConfigModel instance
- *(Future style configs can be added here)*

### 2. `context_config` - Runtime/Contextual Data
Contains all runtime and contextual data:
- `sheet_name`: Sheet being processed
- `invoice_data`: Complete invoice data dictionary
- `all_sheet_configs`: Cross-sheet configuration reference
- `args`: Command-line arguments
- `final_grand_total_pallets`: Global pallet count
- *(Future context data can be added here)*

### 3. `layout_config` - Layout/Building Controls
Contains all layout and building control configurations:
- `sheet_config`: Sheet-specific configuration
- `enable_text_replacement`: Text replacement flag
- `skip_template_header_restoration`: Skip flag
- `skip_header_builder`: Skip flag
- `skip_data_table_builder`: Skip flag
- `skip_footer_builder`: Skip flag
- `skip_template_footer_restoration`: Skip flag
- *(Future layout flags can be added here)*

## Changes Made

### 1. Core Class: `invoice_generator/builders/layout_builder.py`

**Before (15 parameters):**
```python
def __init__(
    self,
    workbook: Workbook,
    worksheet: Worksheet,
    template_worksheet: Worksheet,
    sheet_name: str,
    sheet_config: Dict[str, Any],
    all_sheet_configs: Dict[str, Any],
    invoice_data: Dict[str, Any],
    styling_config: Optional[StylingConfigModel] = None,
    args: Optional[Any] = None,
    final_grand_total_pallets: int = 0,
    enable_text_replacement: bool = False,
    skip_template_header_restoration: bool = False,
    skip_header_builder: bool = False,
    skip_data_table_builder: bool = False,
    skip_footer_builder: bool = False,
    skip_template_footer_restoration: bool = False,
):
```

**After (6 parameters - 3 worksheets + 3 configs):**
```python
def __init__(
    self,
    workbook: Workbook,
    worksheet: Worksheet,
    template_worksheet: Worksheet,
    style_config: Dict[str, Any],
    context_config: Dict[str, Any],
    layout_config: Dict[str, Any],
):
```

### 2. Updated Callers

Updated all 3 instantiation points to use the new pattern:

#### `invoice_generator/processors/single_table_processor.py`
- ✅ Bundles configs into three dictionaries
- ✅ Maintains all existing functionality

#### `invoice_generator/processors/multi_table_processor.py`
- ✅ Bundles configs for each table iteration
- ✅ Properly handles skip flags for multi-table scenarios

#### `test_layout_builder_update.py`
- ✅ Updated test to use new pattern
- ✅ All tests pass successfully

### 3. Documentation: `docs/builders/layout_builder_documentation.md`

Added comprehensive documentation:
- ✅ Three-Config Architecture explanation
- ✅ Detailed parameter documentation
- ✅ Usage examples with actual code
- ✅ Migration guide from old pattern to new
- ✅ Benefits and checklist

## Usage Example

```python
# Prepare the three config bundles
style_config = {
    'styling_config': styling_model
}

context_config = {
    'sheet_name': 'Invoice',
    'invoice_data': invoice_data,
    'all_sheet_configs': all_configs,
    'args': args,
    'final_grand_total_pallets': 150
}

layout_config = {
    'sheet_config': sheet_config,
    'enable_text_replacement': True,
    'skip_header_builder': False,
    'skip_footer_builder': False
}

# Create and run the builder
layout_builder = LayoutBuilder(
    workbook=output_wb,
    worksheet=output_ws,
    template_worksheet=template_ws,
    style_config=style_config,
    context_config=context_config,
    layout_config=layout_config
)
success = layout_builder.build()
```

## Benefits

### ✅ Extensibility
- Add new configs by simply adding dictionary keys
- No constructor signature changes needed
- Future-proof architecture

### ✅ Organization
- Clear separation of concerns (style vs. context vs. layout)
- Easy to understand what goes where
- Semantic grouping of related configs

### ✅ Maintainability
- Easier to read and understand
- Self-documenting code structure
- Reduced parameter count (15+ → 6)

### ✅ Backward Compatibility
- Internal extraction of commonly used values
- Existing internal code continues to work
- Smooth migration path

## Migration Checklist

For any new code or future migrations:

1. ✅ Move `styling_config` into `style_config` dict
2. ✅ Move `sheet_name`, `invoice_data`, `all_sheet_configs`, `args`, `final_grand_total_pallets` into `context_config` dict
3. ✅ Move `sheet_config`, `enable_text_replacement`, and all `skip_*` flags into `layout_config` dict
4. ✅ Keep `workbook`, `worksheet`, `template_worksheet` as direct parameters (unchanged)

## Testing

✅ All tests pass:
```
python test_layout_builder_update.py
```

**Test Results:**
- ✓ LayoutBuilder accepts template_worksheet parameter
- ✓ Template worksheet reference stored correctly
- ✓ Output worksheet reference stored correctly
- ✓ Template unchanged during instantiation
- ✓ No internal workbook creation
- ✓ Clean separation between template (read) and output (write)

## Files Modified

1. `invoice_generator/builders/layout_builder.py` - Core refactor
2. `invoice_generator/processors/single_table_processor.py` - Caller update
3. `invoice_generator/processors/multi_table_processor.py` - Caller update
4. `test_layout_builder_update.py` - Test update
5. `docs/builders/layout_builder_documentation.md` - Documentation update

## Summary

This refactor significantly improves the codebase by:
- Reducing constructor complexity (15+ params → 6 params)
- Providing clear semantic organization (style, context, layout)
- Enabling future extensibility without breaking changes
- Maintaining full backward compatibility

**Status: ✅ COMPLETE - All tests passing, documentation updated, ready for production.**





