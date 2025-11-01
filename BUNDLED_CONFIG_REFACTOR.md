# Bundled Config Refactor - Summary

## Overview

The invoice generator has been refactored to support the new **bundled config format (v2.0+)** while maintaining backward compatibility with the old config format.

## What Was Changed

### 1. Config Loader (`invoice_generator/config/loader.py`)

Added a new `BundledConfigLoader` class that:
- Reads the new bundled config format with sections:
  - `_meta`: Config metadata (version, customer, etc.)
  - `processing`: Sheets to process and data sources
  - `layout_bundle`: Layout configuration per sheet (structure, data_flow, content, footer)
  - `styling_bundle`: Styling configuration per sheet
  - `features`: Feature flags
  - `defaults`: Default settings
  
- Provides clean accessor methods for each section
- **Converts bundled config to legacy format** using `build_legacy_sheet_config()` method
  - Transforms `structure.columns` → `header_to_write`
  - Transforms `data_flow.mappings` → `mappings` with format support
  - Extracts static content and footer config

### 2. Main Generator (`invoice_generator/generate_invoice.py`)

Updated to:
- **Auto-detect config format** by checking for `_meta` and `config_version` fields
- Use `BundledConfigLoader` when bundled config is detected
- Convert bundled config to legacy format for processing
- Maintain full backward compatibility with old configs

### 3. Path Resolution

Updated `derive_paths()` to:
- Check for bundled configs in `config_bundled/` directory first
- Look for files named `{name}_bundled_v2.json`
- Fall back to regular configs if bundled not found

## How It Works

```
Bundled Config (new format)
    ↓
BundledConfigLoader
    ↓
build_legacy_sheet_config()
    ↓
Legacy Config Format
    ↓
Existing Processors & Builders
    (no changes needed!)
```

This approach provides a **gradual refactoring path** - the bundled config is converted to the legacy format, so all existing processors and builders continue to work without modification.

## Config Format Comparison

### Old Format
```json
{
  "sheets_to_process": ["Invoice", "Contract"],
  "sheet_data_map": {
    "Invoice": "aggregation"
  },
  "data_mapping": {
    "Invoice": {
      "start_row": 21,
      "header_to_write": [...],
      "mappings": {...},
      "styling": {...}
    }
  }
}
```

### New Bundled Format
```json
{
  "_meta": {
    "config_version": "2.1_developer_optimized",
    "customer": "JF"
  },
  "processing": {
    "sheets": ["Invoice", "Contract"],
    "data_sources": {
      "Invoice": "aggregation"
    }
  },
  "layout_bundle": {
    "Invoice": {
      "structure": {
        "start_row": 21,
        "columns": [...]
      },
      "data_flow": {
        "mappings": {...}
      }
    }
  },
  "styling_bundle": {
    "Invoice": {...}
  }
}
```

## Usage

### With Auto-Detection (Recommended)

The generator will automatically detect and use bundled configs:

```bash
python -m invoice_generator.generate_invoice JF.json
```

The path resolution will check:
1. `config_bundled/JF_bundled_v2.json` (preferred)
2. `config/JF_config.json` (fallback)

### Explicit Config Path

You can also specify the config path directly:

```bash
python -m invoice_generator.generate_invoice JF.json -c ./config_bundled/
```

## Testing

The refactor was tested with `JF_bundled_v2.json` and successfully:
- ✅ Loaded bundled config with version 2.1_developer_optimized
- ✅ Extracted customer info (JF)
- ✅ Identified 3 sheets to process (Invoice, Contract, Packing list)
- ✅ Mapped data sources correctly
- ✅ Converted all sheets to legacy format with proper:
  - start_row values
  - Header definitions (7-10 headers per sheet)
  - Mappings with number formats (6-8 mappings per sheet)
  - Styling and footer configs

## Benefits

1. **Backward Compatible**: Old configs still work
2. **Gradual Migration**: Can convert configs one at a time
3. **No Changes to Core Logic**: Processors and builders unchanged
4. **Clean Separation**: Config loading isolated in loader module
5. **Better Organization**: New format is more DRY and maintainable

## Next Steps (Future Improvements)

When ready, you can:
1. Update processors to directly use bundled config accessors
2. Update builders to read from layout_bundle/styling_bundle directly
3. Remove legacy format conversion layer
4. Delete old config files

But for now, the system works with both formats seamlessly!

## Files Modified

1. `invoice_generator/config/loader.py` - Added BundledConfigLoader class
2. `invoice_generator/generate_invoice.py` - Updated to detect and use bundled configs

## Files Created

- `config_bundled/JF_bundled_v2.json` - Example bundled config (already existed)

---

**Status**: ✅ Complete - Ready to use bundled configs!




