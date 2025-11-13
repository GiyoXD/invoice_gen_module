# Bundled Config Quick Start Guide

## Running with Bundled Config

### Option 1: Auto-Detection (Easiest)
The generator automatically finds bundled configs in `config_bundled/`:

```bash
python -m invoice_generator.generate_invoice JF.json
```

**What happens:**
1. Looks for `config_bundled/JF_bundled_v2.json` [OK]
2. Falls back to `config/JF_config.json` if not found
3. Auto-detects format and processes accordingly

### Option 2: Specify Directory
```bash
python -m invoice_generator.generate_invoice JF.json -c ./invoice_generator/config_bundled/
```

### Option 3: Full Paths
```bash
python -m invoice_generator.generate_invoice \
  ./invoice_generator/JF.json \
  -t ./invoice_generator/template/ \
  -c ./invoice_generator/config_bundled/ \
  -o ./output/JF_result.xlsx
```

## Bundled Config Structure

```
config_bundled/
└── JF_bundled_v2.json
    ├── _meta                    # Config metadata
    ├── processing               # Sheets & data sources
    │   ├── sheets              # List of sheets to process
    │   └── data_sources        # Map: sheet → data source type
    ├── layout_bundle           # Layout per sheet
    │   └── {SheetName}
    │       ├── structure       # Columns, start_row
    │       ├── data_flow       # Field mappings
    │       ├── content         # Static content
    │       └── footer          # Footer configuration
    ├── styling_bundle          # Styling per sheet
    │   └── {SheetName}
    │       ├── header          # Header styles
    │       ├── data            # Data row styles
    │       ├── footer          # Footer styles
    │       └── dimensions      # Column widths
    ├── features                # Feature flags
    └── defaults                # Default settings
```

## Key Differences from Old Config

| Aspect | Old Config | New Bundled Config |
|--------|-----------|-------------------|
| **Sheets to process** | `sheets_to_process` | `processing.sheets` |
| **Data sources** | `sheet_data_map` | `processing.data_sources` |
| **Start row** | `data_mapping.{sheet}.start_row` | `layout_bundle.{sheet}.structure.start_row` |
| **Headers** | `header_to_write` (list) | `structure.columns` (list with metadata) |
| **Mappings** | `mappings` (complex) | `data_flow.mappings` (simplified) |
| **Styling** | Mixed in sheet config | Separate `styling_bundle.{sheet}` |
| **Static content** | `static_content` | `content.static` |

## Example: Adding a New Sheet

In bundled config:

```json
{
  "processing": {
    "sheets": ["Invoice", "NewSheet"],
    "data_sources": {
      "NewSheet": "aggregation"
    }
  },
  "layout_bundle": {
    "NewSheet": {
      "structure": {
        "start_row": 10,
        "columns": [
          {"id": "col_po", "header": "PO Number", "format": "@"}
        ]
      },
      "data_flow": {
        "mappings": {
          "po": {"column": "col_po", "source_key": 0}
        }
      }
    }
  },
  "styling_bundle": {
    "NewSheet": {
      "header": {
        "font": {"bold": true, "size": 12}
      }
    }
  }
}
```

## Column Definition (New Format)

### Simple Column
```json
{
  "id": "col_po",
  "header": "P.O. №",
  "format": "@",           // Text format
  "rowspan": 1,
  "colspan": 1
}
```

### Column with Children (Multi-level header)
```json
{
  "id": "col_qty_header",
  "header": "Quantity(SF)",
  "colspan": 2,
  "children": [
    {"id": "col_qty_pcs", "header": "PCS", "format": "#,##0"},
    {"id": "col_qty_sf", "header": "SF", "format": "#,##0.00"}
  ]
}
```

## Mapping Definition (New Format)

### From Data Key
```json
"po": {
  "column": "col_po",
  "source_key": 0,           // Index in data tuple
  "fallback": "N/A"
}
```

### From Aggregated Value
```json
"sqft": {
  "column": "col_qty_sf",
  "source_value": "sqft_sum" // Key in aggregation result
}
```

### Formula
```json
"amount": {
  "column": "col_amount",
  "formula": "{col_qty_sf} * {col_unit_price}"
}
```

## Checking Config Status

To verify your config works:

```python
from invoice_generator.config.loader import load_bundled_config

loader = load_bundled_config("config_bundled/JF_bundled_v2.json")

print(f"Version: {loader.config_version}")
print(f"Customer: {loader.customer}")
print(f"Sheets: {loader.get_sheets_to_process()}")

# Get config for a sheet
invoice_config = loader.build_legacy_sheet_config("Invoice")
print(f"Start row: {invoice_config['start_row']}")
```

## Migration Strategy

1. ✅ **Current State**: Both formats work
2. **Phase 1**: Convert one config at a time to bundled format
3. **Phase 2**: Test each converted config
4. **Phase 3**: Once all configs converted, can remove legacy support

## Features

### Enabled/Disabled
```json
"features": {
  "enable_text_replacement": false,
  "enable_auto_calculations": true,
  "debug_mode": false
}
```

Check in code:
```python
if config_loader.is_feature_enabled("debug_mode"):
    print("Debug info...")
```

## Troubleshooting

### Config not found
- Check file is in `config_bundled/` directory
- Check filename matches pattern: `{name}_bundled_v2.json`
- Use `-c` flag to specify directory

### Format errors
- Ensure `_meta` section exists with `config_version`
- Validate JSON syntax
- Check column IDs match between structure and mappings

### Data not showing
- Verify `data_flow.mappings` column IDs match `structure.columns` IDs
- Check `source_key` indices match your data structure
- Ensure data source type in `processing.data_sources` is correct

---

**Pro Tip**: Keep the old config as backup until you verify the bundled config works correctly!




