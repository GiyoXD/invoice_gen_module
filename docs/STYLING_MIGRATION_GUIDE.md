# Styling Migration Guide: From Section-Based to ID-Driven

## Overview

The styling system is evolving from **section-based styling** (header/data/footer) to **ID-driven styling** (columns + row_contexts). This provides:

- **DRY principle**: Define column format once, reuse across contexts
- **Centralized control**: All styling in one place
- **Type safety**: Column defines WHAT (number, text), row context defines HOW (bold, colors)
- **Override flexibility**: Easy per-cell customization

## Architecture

### Old Format (Section-Based)

```json
{
  "styling_bundle": {
    "Invoice": {
      "header": {
        "font": {"bold": true, "size": 12},
        "alignment": {"horizontal": "center"}
      },
      "data": {
        "font": {"size": 12},
        "alignment": {"horizontal": "center"}
      },
      "footer": {
        "font": {"bold": true, "size": 12},
        "alignment": {"horizontal": "center"}
      },
      "column_specific": {
        "col_cbm": {
          "alignment": {"horizontal": "right"}
        }
      }
    }
  }
}
```

**Problems:**
- Column format (e.g., "0.00" for CBM) repeated in header/data/footer
- Alignment repeated 3 times
- Hard to maintain consistency across sections

### New Format (ID-Driven)

```json
{
  "styling_bundle": {
    "Invoice": {
      "columns": {
        "col_po": {
          "format": "@",
          "alignment": "center",
          "width": 28
        },
        "col_cbm": {
          "format": "0.00",
          "alignment": "right",
          "width": 12
        },
        "col_desc": {
          "format": "@",
          "alignment": "center",
          "wrap_text": true,
          "width": 20
        }
      },
      "row_contexts": {
        "header": {
          "bold": true,
          "font_size": 12,
          "fill_color": "CCCCCC",
          "row_height": 35
        },
        "data": {
          "bold": false,
          "font_size": 12,
          "row_height": 35
        },
        "footer": {
          "bold": true,
          "font_size": 12,
          "font_name": "Times New Roman",
          "row_height": 35
        }
      }
    }
  }
}
```

**Benefits:**
- Column format defined once: `col_cbm` always uses "0.00"
- Row context adds emphasis: `header` makes it bold, `data` keeps it normal
- Cell style = Column base + Row context merged
- Easy overrides: Pass `overrides={'bold': False}` for special cases

## How Cell Styles Are Resolved

### Merge Priority (CRITICAL DESIGN)

```
Cell Style = Column Base → Row Context (NON-CONFLICTING ONLY) → Overrides
```

**Column-Owned Properties** (NEVER overridden by context):
- `format` - Number format (e.g., "@", "0.00", "#,##0")
- `alignment` - Horizontal/vertical alignment
- `width` - Column width
- `wrap_text` - Text wrapping

**Context-Additive Properties** (added by row context):
- `bold`, `italic` - Text emphasis
- `font_size`, `font_name`, `font_color` - Font styling
- `fill_color` - Background color
- `border_style` - Border styling
- `row_height` - Row height

**Why This Matters**: Column properties define WHAT the data is (CBM is always "0.00", always right-aligned). Row context adds HOW it looks (header is bold, data is normal). The column's format/alignment should NEVER change between header/data/footer rows.

Example for header cell in description column:
```python
# Step 1: Get column base
col_desc = {
    "format": "@",
    "alignment": "left",    # Column-owned
    "width": 20
}

# Step 2: Get row context
header = {
    "bold": True,
    "alignment": "center",  # This will be IGNORED (column owns alignment)
    "fill_color": "CCCCCC"
}

# Step 3: Merge (context CANNOT override column-owned properties)
merged = {
    "format": "@",          # From column (protected)
    "alignment": "left",    # From column (protected, NOT overridden by context)
    "width": 20,            # From column (protected)
    "bold": True,           # From context (additive)
    "fill_color": "CCCCCC"  # From context (additive)
}

# Step 4: Apply overrides (can override ANYTHING, even column-owned)
if overrides:
    merged.update(overrides)
```

### Builder Usage

**HeaderBuilder** (uses 'header' context):
```python
style = registry.get_style('col_cbm', context='header')
# Returns: {"format": "0.00", "alignment": "right", "bold": True, "fill_color": "CCCCCC"}
styler.apply(cell, style)
```

**DataTableBuilder** (uses 'data' context):
```python
style = registry.get_style('col_cbm', context='data')
# Returns: {"format": "0.00", "alignment": "right", "bold": False}
styler.apply(cell, style)
```

**FooterBuilder** (uses 'footer' context):
```python
style = registry.get_style('col_cbm', context='footer')
# Returns: {"format": "0.00", "alignment": "right", "bold": True}
styler.apply(cell, style)
```

## Migration Steps

### 1. Identify Common Column Properties

Extract properties that **describe the data** (not the emphasis):

**Column Base Properties:**
- `format`: Number format (e.g., "@" for text, "0.00" for decimals)
- `alignment`: Horizontal/vertical alignment
- `width`: Column width in Excel units
- `wrap_text`: Whether to wrap text

**Row Context Properties:**
- `bold`, `italic`: Text emphasis
- `font_size`, `font_name`, `font_color`: Font styling
- `fill_color`: Background color
- `border_style`: Border styling
- `row_height`: Row height

### 2. Convert Section-Based to Columns

**Old Config:**
```json
{
  "header": {
    "font": {"bold": true, "size": 12},
    "alignment": {"horizontal": "center"}
  },
  "data": {
    "font": {"size": 12},
    "alignment": {"horizontal": "center"}
  },
  "column_specific": {
    "col_cbm": {
      "alignment": {"horizontal": "right"}
    }
  },
  "dimensions": {
    "column_widths": {
      "col_cbm": 12
    }
  }
}
```

**New Config:**
```json
{
  "columns": {
    "col_cbm": {
      "format": "0.00",
      "alignment": "right",
      "width": 12
    },
    "col_default": {
      "format": "@",
      "alignment": "center",
      "width": 15
    }
  },
  "row_contexts": {
    "header": {
      "bold": true,
      "font_size": 12
    },
    "data": {
      "font_size": 12
    },
    "footer": {
      "bold": true,
      "font_size": 12
    }
  }
}
```

### 3. Use Default Column

For columns without specific formatting, define a **default column**:

```json
{
  "columns": {
    "col_default": {
      "format": "@",
      "alignment": "center",
      "width": 15
    },
    "col_po": {
      "format": "@",
      "alignment": "center",
      "width": 28
    },
    "col_cbm": {
      "format": "0.00",
      "alignment": "right",
      "width": 12
    }
  }
}
```

The StyleRegistry will fall back to `col_default` if a specific column ID is not found.

## Backward Compatibility

The new styling system **does NOT break existing configs**. Builders check for the new format and fall back to legacy styling:

```python
# In HeaderBuilder, DataTableBuilder, FooterBuilder
if self.style_registry and col_id:
    # NEW: Use StyleRegistry
    style = self.style_registry.get_style(col_id, context='header')
    self.cell_styler.apply(cell, style)
else:
    # LEGACY: Use old styling methods
    apply_header_style(cell, self.sheet_styling_config)
    apply_cell_style(cell, self.sheet_styling_config, context)
```

**Migration Strategy:**
1. Keep old configs working (no changes needed)
2. Add new configs gradually (sheet by sheet)
3. Test both formats work correctly
4. Eventually deprecate old format (Phase 4)

## Example: Complete Sheet Config

```json
{
  "styling_bundle": {
    "Invoice": {
      "columns": {
        "col_default": {
          "format": "@",
          "alignment": "center",
          "wrap_text": true,
          "width": 15
        },
        "col_po": {
          "format": "@",
          "alignment": "center",
          "width": 28
        },
        "col_desc": {
          "format": "@",
          "alignment": "left",
          "wrap_text": true,
          "width": 20
        },
        "col_cbm": {
          "format": "0.00",
          "alignment": "right",
          "width": 12
        },
        "col_qty_pcs": {
          "format": "#,##0",
          "alignment": "right",
          "width": 12
        },
        "col_unit_price": {
          "format": "0.00",
          "alignment": "right",
          "width": 15
        },
        "col_amount": {
          "format": "#,##0.00",
          "alignment": "right",
          "width": 18
        }
      },
      "row_contexts": {
        "header": {
          "bold": true,
          "font_size": 12,
          "font_name": "Times New Roman",
          "fill_color": "CCCCCC",
          "border_style": "thin",
          "row_height": 35
        },
        "data": {
          "bold": false,
          "font_size": 12,
          "font_name": "Times New Roman",
          "border_style": "thin",
          "row_height": 35
        },
        "footer": {
          "bold": true,
          "font_size": 12,
          "font_name": "Times New Roman",
          "border_style": "thin",
          "row_height": 35
        },
        "static_row": {
          "bold": false,
          "italic": true,
          "font_size": 11,
          "fill_color": "F0F0F0"
        }
      }
    }
  }
}
```

## Testing Strategy

1. **Unit Tests**: Test StyleRegistry and CellStyler (already complete)
2. **Integration Tests**: Test builders with new config format
3. **Regression Tests**: Verify old configs still work
4. **Visual Tests**: Compare output Excel files (old vs new styling)

## Next Steps

1. ✅ **Phase 1 Complete**: Core styling system (StyleRegistry + CellStyler)
2. ✅ **Phase 2 Complete**: Integrate into builders (HeaderBuilder, DataTableBuilder, FooterBuilder)
3. ⏳ **Phase 3 Pending**: Create example configs with new format
4. ⏳ **Phase 4 Pending**: Add backward compatibility adapter
5. ⏳ **Phase 5 Pending**: Test all sheets render correctly
6. ⏳ **Phase 6 Pending**: Update documentation and deprecate old format
