# Bundle Cascade Pattern - Pure Bundle Architecture

## Overview

The **Bundle Cascade Pattern** is an architectural pattern where configuration bundles flow down from a parent component (Director) to child components (Builders) without extraction or transformation. Child components use `@property` decorators to access bundle values on-demand.

**Status:** ✅ Implemented in LayoutBuilder → DataTableBuilder, FooterBuilder

---

## The Pattern

### Core Principle

> **"Store bundles, access on-demand"** - Don't extract config values in constructors; access them lazily via properties.

```python
# ❌ OLD: Extract everything in constructor (50+ lines)
class Builder:
    def __init__(self, worksheet, sheet_name, config, data, styling, ...):  # 20+ params
        self.worksheet = worksheet
        self.sheet_name = sheet_name
        self.config = config
        # ... 20+ more assignments ...

# ✅ NEW: Store bundles, extract via properties
class Builder:
    def __init__(self, worksheet, style_config, context_config, layout_config, data_config):
        self.worksheet = worksheet
        self.style_config = style_config
        self.context_config = context_config
        self.layout_config = layout_config
        self.data_config = data_config
        # Done! 5 lines instead of 50!
    
    @property
    def sheet_name(self):
        return self.context_config.get('sheet_name', '')
```

---

## Architecture Diagram

```
┌─────────────────────────────────────────────────────────┐
│              LayoutBuilder (Director)                    │
│  Receives:                                               │
│    • style_config    ─┐                                  │
│    • context_config  ─┼─► Stores & cascades down        │
│    • layout_config   ─┤                                  │
│    • (creates data)  ─┘                                  │
└──────────────┬──────────────────────────────────────────┘
               │
               │ CASCADES BUNDLES ↓
               │
    ┌──────────┴──────────┬──────────────┬───────────────┐
    ▼                     ▼              ▼               ▼
┌────────────┐     ┌─────────────┐  ┌──────────┐  ┌─────────────┐
│DataTable   │     │ Footer      │  │ Header   │  │ Template    │
│Builder     │     │ Builder     │  │ Builder  │  │ State       │
├────────────┤     ├─────────────┤  ├──────────┤  ├─────────────┤
│• style     │     │• style      │  │• style   │  │• (simple,   │
│• context   │     │• context    │  │• layout  │  │  no bundles)│
│• layout    │     │• data       │  │          │  │             │
│• data      │     │             │  │          │  │             │
├────────────┤     ├─────────────┤  ├──────────┤  ├─────────────┤
│@property   │     │@property    │  │(4 params │  │             │
│sheet_name  │     │header_info  │  │ no       │  │             │
│data_source │     │pallet_count │  │ bundles) │  │             │
│styling     │     │styling      │  │          │  │             │
│... (20+)   │     │... (13)     │  │          │  │             │
└────────────┘     └─────────────┘  └──────────┘  └─────────────┘
     │                    │              │              │
     └────────────────────┴──────────────┴──────────────┘
                          │
              Each extracts what IT needs via @property
```

---

## Standard Bundle Taxonomy

### 1. `style_config` - Style Configuration Bundle
**Purpose:** All styling and visual configurations

**Contents:**
```python
{
    'styling_config': StylingConfigModel instance
}
```

**Used by:** All builders

---

### 2. `context_config` - Runtime Context Bundle
**Purpose:** Runtime state, business context, and session data

**Contents:**
```python
{
    'sheet_name': str,
    'invoice_data': Dict,
    'all_sheet_configs': Dict,
    'args': CommandLineArgs,
    'final_grand_total_pallets': int,
    'header_info': Dict,           # (for child builders)
    'pallet_count': int,            # (for child builders)
    'is_last_table': bool,          # (for child builders)
    'dynamic_desc_used': bool       # (for child builders)
}
```

**Used by:** LayoutBuilder, DataTableBuilder, FooterBuilder

---

### 3. `layout_config` - Layout Control Bundle
**Purpose:** Layout positioning, spacing, and structural controls

**Contents:**
```python
{
    'sheet_config': Dict,
    'add_blank_after_header': bool,
    'static_content_after_header': Dict,
    'add_blank_before_footer': bool,
    'static_content_before_footer': Dict,
    'merge_rules_after_header': Dict,
    'merge_rules_before_footer': Dict,
    'merge_rules_footer': Dict,
    'data_cell_merging_rules': Dict,
    'max_rows_to_fill': int,
    'enable_text_replacement': bool,
    'skip_*': bool  # Various skip flags
}
```

**Used by:** LayoutBuilder, DataTableBuilder

---

### 4. `data_config` - Data Source Bundle
**Purpose:** Data sources, mappings, and data-related configs

**Contents:**
```python
{
    'data_source': Union[Dict, List],
    'data_source_type': str,
    'header_info': Dict,
    'mapping_rules': Dict,
    'all_tables_data': Dict,
    'table_keys': List,
    'sum_ranges': List[Tuple[int, int]],
    'footer_config': Dict,
    'DAF_mode': bool,
    'override_total_text': str
}
```

**Used by:** DataTableBuilder, FooterBuilder

---

## Implementation Examples

### Example 1: LayoutBuilder (Director)

```python
class LayoutBuilder:
    def __init__(
        self,
        workbook: Workbook,
        worksheet: Worksheet,
        template_worksheet: Worksheet,
        style_config: Dict[str, Any],
        context_config: Dict[str, Any],
        layout_config: Dict[str, Any],
    ):
        # Store worksheets and bundles (NO extraction!)
        self.workbook = workbook
        self.worksheet = worksheet
        self.template_worksheet = template_worksheet
        self.style_config = style_config
        self.context_config = context_config
        self.layout_config = layout_config
        
        # Extract ONLY what LayoutBuilder itself needs
        self.sheet_name = context_config.get('sheet_name')
        self.sheet_config = layout_config.get('sheet_config', {})
        # ... minimal extraction ...
    
    def build(self):
        # Create data config bundle for this build
        data_config_bundle = {
            'data_source': computed_data,
            'data_source_type': 'aggregation',
            'header_info': self.header_info,
            'mapping_rules': self.sheet_config.get('mappings', {})
        }
        
        # CASCADE: Pass bundles to child builder
        data_table_builder = DataTableBuilder(
            worksheet=self.worksheet,
            style_config=self.style_config,      # ← CASCADE
            context_config=self.context_config,  # ← CASCADE
            layout_config=self.layout_config,    # ← CASCADE
            data_config=data_config_bundle       # ← CREATE & PASS
        )
        data_table_builder.build()
```

---

### Example 2: DataTableBuilder (Worker)

```python
class DataTableBuilder:
    def __init__(
        self,
        worksheet: Worksheet,
        style_config: Dict[str, Any],
        context_config: Dict[str, Any],
        layout_config: Dict[str, Any],
        data_config: Dict[str, Any]
    ):
        # Store worksheet and bundles (NO extraction!)
        self.worksheet = worksheet
        self.style_config = style_config
        self.context_config = context_config
        self.layout_config = layout_config
        self.data_config = data_config
        
        # Initialize output state variables (NOT from configs!)
        self.data_start_row = -1
        self.data_end_row = -1
        self.dynamic_desc_used = False
    
    # Properties for frequently accessed values
    @property
    def sheet_name(self) -> str:
        return self.context_config.get('sheet_name', '')
    
    @property
    def data_source(self):
        return self.data_config.get('data_source')
    
    @property
    def sheet_styling_config(self):
        return self.style_config.get('styling_config')
    
    # ... 20+ more properties ...
    
    def build(self):
        # Use properties for clean access
        if self.sheet_name == 'Invoice':
            data = self.data_source
            styling = self.sheet_styling_config
            # ... use them naturally ...
```

---

## Benefits

### 1. Tiny Constructors
- **Before:** 50+ lines of parameter extraction
- **After:** 5-10 lines to store bundles
- **Reduction:** 80-90% less constructor code

### 2. Zero Maintenance for New Configs
- **Before:** Add config → update constructor → update all callers
- **After:** Add config → done! Automatically available via bundle

### 3. Consistent Interface
- All builders use same bundle names
- Easy to remember: `style_config`, `context_config`, `layout_config`, `data_config`
- Predictable pattern across codebase

### 4. Clean Access Syntax
- Use `@property` for frequently accessed values: `self.sheet_name`
- Access bundles directly for rare values: `self.layout_config.get('rare_flag')`

### 5. Single Source of Truth
- Config stored once in bundle
- No duplication (bundle + extracted)
- Less memory, less confusion

---

## When to Use Properties vs Direct Access

### Use `@property` for:
✅ Values used **3+ times** across methods
✅ Core builder functionality (sheet_name, data_source, styling)
✅ Values that make code cleaner (better than dict access everywhere)

### Use Direct Access for:
✅ Values used **1-2 times** total
✅ Rare edge cases or optional features
✅ Values only needed in one specific method

**Example:**
```python
# Property: Used everywhere
@property
def sheet_name(self):
    return self.context_config.get('sheet_name', '')

# Direct access: Used once
def _rare_feature(self):
    if self.layout_config.get('enable_rare_feature', False):
        ...
```

---

## Migration Checklist

When converting a builder to pure bundle pattern:

### Phase 1: Update Constructor
- [ ] Change params from 15+ individuals to 4 bundles
- [ ] Store bundles only (no extraction)
- [ ] Keep output state variables (those are legitimate)

### Phase 2: Add Properties
- [ ] Add `@property` for frequently accessed values (3-5 uses)
- [ ] Document what each property returns
- [ ] Test that properties work correctly

### Phase 3: Update Callers
- [ ] Find all places that instantiate this builder
- [ ] Convert to bundle passing
- [ ] Test each caller

### Phase 4: Verify
- [ ] Run linter (no errors)
- [ ] Run tests (all pass)
- [ ] Update documentation

---

## Real-World Stats

### DataTableBuilder Transformation
- **Before:** 24 individual parameters
- **After:** 4 bundle parameters
- **Constructor lines:** 50 → 10 (80% reduction)
- **Properties added:** 21 (for clean access)
- **Result:** Massively cleaner, infinitely extensible

### FooterBuilder Transformation
- **Before:** 15 individual parameters
- **After:** 3 bundle parameters + footer_row_num
- **Constructor lines:** 35 → 10 (71% reduction)
- **Properties added:** 13 (for clean access)
- **Result:** Clean and maintainable

### LayoutBuilder Update
- **Before:** 60+ lines extracting and passing params to children
- **After:** 20 lines creating and passing bundles
- **Reduction:** 67% less mapping code

---

## Future Extensions

### Adding New Config (Example)

**Want to add `enable_watermark` flag?**

**OLD WAY (Parameter Hell):**
1. ❌ Add param to DataTableBuilder constructor
2. ❌ Update all 3 call sites
3. ❌ Update documentation (3 places)
4. ❌ Risk breaking existing code
5. ❌ 30 minutes of work

**NEW WAY (Bundle Paradise):**
1. ✅ Add to appropriate bundle at call site:
```python
layout_config = {
    # ... existing ...
    'enable_watermark': True  # ← JUST ADD THIS!
}
```
2. ✅ Access in builder:
```python
def build(self):
    if self.layout_config.get('enable_watermark', False):
        self._add_watermark()
```
3. ✅ Done! 2 minutes of work

---

## Best Practices

### DO ✅
- Store bundles as-is in constructor
- Use `@property` for frequently accessed values (3+ uses)
- Access bundles directly for rare values (1-2 uses)
- Keep bundle names consistent across all builders
- Document what each bundle contains

### DON'T ❌
- Extract everything in constructor (defeats the purpose!)
- Mix bundle types (style in context, etc.)
- Forget to provide defaults in `.get()`
- Rename bundles (use standard: style, context, layout, data)

---

## Summary

The **Bundle Cascade Pattern** transforms parameter-heavy constructors into clean, extensible builders by:

1. **Cascading bundles** from parent to children
2. **Storing bundles** without extraction
3. **Using properties** for clean access
4. **Maintaining single source of truth**

**Result:** 80% less constructor code, infinite extensibility, zero maintenance for new configs.

**Status:** ✅ Fully implemented in DataTableBuilder and FooterBuilder
**Next:** Consider applying to other complex components

---

**See Also:**
- `docs/REFACTOR_THREE_CONFIG_PATTERN.md` - Original three-config refactor
- `docs/builders/data_table_builder_documentation.md` - DataTableBuilder details
- `docs/builders/footer_builder_documentation.md` - FooterBuilder details





