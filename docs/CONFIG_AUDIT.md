# Configuration Structure Audit

**Date**: October 29, 2025  
**Purpose**: Document current configuration structure to guide refactoring towards typed configuration classes  
**Status**: Analysis of 22 config files in `invoice_generator/config/`

---

## Executive Summary

### Current Problems Identified

1. **High Parameter Count**: Builders accept 14-24 parameters each
2. **Configuration Fragmentation**: Related settings scattered across multiple dictionaries
3. **No Type Safety**: Dictionary-based configs offer no IDE support or validation
4. **Hidden Dependencies**: Typos in config keys fail silently
5. **Maintenance Burden**: Adding new config requires updating multiple call sites

### Recommendations

1. **Create Typed Configuration Classes** using `@dataclass`
2. **Consolidate Related Settings** into logical groups
3. **Implement Validation** in class constructors
4. **Provide Migration Path** from JSON → typed classes

---

## Configuration File Structure

### Top-Level Structure

All config files follow this pattern:

```json
{
  "sheets_to_process": ["Invoice", "Contract", "Packing list"],
  "sheet_data_map": { ... },
  "data_mapping": {
    "Invoice": { ... },
    "Contract": { ... },
    "Packing list": { ... }
  }
}
```

**Fields:**
- `sheets_to_process` (List[str]): Sheet names to process
- `sheet_data_map` (Dict[str, str]): Maps sheet names to data source types
- `data_mapping` (Dict[str, SheetConfig]): Per-sheet configurations

---

## Sheet Configuration Schema

Each sheet in `data_mapping` has this structure:

### Core Layout Settings

| Field | Type | Required | Default | Description | Example |
|-------|------|----------|---------|-------------|---------|
| `start_row` | int | ✅ Yes | - | Row where data starts (after header) | `21` |
| `header_to_write` | List[HeaderCell] | ✅ Yes | - | Header cell definitions | See below |
| `mappings` | Dict[str, Mapping] | ✅ Yes | - | Column mapping rules | See below |
| `footer_configurations` | FooterConfig | ✅ Yes | - | Footer settings | See below |
| `styling` | StylingConfig | ✅ Yes | - | Styling rules | See below |

### Optional Layout Settings

| Field | Type | Required | Default | Description | Usage Frequency |
|-------|------|----------|---------|-------------|-----------------|
| `add_blank_before_footer` | bool | ❌ No | `false` | Insert blank row before footer | **High** (15/22) |
| `add_blank_after_header` | bool | ❌ No | `false` | Insert blank row after header | Low (2/22) |
| `static_content_before_footer` | Dict[str, str] | ❌ No | `{}` | Static text before footer | **High** (18/22) |
| `static_content_after_header` | Dict[str, str] | ❌ No | `{}` | Static text after header | Low (1/22) |
| `merge_rules_before_footer` | Dict[str, int] | ❌ No | `{}` | Cell merges before footer | Medium (8/22) |
| `merge_rules_after_header` | Dict[str, int] | ❌ No | `{}` | Cell merges after header | Low (1/22) |
| `merge_rules_footer` | Dict[str, int] | ❌ No | `{}` | Cell merges in footer | Low (3/22) |
| `data_cell_merging_rule` | Dict[str, Dict] | ❌ No | `null` | Data cell merge rules | **High** (14/22) |
| `weight_summary_config` | WeightSummaryConfig | ❌ No | `null` | Weight summary settings | Medium (6/22) |
| `summary` | bool | ❌ No | `false` | Enable summary section (DAF mode) | Low (Packing list only) |
| `data_source` | str | ❌ No | From `sheet_data_map` | Override data source | Low (2/22) |

---

## Detailed Schema Definitions

### 1. HeaderCell Schema

```python
# Current JSON structure
{
  "row": int,           # 0-based row offset
  "col": int,           # 0-based column offset
  "text": str,          # Display text
  "id": str,            # Column identifier (e.g., "col_desc")
  "rowspan": int,       # Optional, default 1
  "colspan": int        # Optional, default 1
}
```

**Example:**
```json
{
  "row": 0,
  "col": 3,
  "text": "Description",
  "id": "col_desc",
  "rowspan": 2,
  "colspan": 1
}
```

**Proposed Class:**
```python
@dataclass
class HeaderCell:
    row: int
    col: int
    text: str
    id: Optional[str] = None
    rowspan: int = 1
    colspan: int = 1
```

### 2. Mapping Rules Schema

Mappings have different types based on purpose:

#### Type 1: Key Index Mapping (Dynamic Data)
```json
{
  "po": {
    "key_index": 0,
    "id": "col_po",
    "number_format": "@"  // Optional
  }
}
```

#### Type 2: Value Key Mapping (Aggregated Data)
```json
{
  "sqft": {
    "value_key": "sqft_sum",
    "id": "col_qty_sf",
    "number_format": "#,##0.00"
  }
}
```

#### Type 3: Formula Mapping
```json
{
  "amount": {
    "id": "col_amount",
    "type": "formula",
    "formula_template": "{col_ref_1}{row} * {col_ref_0}{row}",
    "inputs": ["col_qty_sf", "col_unit_price"],
    "number_format": "#,##0.00"
  }
}
```

#### Type 4: Initial Static Rows
```json
{
  "initial_static": {
    "type": "initial_static_rows",
    "column_header_id": "col_static",
    "values": ["VENDOR#:", "Des: LEATHER", "MADE IN CAMBODIA"]
  }
}
```

#### Type 5: Description with Fallbacks
```json
{
  "description": {
    "key_index": 3,
    "id": "col_desc",
    "fallback_on_none": "LEATHER",
    "fallback_on_DAF": "LEATHER"  // Optional
  }
}
```

**Proposed Class:**
```python
@dataclass
class MappingRule:
    id: str
    type: Literal["key_index", "value_key", "formula", "initial_static"] = "key_index"
    
    # For key_index and value_key types
    key_index: Optional[int] = None
    value_key: Optional[str] = None
    
    # For formula type
    formula_template: Optional[str] = None
    inputs: Optional[List[str]] = None
    
    # For initial_static type
    column_header_id: Optional[str] = None
    values: Optional[List[str]] = None
    
    # Common fields
    number_format: Optional[str] = None
    fallback_on_none: Optional[str] = None
    fallback_on_DAF: Optional[str] = None
```

### 3. Footer Configuration Schema

```python
# Current JSON structure
{
  "type": "regular" | "grand_total",  # Optional, default "regular"
  "total_text": str,                   # e.g., "TOTAL:"
  "total_text_column_id": str | int,   # Column ID or 0-based index
  "pallet_count_column_id": str | int, # Optional
  "sum_column_ids": List[str],         # Column IDs to sum
  "number_formats": Dict[str, Dict],   # Number formats per column
  "merge_rules": List[MergeRule],      # Footer merge rules
  "style": StyleDict,                  # Footer styling
  "add_ons": List[str]                 # e.g., ["summary"]
}
```

**Example:**
```json
{
  "total_text": "TOTAL:",
  "total_text_column_id": 1,
  "pallet_count_column_id": 2,
  "sum_column_ids": ["col_qty_sf", "col_amount"],
  "number_formats": {
    "col_qty_sf": {"number_format": "#,##0.00"},
    "col_amount": {"number_format": "#,##0.00"}
  },
  "merge_rules": [
    {"start_column_id": "col_po", "colspan": 1}
  ],
  "style": {
    "font": {"name": "Times New Roman", "size": 12, "bold": true},
    "alignment": {"horizontal": "center", "vertical": "center"},
    "border": {"apply": true}
  }
}
```

**Proposed Class:**
```python
@dataclass
class FooterMergeRule:
    start_column_id: Union[str, int]
    colspan: int

@dataclass
class FooterConfig:
    total_text: str
    total_text_column_id: Union[str, int]
    sum_column_ids: List[str]
    number_formats: Dict[str, Dict[str, str]]
    
    type: Literal["regular", "grand_total"] = "regular"
    pallet_count_column_id: Optional[Union[str, int]] = None
    merge_rules: List[FooterMergeRule] = field(default_factory=list)
    style: Optional[Dict[str, Any]] = None
    add_ons: List[str] = field(default_factory=list)
```

### 4. Styling Configuration Schema

**Already partially implemented as `StylingConfigModel`** but currently passed as dict.

```python
# Current JSON structure
{
  "default_font": {"name": str, "size": float, "bold": bool},
  "header_font": {"name": str, "size": float, "bold": bool},
  "default_alignment": {"horizontal": str, "vertical": str, "wrap_text": bool},
  "header_alignment": {"horizontal": str, "vertical": str, "wrap_text": bool},
  "force_text_format_ids": List[str],
  "column_ids_with_full_grid": List[str],
  "column_id_styles": Dict[str, Dict],
  "column_id_widths": Dict[str, float],
  "row_heights": {"header": float, "data_default": float, "footer": float}
}
```

**Current Model** (already exists in `invoice_generator/styling/models.py`):
```python
class StylingConfigModel(BaseModel):
    defaultFont: Optional[FontModel] = None
    headerFont: Optional[FontModel] = None
    defaultAlignment: Optional[AlignmentModel] = None
    headerAlignment: Optional[AlignmentModel] = None
    # ... etc
```

**Issue**: Config uses snake_case, model uses camelCase!

### 5. Data Cell Merging Rule Schema

```python
# Current JSON structure
{
  "col_item": {"rowspan": 1},
  "col_desc": {"rowspan": 1}
}
```

**Proposed Class:**
```python
@dataclass
class DataCellMergeRule:
    rowspan: int
    colspan: int = 1

# Usage: Dict[str, DataCellMergeRule]
```

### 6. Weight Summary Config Schema

```python
# Current JSON structure
{
  "enabled": bool,
  "label_col_id": str,
  "value_col_id": str
}
```

**Proposed Class:**
```python
@dataclass
class WeightSummaryConfig:
    enabled: bool
    label_col_id: str
    value_col_id: str
```

---

## Configuration Consolidation Plan

### Proposed Configuration Classes

#### Class 1: `LayoutConfig`
**Purpose**: Consolidate layout-related settings

```python
@dataclass
class LayoutConfig:
    """Layout configuration for a sheet."""
    start_row: int
    
    # Optional blank rows
    add_blank_after_header: bool = False
    add_blank_before_footer: bool = False
    
    # Static content
    static_content_after_header: Dict[str, str] = field(default_factory=dict)
    static_content_before_footer: Dict[str, str] = field(default_factory=dict)
    
    # Merge rules
    merge_rules_after_header: Dict[str, int] = field(default_factory=dict)
    merge_rules_before_footer: Dict[str, int] = field(default_factory=dict)
    merge_rules_footer: Dict[str, int] = field(default_factory=dict)
    
    # Data cell merging
    data_cell_merging_rules: Optional[Dict[str, DataCellMergeRule]] = None
    
    # Weight summary
    weight_summary_config: Optional[WeightSummaryConfig] = None
    
    # Summary section (DAF mode)
    enable_summary: bool = False
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'LayoutConfig':
        """Create from JSON config dict."""
        return cls(
            start_row=data['start_row'],
            add_blank_after_header=data.get('add_blank_after_header', False),
            add_blank_before_footer=data.get('add_blank_before_footer', False),
            static_content_after_header=data.get('static_content_after_header', {}),
            static_content_before_footer=data.get('static_content_before_footer', {}),
            merge_rules_after_header=data.get('merge_rules_after_header', {}),
            merge_rules_before_footer=data.get('merge_rules_before_footer', {}),
            merge_rules_footer=data.get('merge_rules_footer', {}),
            data_cell_merging_rules=cls._parse_merge_rules(
                data.get('data_cell_merging_rule')
            ),
            weight_summary_config=cls._parse_weight_summary(
                data.get('weight_summary_config')
            ),
            enable_summary=data.get('summary', False)
        )
    
    @staticmethod
    def _parse_merge_rules(rules_dict):
        if not rules_dict:
            return None
        return {
            col_id: DataCellMergeRule(**rule_data)
            for col_id, rule_data in rules_dict.items()
        }
    
    @staticmethod
    def _parse_weight_summary(config_dict):
        if not config_dict:
            return None
        return WeightSummaryConfig(**config_dict)
```

**Impact**: Reduces 11 individual parameters → 1 object

#### Class 2: `SheetConfig`
**Purpose**: Complete sheet configuration

```python
@dataclass
class SheetConfig:
    """Complete configuration for a single sheet."""
    layout: LayoutConfig
    header_cells: List[HeaderCell]
    mappings: Dict[str, MappingRule]
    footer_config: FooterConfig
    styling: StylingConfigModel
    
    # Override data source (optional)
    data_source: Optional[str] = None
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'SheetConfig':
        """Create from JSON config dict."""
        return cls(
            layout=LayoutConfig.from_dict(data),
            header_cells=[HeaderCell(**cell) for cell in data['header_to_write']],
            mappings=cls._parse_mappings(data['mappings']),
            footer_config=FooterConfig.from_dict(data['footer_configurations']),
            styling=StylingConfigModel.from_dict(data['styling']),
            data_source=data.get('data_source')
        )
    
    @staticmethod
    def _parse_mappings(mappings_dict):
        # Parse different mapping types
        parsed = {}
        for key, value in mappings_dict.items():
            if value.get('type') == 'initial_static_rows':
                parsed[key] = InitialStaticMapping(**value)
            elif value.get('type') == 'formula':
                parsed[key] = FormulaMapping(**value)
            else:
                parsed[key] = StandardMapping(**value)
        return parsed
```

#### Class 3: `InvoiceConfig`
**Purpose**: Top-level configuration

```python
@dataclass
class InvoiceConfig:
    """Top-level invoice configuration."""
    sheets_to_process: List[str]
    sheet_data_map: Dict[str, str]
    sheets: Dict[str, SheetConfig]
    
    @classmethod
    def from_json_file(cls, path: str) -> 'InvoiceConfig':
        """Load from JSON file."""
        with open(path, 'r') as f:
            data = json.load(f)
        
        return cls(
            sheets_to_process=data['sheets_to_process'],
            sheet_data_map=data['sheet_data_map'],
            sheets={
                name: SheetConfig.from_dict(config)
                for name, config in data['data_mapping'].items()
            }
        )
```

---

## Usage Frequency Analysis

### By Field

| Field | Usage Count | Percentage | Priority |
|-------|-------------|------------|----------|
| `start_row` | 22/22 | 100% | ⭐⭐⭐ Critical |
| `header_to_write` | 22/22 | 100% | ⭐⭐⭐ Critical |
| `mappings` | 22/22 | 100% | ⭐⭐⭐ Critical |
| `footer_configurations` | 22/22 | 100% | ⭐⭐⭐ Critical |
| `styling` | 22/22 | 100% | ⭐⭐⭐ Critical |
| `static_content_before_footer` | 18/22 | 82% | ⭐⭐ High |
| `add_blank_before_footer` | 15/22 | 68% | ⭐⭐ High |
| `data_cell_merging_rule` | 14/22 | 64% | ⭐⭐ High |
| `merge_rules_before_footer` | 8/22 | 36% | ⭐ Medium |
| `weight_summary_config` | 6/22 | 27% | ⭐ Medium |
| `merge_rules_footer` | 3/22 | 14% | Low |
| `add_blank_after_header` | 2/22 | 9% | Low |
| `data_source` override | 2/22 | 9% | Low |
| `static_content_after_header` | 1/22 | 5% | Low |
| `merge_rules_after_header` | 1/22 | 5% | Low |
| `summary` | 1/22 | 5% | Low (DAF-specific) |

---

## Migration Strategy

### Phase 1: Create Config Models ✅ SAFE
**No breaking changes**

1. Create `invoice_generator/config/builder_models.py`
2. Define all dataclasses with `from_dict()` methods
3. Write comprehensive tests
4. Create config loader utility

**Files to create:**
- `invoice_generator/config/builder_models.py`
- `tests/config/test_builder_models.py`

### Phase 2: Update Builders (Backward Compatible)
**No breaking changes**

1. Update builders to accept **both** old and new formats:
   ```python
   def __init__(
       self,
       # NEW: preferred
       layout_config: Optional[LayoutConfig] = None,
       # OLD: backward compatible
       start_row: Optional[int] = None,
       add_blank_after_header: Optional[bool] = None,
       # ...
   ):
       # Convert old → new if needed
       if layout_config is None:
           layout_config = LayoutConfig(
               start_row=start_row or 1,
               add_blank_after_header=add_blank_after_header or False,
               # ...
           )
   ```

2. Update one builder at a time
3. Test thoroughly after each change

### Phase 3: Migrate Call Sites
**Gradual migration**

1. Update processors to use new classes
2. Keep old JSON configs working via `from_dict()`
3. No rush - both formats work

### Phase 4: Deprecate Old Parameters
**After all code migrated**

1. Remove old parameters
2. Clean up backward compatibility code
3. Update documentation

---

## Benefits of Refactoring

### Before: DataTableBuilder
```python
DataTableBuilder(
    worksheet,                        # 1
    sheet_name,                       # 2
    sheet_config,                     # 3 (dict)
    all_sheet_configs,                # 4 (dict)
    data_source,                      # 5
    data_source_type,                 # 6
    header_info,                      # 7 (dict)
    mapping_rules,                    # 8 (dict)
    sheet_styling_config,             # 9 (dict)
    add_blank_after_header,           # 10
    static_content_after_header,      # 11 (dict)
    add_blank_before_footer,          # 12
    static_content_before_footer,     # 13 (dict)
    merge_rules_after_header,         # 14 (dict)
    merge_rules_before_footer,        # 15 (dict)
    merge_rules_footer,               # 16 (dict)
    max_rows_to_fill,                 # 17
    grand_total_pallets,              # 18
    custom_flag,                      # 19
    data_cell_merging_rules,          # 20 (dict)
    DAF_mode,                         # 21
    all_tables_data,                  # 22 (dict)
    table_keys,                       # 23
    is_last_table,                    # 24
)
```

### After: DataTableBuilder
```python
DataTableBuilder(
    worksheet,                        # 1
    sheet_name,                       # 2
    config=sheet_config,              # 3 - SheetConfig object
    data_source=data_source,          # 4 - DataSourceConfig object
    header_info=header_info,          # 5 - HeaderInfo object
    build_context=context,            # 6 - BuildContext object (DAF, custom, etc.)
)
```

**Result**: 24 parameters → 6 parameters (75% reduction!)

### Type Safety Example

```python
# Before: No IDE help, typos possible
config = {
    "add_blank_after_header": True,
    "add_blankk_before_footer": True,  # Typo! Silent failure
}

# After: IDE autocomplete, type checking
config = LayoutConfig(
    add_blank_after_header=True,
    add_blank_before_footer=True,  # IDE catches typos
)
```

### Validation Example

```python
@dataclass
class LayoutConfig:
    start_row: int
    
    def __post_init__(self):
        if self.start_row < 1:
            raise ValueError(f"start_row must be >= 1, got {self.start_row}")
        
        # Validate static content keys are valid column references
        for key in self.static_content_before_footer.keys():
            if not key.isdigit():
                raise ValueError(f"Static content key must be numeric, got '{key}'")
```

---

## Next Steps

1. ✅ **Review this audit** - Ensure accuracy
2. ⬜ **Create `builder_models.py`** - Define all dataclasses
3. ⬜ **Create config loader** - JSON → typed classes
4. ⬜ **Write tests** - Ensure conversion works perfectly
5. ⬜ **Proof of concept** - Refactor one builder
6. ⬜ **Roll out gradually** - One builder at a time

---

## Questions for Discussion

1. **Snake_case vs camelCase**: Should config models use snake_case to match JSON, or camelCase to match existing `StylingConfigModel`?

2. **Pydantic vs Dataclasses**: Should we use Pydantic (like `StylingConfigModel`) or dataclasses for consistency?

3. **Validation Level**: How strict should validation be? Fail-fast or permissive?

4. **Migration Timeline**: How quickly to migrate? Gradual vs all-at-once?

5. **Config Format**: Keep JSON or migrate to Python modules?

---

## Conclusion

The current configuration system is fragile due to:
- Dictionary-based configs (no type safety)
- Scattered parameters across 24+ function arguments
- Silent failures from typos

**Proposed solution:**
- Typed configuration classes
- Consolidation into logical groups
- Backward-compatible migration
- Gradual rollout

**Expected benefits:**
- 75% reduction in parameters
- Type safety and IDE support
- Validation and clear errors
- Easier testing and maintenance

**Risk**: Low (backward compatible approach)  
**Effort**: Medium (one builder at a time)  
**Impact**: **High** (significantly improves maintainability)








