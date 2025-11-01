# TableDataResolver Implementation Summary

## Overview
Created a dedicated `TableDataResolver` class that centralizes data preparation logic, eliminating the need for builders to directly call data preparation functions.

## Architecture

### Resolver Chain
```
BundledConfigLoader
    ↓
BuilderConfigResolver
    ↓
TableDataResolver  ← NEW!
    ↓
DataTableBuilder
```

## What TableDataResolver Does

### Primary Responsibility
Transforms raw invoice data into table-ready row dictionaries that builders can directly use.

### Input
- **data_source**: Raw invoice data (aggregations, processed tables, etc.)
- **data_source_type**: Type of data source ('aggregation', 'DAF_aggregation', 'processed_tables_multi', etc.)
- **mapping_rules**: Configuration rules for mapping data to columns
- **header_info**: Column mappings and IDs
- **DAF_mode**: Whether DAF mode is enabled

### Output
Dictionary containing:
```python
{
    'data_rows': List[Dict[int, Any]],      # Prepared row data (col_index → value)
    'pallet_counts': List[int],             # Pallet counts per row
    'dynamic_desc_used': bool,              # Whether dynamic descriptions were used
    'num_data_rows': int,                   # Number of data rows from source
    'static_info': {                        # Static configuration
        'col1_index': int,
        'num_static_labels': int,
        'initial_static_col1_values': List,
        'static_column_header_name': str,
        'apply_special_border_rule': bool
    },
    'formula_rules': Dict                   # Formula rules by column
}
```

## Implementation Details

### File Location
`invoice_generator/config/table_data_resolver.py`

### Key Methods

#### `resolve() -> Dict[str, Any]`
Main entry point that:
1. Parses mapping rules
2. Extracts table-specific data (for multi-table scenarios)
3. Calls `prepare_data_rows()` from `data_preparer.py`
4. Returns comprehensive result dictionary

#### `create_from_bundles(data_config, context_config) -> TableDataResolver`
Factory method that creates a resolver from configuration bundles. This is the **recommended** way to instantiate the resolver.

```python
# Usage
table_resolver = TableDataResolver.create_from_bundles(
    data_config=builder_resolver.get_data_bundle(table_key='1'),
    context_config=builder_resolver.get_context_bundle()
)
result = table_resolver.resolve()
```

### Integration with BuilderConfigResolver

Added `get_table_data_resolver()` method to `BuilderConfigResolver`:

```python
def get_table_data_resolver(self, table_key: Optional[str] = None) -> TableDataResolver:
    """
    Get a TableDataResolver instance for preparing table data.
    
    Args:
        table_key: Optional table key for multi-table scenarios
        
    Returns:
        Configured TableDataResolver instance
    """
    data_config = self.get_data_bundle(table_key=table_key)
    context_config = self.get_context_bundle()
    
    return TableDataResolver.create_from_bundles(
        data_config=data_config,
        context_config=context_config
    )
```

## Data Source Type Support

### Supported Types
- **aggregation**: Standard aggregation data (tuple keys → dict values)
- **DAF_aggregation**: DAF mode aggregation
- **custom_aggregation**: Custom aggregation format
- **processed_tables**: Single table with column arrays
- **processed_tables_multi**: Multiple tables (requires `table_key`)

### Multi-Table Processing
For `processed_tables_multi`, the resolver works with `BuilderConfigResolver` to extract specific table data:

```python
# BuilderConfigResolver extracts table '1' data
data_config = builder_resolver.get_data_bundle(table_key='1')

# TableDataResolver processes the extracted table
table_resolver = builder_resolver.get_table_data_resolver(table_key='1')
result = table_resolver.resolve()
```

## Testing

### Test Coverage
Created `tests/config/test_table_data_resolver.py` with **14 comprehensive tests**:

#### Initialization Tests (2)
- ✅ `test_initialization_with_standard_aggregation`
- ✅ `test_initialization_with_table_key`

#### Factory Method Tests (2)
- ✅ `test_create_from_bundles`
- ✅ `test_create_from_bundles_with_daf_mode`

#### Integration Tests (2)
- ✅ `test_get_table_data_resolver_from_builder_resolver`
- ✅ `test_get_table_data_resolver_with_table_key`

#### Helper Method Tests (2)
- ✅ `test_idx_to_header_map_generation`
- ✅ `test_get_desc_col_idx_finds_description_column`
- ✅ `test_extract_table_data_for_multi_table`

#### Resolution Tests (3)
- ✅ `test_resolve_with_real_invoice_config`
- ✅ `test_resolve_with_packing_list_config`
- ✅ `test_static_info_extraction`

#### Edge Case Tests (2)
- ✅ `test_resolver_with_empty_data_source`
- ✅ `test_resolver_with_missing_header_info`

### Total Test Count
**116 tests** across all resolver components:
- 21 tests: `BundleAccessor`
- 44 tests: `BuilderConfigResolver`
- 30 tests: JF config integration
- 7 tests: Layout convenience methods
- 14 tests: `TableDataResolver` ← NEW!

All tests passing! ✅

## Benefits

### 1. Separation of Concerns
- **Before**: DataTableBuilder directly called `prepare_data_rows()`
- **After**: TableDataResolver handles data preparation; builders just consume prepared data

### 2. Single Responsibility
- **TableDataResolver**: Transform raw data → table-ready rows
- **BuilderConfigResolver**: Extract/organize config bundles
- **DataTableBuilder**: Build Excel tables with styled cells

### 3. Testability
- Isolated testing of data preparation logic
- Mock-friendly factory method
- Clear input/output contracts

### 4. Consistency
- Standardized data preparation across all tables
- Centralized handling of different data source types
- Uniform error handling

### 5. Maintainability
- Single place to update data preparation logic
- Clear dependency chain
- Self-documenting code with type hints

## Data Preparation Fix

### Issue Found
`data_preparer.py` only handled `'processed_tables'` data source type, not `'processed_tables_multi'`.

### Fix Applied
```python
# Before
elif data_source_type == 'processed_tables':
    ...

# After
elif data_source_type in ['processed_tables', 'processed_tables_multi']:
    ...
```

This fix ensures Packing list (which uses `processed_tables_multi`) works correctly.

## Usage Examples

### Single Table (Invoice Sheet)
```python
from openpyxl import Workbook
from invoice_generator.config.config_loader import BundledConfigLoader
from invoice_generator.config.builder_config_resolver import BuilderConfigResolver

workbook = Workbook()
worksheet = workbook.active

config_loader = BundledConfigLoader('config/JF_config.json')
builder_resolver = BuilderConfigResolver(
    config_loader=config_loader,
    sheet_name='Invoice',
    worksheet=worksheet,
    args=args,
    invoice_data=invoice_data
)

# Get table data resolver
table_resolver = builder_resolver.get_table_data_resolver()

# Resolve data
result = table_resolver.resolve()

# Use prepared data
data_rows = result['data_rows']
pallet_counts = result['pallet_counts']
```

### Multi-Table (Packing List Sheet)
```python
# For table 1
table_resolver = builder_resolver.get_table_data_resolver(table_key='1')
result_table_1 = table_resolver.resolve()

# For table 2
table_resolver = builder_resolver.get_table_data_resolver(table_key='2')
result_table_2 = table_resolver.resolve()
```

## Next Steps

### 1. Update DataTableBuilder
Refactor `DataTableBuilder` to use `TableDataResolver` instead of directly calling `prepare_data_rows()`:

```python
# Before (legacy)
data_rows, pallet_counts, dynamic_desc_used, num_data_rows = prepare_data_rows(...)

# After (modern)
table_resolver = builder_resolver.get_table_data_resolver()
result = table_resolver.resolve()
data_rows = result['data_rows']
pallet_counts = result['pallet_counts']
```

### 2. Update multi_table_processor
Migrate multi-table processor to use the resolver pattern for each table.

### 3. Create Integration Examples
Add comprehensive examples in `docs/examples/` showing:
- Full resolver chain usage
- Multi-table processing
- DAF mode handling
- Error handling patterns

### 4. Deprecation Warning
Add logging to identify legacy code paths still calling `prepare_data_rows()` directly (similar to what was done in `layout_builder.py`).

## Related Documentation
- `docs/examples/layout_builder_resolver_pattern.py` - Usage patterns
- `docs/builders/data_table_builder_documentation.md` - Builder documentation
- `invoice_generator/data/data_preparer.py` - Underlying data preparation logic

## Summary
The `TableDataResolver` represents a major architectural improvement, establishing clear separation between configuration resolution and data preparation. This follows the single responsibility principle and makes the codebase more maintainable, testable, and easier to understand.

**Status**: ✅ Implemented and tested (116 tests passing)
