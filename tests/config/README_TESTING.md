# Testing the TableDataResolver/Adapter

## Overview

The `TableDataResolver` (candidate for renaming to `TableDataAdapter`) is responsible for transforming raw invoice data into builder-ready table rows. This is a critical adapter component that sits between data sources and builders.

## Running Tests

### Run All TableDataResolver Tests
```powershell
# Verbose output
python -m pytest tests/config/test_table_data_resolver.py -v

# With print statements (if any)
python -m pytest tests/config/test_table_data_resolver.py -v -s

# Run with coverage
python -m pytest tests/config/test_table_data_resolver.py --cov=invoice_generator.config.table_data_resolver --cov-report=html
```

### Run Specific Test Classes
```powershell
# Main test suite
python -m pytest tests/config/test_table_data_resolver.py::TestTableDataResolver -v

# Edge cases
python -m pytest tests/config/test_table_data_resolver.py::TestTableDataResolverEdgeCases -v

# Isolated unit tests
python -m pytest tests/config/test_table_data_resolver.py::TestTableDataResolverIsolated -v
```

### Run Individual Tests
```powershell
# Test with real config
python -m pytest tests/config/test_table_data_resolver.py::TestTableDataResolver::test_resolve_with_real_invoice_config -v

# Test mapping conversion
python -m pytest tests/config/test_table_data_resolver.py::TestTableDataResolverIsolated::test_bundled_to_legacy_mapping_conversion -v

# Test all data source types
python -m pytest tests/config/test_table_data_resolver.py::TestTableDataResolverIsolated::test_all_three_data_source_types -v
```

## Test Structure

### 1. TestTableDataResolver (Main Suite)
Tests core functionality with real JF config:
- Initialization with different data sources
- Factory method (`create_from_bundles`)
- Integration with BuilderConfigResolver
- Full resolution workflow
- Static info extraction

### 2. TestTableDataResolverEdgeCases
Tests error handling and edge cases:
- Empty data sources
- Missing header info
- Null values

### 3. TestTableDataResolverIsolated (New)
Pure unit tests without config dependencies:
- Bundled → Legacy mapping conversion
- All 3 data source types (aggregation, DAF_aggregation, processed_tables_multi)
- Caching behavior

## Test Data

The tests use sample invoice data with:
- **Aggregation data**: Key-value dict format from standard processing
- **Multi-table data**: Table '1' and '2' with column arrays (PO, Item, Description, etc.)

Example:
```python
invoice_data = {
    'standard_aggregation_results': {
        ('PO123', 'ITEM001', 5.25, 'LEATHER TYPE A'): {
            'sqft_sum': 100.50,
            'amount_sum': 527.625
        }
    },
    'processed_tables_data': {
        '1': {
            'po': ['PO123', 'PO123'],
            'item': ['ITEM001', 'ITEM002'],
            'pallet_count': [2, 3]
        }
    }
}
```

## What Gets Tested

### ✅ Covered
- Initialization with all data source types
- Factory method creation
- Integration with BuilderConfigResolver
- Mapping rule conversion (bundled → legacy)
- Multi-table data extraction
- Static info extraction
- DAF mode detection
- Caching of parsed rules

### 🔄 Could Add More Coverage
- Formula rule processing
- Dynamic description handling
- Pallet count calculations
- Error handling for malformed configs
- Performance with large datasets

## Adapter Pattern (Potential Rename)

The `TableDataResolver` is actually implementing the **Adapter Pattern**:
- **Purpose**: Adapts raw invoice data format → builder-compatible format
- **Input**: Heterogeneous data (aggregations, processed tables, DAF data)
- **Output**: Homogeneous table rows (List[Dict[int, Any]])

Consider renaming:
- `TableDataResolver` → `TableDataAdapter`
- `table_data_resolver.py` → `table_data_adapter.py`
- Update all imports and references

## Dependencies

Required:
- `invoice_generator.config.config_loader.BundledConfigLoader`
- `invoice_generator.config.builder_config_resolver.BuilderConfigResolver`
- `invoice_generator.data.data_preparer` (prepare_data_rows, parse_mapping_rules)
- Real config file: `invoice_generator/config_bundled/JF_config/JF_config.json`

## Continuous Integration

Add to CI pipeline:
```yaml
- name: Test TableDataResolver
  run: python -m pytest tests/config/test_table_data_resolver.py --cov --cov-report=xml
```
