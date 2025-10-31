# Builder Architecture: `layout_builder.py`

This document explains the structure and purpose of the `LayoutBuilder` class, which serves as the **Director** in the Builder pattern, orchestrating all specialized builders to construct the complete document layout.

## Overview

The `LayoutBuilder` is the central orchestrator that coordinates multiple specialized builders to construct a complete worksheet layout. It implements the **Director** role in the classic Builder pattern, managing the sequence of building operations and ensuring proper coordination between components.

- **Purpose**: To orchestrate the complete document layout construction by coordinating multiple specialized builders in the correct sequence.
- **Pattern**: **Director pattern** - Directs and coordinates multiple Builder objects to construct a complex product.
- **Key Responsibility**: Manages the entire layout building process from template state capture through header, data, and footer construction.

## Design Pattern: Director Pattern

The Director pattern separates the construction algorithm from the parts being constructed. The `LayoutBuilder` (Director) uses various specialized builders to construct the final product:

```
       LayoutBuilder (Director)
              |
              ├── TemplateStateBuilder (captures template state)
              ├── TextReplacementBuilder (replaces placeholders)
              ├── HeaderBuilder (builds headers)
              ├── DataTableBuilder (builds data tables)
              └── FooterBuilder (builds footers)
```

## `LayoutBuilder` Class

### Architecture: Three-Config Design Pattern

`LayoutBuilder` uses a **three-config architecture** for maximum extensibility and maintainability. All configurations are bundled into three organized dictionaries, allowing you to add new configurations without changing the constructor signature.

**The Three Config Bundles:**

1. **`style_config`**: All styling-related configurations
2. **`context_config`**: All runtime/contextual data
3. **`layout_config`**: All layout/building control configurations

### `__init__(...)` - The Constructor

The constructor initializes the director with three organized config bundles, providing clean separation of concerns.

- **Purpose**: To configure the layout builder with worksheets and three extensible config bundles (style, context, layout).
- **Parameters**:

    #### Core Worksheet Parameters
    - `workbook: Workbook`: The output workbook object (writable) where the final result will be saved.
    - `worksheet: Worksheet`: The output worksheet object (writable) where the layout will be constructed.
    - `template_worksheet: Worksheet`: The template worksheet object (read-only usage) from which template state is captured.

    #### `style_config: Dict[str, Any]` - Style Configuration Bundle
    Contains all styling-related configurations:
    - `styling_config: StylingConfigModel`: The styling configuration model for the sheet.
    - *(Future style configs can be added here without changing the signature)*

    #### `context_config: Dict[str, Any]` - Context Configuration Bundle
    Contains all runtime/contextual data:
    - `sheet_name: str`: The name of the sheet being processed (e.g., `"Invoice"`, `"Packing list"`).
    - `invoice_data: Dict[str, Any]`: The complete invoice data dictionary containing all data sources:
        - `standard_aggregation_results`
        - `DAF_aggregation` / `final_DAF_compounded_result`
        - `custom_aggregation_results`
        - `processed_tables_data`
    - `all_sheet_configs: Dict[str, Any]`: The full configuration dictionary for all sheets (allows cross-sheet references).
    - `args: Optional[Any]`: Command-line arguments object with flags:
        - `args.DAF`: Delivery At Frontier mode flag.
        - `args.custom`: Custom processing mode flag.
    - `final_grand_total_pallets: int`: The global total pallet count across all sheets (default: `0`).
    - *(Future context data can be added here without changing the signature)*

    #### `layout_config: Dict[str, Any]` - Layout Configuration Bundle
    Contains all layout/building control configurations:
    - `sheet_config: Dict[str, Any]`: The complete configuration for this specific sheet, including:
        - `start_row`: Row where data starts in the template.
        - `header_to_write`: Header layout configuration.
        - `mappings`: Data column mapping rules.
        - `footer_configurations`: Footer configuration.
        - `add_blank_after_header`, `add_blank_before_footer`: Layout flags.
        - `static_content_after_header`, `static_content_before_footer`: Static content dictionaries.
        - `merge_rules_*`: Various merge rule configurations.
        - `data_source`: Indicator for which data source to use.
    - `enable_text_replacement: bool`: Flag to enable text replacement processing (default: `False`).
    - **Skip Flags for Customization**:
        - `skip_template_header_restoration: bool`: Skip restoring template header (default: `False`).
        - `skip_header_builder: bool`: Skip running HeaderBuilder (default: `False`).
        - `skip_data_table_builder: bool`: Skip running DataTableBuilder (default: `False`).
        - `skip_footer_builder: bool`: Skip running FooterBuilder (default: `False`).
        - `skip_template_footer_restoration: bool`: Skip restoring template footer (default: `False`).
    - *(Future layout flags can be added here without changing the signature)*

- **Instance Variables Initialized**:
    - **Input Context**: All constructor parameters stored as instance variables.
    - **Config Bundles**: `style_config`, `context_config`, `layout_config` stored as-is for extensibility.
    - **Extracted Conveniences**: Commonly used values extracted from bundles (e.g., `self.sheet_name`, `self.invoice_data`).
    - **Output State Variables**:
        - `header_info`: Stores header metadata after HeaderBuilder runs.
        - `next_row_after_footer`: Row index after footer section.
        - `data_start_row`: Starting row of data section (exposed for multi-table sum calculations).
        - `data_end_row`: Ending row of data section (exposed for multi-table sum calculations).
        - `dynamic_desc_used`: Flag indicating if dynamic descriptions were used (for summary add-on).
        - `template_state_builder`: Reference to TemplateStateBuilder instance.

### Usage Example - Three-Config Pattern

```python
# Prepare the three config bundles
style_config = {
    'styling_config': styling_model,  # StylingConfigModel instance
    # Add any future style configs here
}

context_config = {
    'sheet_name': 'Invoice',
    'invoice_data': {
        'standard_aggregation_results': [...],
        'DAF_aggregation': {...},
        'processed_tables_data': {...}
    },
    'all_sheet_configs': sheet_configs,
    'args': args,
    'final_grand_total_pallets': 150,
    # Add any future context data here
}

layout_config = {
    'sheet_config': {
        'start_row': 10,
        'header_to_write': ['SKU', 'Description', 'Qty'],
        'mappings': {...},
        'footer_configurations': {...}
    },
    'enable_text_replacement': True,
    'skip_header_builder': False,
    'skip_data_table_builder': False,
    'skip_footer_builder': False,
    'skip_template_header_restoration': False,
    'skip_template_footer_restoration': False,
    # Add any future layout flags here
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

**Benefits of This Pattern:**
- ✅ **Extensible**: Add new configs without changing constructor signature
- ✅ **Organized**: Clear separation of style, context, and layout concerns
- ✅ **Maintainable**: Easy to understand what goes where
- ✅ **Backward Compatible**: Extracts commonly used values for existing code

### Migration Guide - Converting to Three-Config Pattern

If you have existing code using the old parameter-based approach, here's how to convert it:

**OLD Pattern (Before):**
```python
layout_builder = LayoutBuilder(
    workbook=output_wb,
    worksheet=output_ws,
    template_worksheet=template_ws,
    sheet_name='Invoice',
    sheet_config=sheet_config,
    all_sheet_configs=all_configs,
    invoice_data=invoice_data,
    styling_config=styling_config,
    args=args,
    final_grand_total_pallets=150,
    enable_text_replacement=True,
    skip_header_builder=False,
    skip_footer_builder=False
)
```

**NEW Pattern (After):**
```python
# Bundle into three organized configs
style_config = {
    'styling_config': styling_config
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

layout_builder = LayoutBuilder(
    workbook=output_wb,
    worksheet=output_ws,
    template_worksheet=template_ws,
    style_config=style_config,
    context_config=context_config,
    layout_config=layout_config
)
```

**Checklist for Migration:**
1. ✅ Move `styling_config` into `style_config` dict
2. ✅ Move `sheet_name`, `invoice_data`, `all_sheet_configs`, `args`, `final_grand_total_pallets` into `context_config` dict
3. ✅ Move `sheet_config`, `enable_text_replacement`, and all `skip_*` flags into `layout_config` dict
4. ✅ Keep `workbook`, `worksheet`, `template_worksheet` as direct parameters (unchanged)

### `build(self) -> bool` - The Main Orchestration Method

This is the core method that orchestrates all builders in the correct sequence to construct the complete layout.

- **Purpose**: To execute the complete layout building process by coordinating all specialized builders in proper sequence.
- **Orchestration Sequence**:

#### **Phase 1: Text Replacement (Optional Pre-processing)**
   - **Condition**: Only if `enable_text_replacement` is `True`.
   - **Action**: Creates and runs `TextReplacementBuilder` to replace placeholders.
   - **Mode-Dependent**: 
     - If `args.DAF`: Runs both placeholder and DAF-specific replacements.
     - Otherwise: Runs only placeholder replacements.

#### **Phase 2: Template Boundary Calculation**
   - **Purpose**: Calculate header and footer boundaries in the template worksheet.
   - **Logic**:
     - `template_header_start_row`: Always `1` (template headers start at row 1).
     - `template_header_end_row`: `original_start_row - 1` (header ends before data starts).
     - `template_footer_start_row`: `original_start_row + 2` (footer in template after header + minimal data).
   - **Important**: Uses the **original** `start_row` from the template config, not dynamic values that change in multi-table scenarios.

#### **Phase 3: Template State Capture**
   - **Action**: Creates `TemplateStateBuilder` and captures header/footer state from `template_worksheet`.
   - **Purpose**: Preserves the template's visual structure (formatting, merges, styles) for restoration.

#### **Phase 3b: Template Header Restoration**
   - **Condition**: Only if `skip_template_header_restoration` is `False`.
   - **Action**: Restores captured header to the output worksheet.
   - **Purpose**: Copies template header structure to output before HeaderBuilder writes content.

#### **Phase 4: Header Building**
   - **Condition**: Only if `skip_header_builder` is `False`.
   - **Action**: Creates and runs `HeaderBuilder` to build header rows.
   - **Output**: Stores `header_info` dictionary for use by downstream builders.
   - **Styling Conversion**: Converts styling config dict to `StylingConfigModel` if needed.
   - **Fallback**: If skipped, provides dummy `header_info` for downstream builders.
   - **Error Handling**: Returns `False` if header building fails.

#### **Phase 5: Data Table Building**
   - **Condition**: Only if `skip_data_table_builder` is `False`.
   - **Steps**:
     1. **Extract Configuration**: Gets mapping rules, merge rules, flags, and data source indicator from `sheet_config`.
     2. **Data Source Resolution**: Determines which data to use:
        - **Custom Mode**: Uses `custom_aggregation_results` if `args.custom` is `True`.
        - **DAF Mode**: Uses `final_DAF_compounded_result` for Invoice/Contract sheets if `args.DAF` is `True`.
        - **Standard Mode**: Uses `standard_aggregation_results` for normal aggregation.
        - **Multi-Table Mode**: Uses data from `processed_tables_data[data_source_indicator]`.
     3. **Builder Execution**: Creates and runs `DataTableBuilder`.
     4. **Output Storage**: Stores:
        - `footer_row_position`: Where footer should be placed.
        - `data_start_row`, `data_end_row`: Data range for sum calculations.
        - `local_chunk_pallets`: Pallet count for this data chunk.
        - `dynamic_desc_used`: Whether dynamic descriptions were used.
   - **Fallback**: If skipped, provides dummy values for downstream builders.
   - **Error Handling**: Returns `False` if data table building fails.

#### **Phase 6: Footer Building**
   - **Condition**: Only if `skip_footer_builder` is `False`.
   - **Steps**:
     1. **Pallet Count Resolution**:
        - For multi-table mode (`processed_tables`): Uses `local_chunk_pallets`.
        - Otherwise: Uses `final_grand_total_pallets`.
     2. **Sum Range Preparation**: Creates list of tuples `[(data_start_row, data_end_row)]` for SUM formulas.
     3. **Builder Execution**: Creates and runs `FooterBuilder`.
     4. **Output Storage**: Stores `next_row_after_footer` for template footer placement.
     5. **Row Height Application**: Applies footer row heights to all footer rows (including add-ons).
   - **Fallback**: If skipped, sets `next_row_after_footer` to footer row position.

#### **Phase 7: Template Footer Restoration**
   - **Condition**: Only if `skip_template_footer_restoration` is `False`.
   - **Action**: Restores captured template footer to output worksheet at `next_row_after_footer`.
   - **Purpose**: Places template's static footer content (e.g., "Manufacturer:", signatures) below the dynamic data footer.

- **Return Value**: `bool`
    - `True`: Layout built successfully.
    - `False`: Critical error occurred during header or data building.

### `_apply_footer_row_height(self, footer_row: int, styling_config)` - Helper Method

Applies the configured row height to a single footer row.

- **Purpose**: To set footer row height based on styling configuration.
- **Parameters**:
    - `footer_row: int`: The Excel row index to apply height to.
    - `styling_config`: The styling configuration model.
- **Height Resolution Logic**:
    1. Check if `footer_matches_header_height` is `True` (default).
    2. If `True`, use header height; otherwise use explicit footer height.
    3. Apply the determined height to row dimensions.
- **Graceful Handling**: Silently skips if configuration is missing or values are invalid.

## Builder Coordination Flow

```
LayoutBuilder.build()
        |
        ├─ [Optional] TextReplacementBuilder
        |       └─ Replace placeholders in workbook
        |
        ├─ Calculate Template Boundaries
        |       └─ Determine header/footer positions in template
        |
        ├─ TemplateStateBuilder (Capture)
        |       └─ Capture template state from template_worksheet
        |
        ├─ TemplateStateBuilder (Restore Header)
        |       └─ Restore header to output worksheet
        |
        ├─ HeaderBuilder
        |       ├─ Build header rows
        |       └─ Return header_info
        |
        ├─ DataTableBuilder
        |       ├─ Resolve data source
        |       ├─ Build data table
        |       └─ Return footer_position, data_range, pallets
        |
        ├─ FooterBuilder
        |       ├─ Build footer row(s)
        |       ├─ Apply row heights
        |       └─ Return next_row_after_footer
        |
        └─ TemplateStateBuilder (Restore Footer)
                └─ Restore template footer to output worksheet
```

## Key Design Decisions

### 1. **Separation of Template and Output Worksheets**
The builder reads from `template_worksheet` (read-only) and writes to `worksheet` (writable), completely avoiding merge conflicts that could occur from modifying the same worksheet.

### 2. **Skip Flags for Flexibility**
Five skip flags (`skip_template_header_restoration`, `skip_header_builder`, `skip_data_table_builder`, `skip_footer_builder`, `skip_template_footer_restoration`) allow custom processors to selectively disable specific build phases, enabling specialized workflows without code duplication.

### 3. **Data Source Auto-Resolution**
The builder automatically determines which data source to use based on mode flags (`args.DAF`, `args.custom`) and sheet type, simplifying processor implementations.

### 4. **Exposed State Variables**
Critical state variables (`data_start_row`, `data_end_row`, `dynamic_desc_used`) are exposed as instance variables, allowing multi-table processors to access them for sum range calculations and summary add-on logic.

### 5. **Template Boundary Calculation**
Template boundaries are always calculated from the **original** `start_row` in the config, not dynamic values, ensuring consistent template state capture regardless of where data is written in multi-table scenarios.

### 6. **Centralized Row Height Management**
The director applies footer row heights for all footer rows (including add-ons), ensuring consistent height application even when FooterBuilder creates multiple rows.

## Data Source Resolution Logic

The builder uses the following priority order to resolve data sources:

```
1. Check if args.custom and data_source == 'aggregation'
   └─ Use invoice_data['custom_aggregation_results']
   
2. Check if args.DAF and sheet is Invoice/Contract
   └─ Use invoice_data['final_DAF_compounded_result']
   
3. Check data_source_indicator value
   ├─ 'DAF_aggregation' → Use invoice_data['final_DAF_compounded_result']
   ├─ 'aggregation' → Use invoice_data['standard_aggregation_results']
   └─ Other key → Use invoice_data['processed_tables_data'][key]
   
4. If no data found, log warning and skip fill
```

## Skip Flag Use Cases

The skip flags enable specialized processing scenarios:

### Use Case 1: Custom Header Pre-Processing
```python
# A processor that builds custom headers before using LayoutBuilder
layout_config = {
    'sheet_config': config,
    'skip_header_builder': True,  # Don't rebuild header
}

layout_builder = LayoutBuilder(
    workbook=workbook,
    worksheet=worksheet,
    template_worksheet=template,
    style_config={'styling_config': styling},
    context_config={'sheet_name': 'Invoice', 'invoice_data': data, ...},
    layout_config=layout_config
)
```

### Use Case 2: Multi-Table Processing
```python
# Multi-table processor that builds data tables separately
layout_config = {
    'sheet_config': config,
    'skip_data_table_builder': True,  # Build tables separately
    'skip_footer_builder': True       # Build grand total footer separately
}

layout_builder = LayoutBuilder(
    workbook=workbook,
    worksheet=worksheet,
    template_worksheet=template,
    style_config={'styling_config': styling},
    context_config={'sheet_name': 'Packing list', 'invoice_data': data, ...},
    layout_config=layout_config
)
```

### Use Case 3: Template-Free Generation
```python
# Generate without template restoration
layout_config = {
    'sheet_config': config,
    'skip_template_header_restoration': True,
    'skip_template_footer_restoration': True
}

layout_builder = LayoutBuilder(
    workbook=workbook,
    worksheet=worksheet,
    template_worksheet=template,
    style_config={'styling_config': styling},
    context_config={'sheet_name': 'Invoice', 'invoice_data': data, ...},
    layout_config=layout_config
)
```

## Dependencies

- **Specialized Builders**:
    - `HeaderBuilder` - Builds table headers.
    - `DataTableBuilder` - Builds data table sections.
    - `FooterBuilder` - Builds footer sections.
    - `TextReplacementBuilder` - Replaces text placeholders.
    - `TemplateStateBuilder` - Captures and restores template state.
- **Style Configuration**: `StylingConfigModel` - Type-safe styling model.
- **openpyxl**: Core Excel manipulation library.

## Usage Example

### Standard Single-Table Sheet

```python
from invoice_generator.builders.layout_builder import LayoutBuilder

# Prepare three config bundles
style_config = {
    'styling_config': styling_config
}

context_config = {
    'sheet_name': 'Invoice',
    'invoice_data': data,
    'all_sheet_configs': all_configs,
    'args': args,  # Contains args.DAF, args.custom flags
    'final_grand_total_pallets': 250
}

layout_config = {
    'sheet_config': invoice_config,
    'enable_text_replacement': True
}

# Standard layout build for Invoice sheet
layout_builder = LayoutBuilder(
    workbook=output_workbook,
    worksheet=output_worksheet,
    template_worksheet=template_worksheet,
    style_config=style_config,
    context_config=context_config,
    layout_config=layout_config
)

success = layout_builder.build()

if success:
    print(f"Layout built successfully")
    print(f"Data range: {layout_builder.data_start_row} to {layout_builder.data_end_row}")
    print(f"Next row after footer: {layout_builder.next_row_after_footer}")
else:
    print("Layout building failed")
```

### Custom Processing with Skip Flags

```python
# Custom processor that handles header building separately
style_config = {
    'styling_config': styling_config
}

context_config = {
    'sheet_name': 'Custom Report',
    'invoice_data': data,
    'all_sheet_configs': all_configs,
    'args': args
}

layout_config = {
    'sheet_config': custom_config,
    'skip_header_builder': True,  # Header already built by custom logic
    'skip_template_header_restoration': True  # Don't restore template header
}

layout_builder = LayoutBuilder(
    workbook=output_workbook,
    worksheet=output_worksheet,
    template_worksheet=template_worksheet,
    style_config=style_config,
    context_config=context_config,
    layout_config=layout_config
)

success = layout_builder.build()
```

### DAF Mode Processing

```python
# DAF mode automatically uses DAF data source for Invoice/Contract
style_config = {
    'styling_config': styling_config
}

context_config = {
    'sheet_name': 'Invoice',
    'invoice_data': data,  # Must contain 'final_DAF_compounded_result'
    'all_sheet_configs': all_configs,
    'args': args,  # args.DAF = True
    'final_grand_total_pallets': 300
}

layout_config = {
    'sheet_config': invoice_config,
    'enable_text_replacement': True  # Runs DAF text replacements
}

layout_builder = LayoutBuilder(
    workbook=output_workbook,
    worksheet=output_worksheet,
    template_worksheet=template_worksheet,
    style_config=style_config,
    context_config=context_config,
    layout_config=layout_config
)

success = layout_builder.build()
# Uses DAF_aggregation data source automatically
```

## Integration with Processors

The `LayoutBuilder` is typically called by sheet processors (e.g., `SingleTableProcessor`, `MultiTableProcessor`):

```python
class SingleTableProcessor(SheetProcessor):
    def process(self) -> bool:
        # ... initialization ...
        
        # Prepare config bundles
        style_config = {
            'styling_config': styling_config
        }
        
        context_config = {
            'sheet_name': self.sheet_name,
            'invoice_data': self.invoice_data,
            'all_sheet_configs': self.data_mapping_config,
            'args': self.cli_args,
            'final_grand_total_pallets': self.final_grand_total_pallets
        }
        
        layout_config = {
            'sheet_config': self.sheet_config
        }
        
        # Create and run LayoutBuilder
        layout_builder = LayoutBuilder(
            workbook=self.workbook,
            worksheet=self.worksheet,
            template_worksheet=template_ws,
            style_config=style_config,
            context_config=context_config,
            layout_config=layout_config
        )
        
        return layout_builder.build()
```

## Notes

- The `LayoutBuilder` assumes both `workbook` and `worksheet` (output) and `template_worksheet` are provided and valid.
- The director pattern centralizes the building sequence, making it easy to modify the order or add new build phases.
- Skip flags do not validate dependencies; it's the caller's responsibility to ensure skipped phases don't break downstream phases.
- Template boundary calculations are critical for multi-table scenarios and use the **original** `start_row` from config, not dynamic values.
- The `header_info` dictionary is essential for all downstream builders; if `skip_header_builder` is `True`, a dummy header_info is provided.
- Text replacement is optional and typically disabled for performance; it's only needed when placeholders exist in the template.
- The builder exposes state variables (`data_start_row`, `data_end_row`, etc.) for multi-table processors that need to calculate sum ranges across multiple tables.
- Row height application for footers occurs in the director (LayoutBuilder), not in FooterBuilder, to ensure consistent height application across all footer rows including add-ons.


