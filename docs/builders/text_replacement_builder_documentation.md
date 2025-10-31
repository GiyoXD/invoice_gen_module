# Builder Architecture: `text_replacement_builder.py`

This document explains the structure and purpose of the `TextReplacementBuilder` class, which is responsible for finding and replacing text placeholders and location-specific terms throughout the workbook.

## Overview

The `TextReplacementBuilder` is a specialized builder that performs two types of text replacement operations: data-driven placeholder replacement (e.g., invoice numbers, dates, customer info) and hardcoded DAF-specific location replacements. It operates at the workbook level, scanning and modifying text across all worksheets.

- **Purpose**: To replace text placeholders with actual data and perform location-specific text substitutions for DAF (Delivery At Frontier) mode.
- **Pattern**: Builder pattern - constructs a workbook with replaced text content.
- **Key Responsibility**: Executes configured text replacement rules across the entire workbook using a centralized replacement engine.

## Key Concepts

### Two Replacement Types

1. **Placeholder Replacement**: Data-driven replacements using values from `invoice_data`
   - Example: `"JFINV"` → `"JF25046"`
   - Example: `"[[CUSTOMER_NAME]]"` → `"ABC Trading Co., Ltd."`

2. **DAF-Specific Replacement**: Hardcoded location/term substitutions
   - Example: `"BINH PHUOC"` → `"BAVET"`
   - Example: `"FCA"` → `"DAF"`

### Workbook-Level Operation

Unlike other builders that work on individual worksheets, this builder operates on the **entire workbook**, scanning all sheets for matching text.

## `TextReplacementBuilder` Class

### `__init__(...)` - The Constructor

The constructor initializes the builder with the workbook and data source.

- **Purpose**: To configure the builder with the workbook to modify and the data to use for replacements.
- **Parameters**:
    - `workbook: openpyxl.Workbook`: The workbook object containing all worksheets to process. This is modified in-place.
    - `invoice_data: Dict[str, Any]`: The complete invoice data dictionary containing:
        - `processed_tables_data`: Multi-table data with invoice numbers, dates, references
        - `customer_info`: Customer name and address
        - Other data sources as needed

### `build(self)` - The Main Build Method

The primary method that executes all text replacement tasks in sequence.

- **Purpose**: To run all configured text replacement operations on the workbook.
- **Process**:
    1. Call `_replace_placeholders()` - Replace data-driven placeholders
    2. Call `_run_daf_specific_replacement()` - Replace DAF-specific terms
- **Return Value**: `None` (modifies workbook in-place)
- **Use Case**: Called by `LayoutBuilder` when `enable_text_replacement=True`

### `_replace_placeholders(self)` - Data-Driven Placeholder Replacement

Replaces placeholders with actual data from `invoice_data`.

- **Purpose**: To populate template placeholders with real invoice data (invoice numbers, dates, customer information).
- **Search Area**: Limited to cells `A1:N14` (first 14 rows and 14 columns) for performance.
- **Replacement Rules**: Defines 5 standard placeholder rules:

| Placeholder | Data Path | Type | Description |
|-------------|-----------|------|-------------|
| `"JFINV"` | `processed_tables_data` → `1` → `inv_no` → `[0]` | String | Invoice number |
| `"JFTIME"` | `processed_tables_data` → `1` → `inv_date` → `[0]` | Date | Invoice date (formatted) |
| `"JFREF"` | `processed_tables_data` → `1` → `inv_ref` → `[0]` | String | Invoice reference |
| `"[[CUSTOMER_NAME]]"` | `customer_info` → `name` | String | Customer company name |
| `"[[CUSTOMER_ADDRESS]]"` | `customer_info` → `address` | String | Customer address |

- **Match Mode**: All rules use `"exact"` match mode (case-sensitive, whole cell match).
- **Date Handling**: `JFTIME` rule has `is_date: True`, enabling date formatting.
- **Delegation**: Calls `find_and_replace()` utility function from `utils.text`.

#### Rule Structure

Each rule is a dictionary with the following keys:
- `find` (`str`): The text to search for
- `data_path` (`List[str]`): Path to navigate through `invoice_data` to find replacement value
- `match_mode` (`str`): `"exact"` (whole cell) or `"substring"` (within cell)
- `is_date` (`bool`, optional): If `True`, formats value as date

### `_run_daf_specific_replacement(self)` - DAF Location and Term Replacement

Replaces location names and shipping terms for DAF (Delivery At Frontier) mode.

- **Purpose**: To convert location names and shipping terms from various formats to standardized DAF format.
- **Search Area**: Expanded to `200 rows × 16 columns` to cover header and data sections.
- **Replacement Rules**: Defines 15 hardcoded replacement rules:

#### Location Replacements (Exact Match)

| Find | Replace | Description |
|------|---------|-------------|
| `"BINH PHUOC"` | `"BAVET"` | Vietnam province → Cambodia border town |
| `"BAVET, SVAY RIENG"` | `"BAVET"` | Remove province suffix |
| `"BAVET,SVAY RIENG"` | `"BAVET"` | Remove province suffix (no space) |
| `"BAVET, SVAYRIENG"` | `"BAVET"` | Remove province suffix (variant spelling) |
| `"BINH DUONG"` | `"BAVET"` | Vietnam province → Cambodia border town |
| `"SVAY RIENG"` | `"BAVET"` | Province name → border town |
| `"PORT KLANG"` | `"BAVET"` | Malaysia port → Cambodia border town |
| `"HCM"` | `"BAVET"` | Ho Chi Minh City → Cambodia border town |

#### Shipping Term Replacements

| Find | Replace | Match Mode | Description |
|------|---------|------------|-------------|
| `"FCA  BAVET,SVAYRIENG"` | `"DAF BAVET"` | Exact | Normalize shipping term with location |
| `"FCA: BAVET,SVAYRIENG"` | `"DAF: BAVET"` | Exact | Normalize with colon format |
| `"DAF  BAVET,SVAYRIENG"` | `"DAF BAVET"` | Exact | Normalize existing DAF term |
| `"DAF: BAVET,SVAYRIENG"` | `"DAF: BAVET"` | Exact | Normalize with colon format |
| `"DAP"` | `"DAF"` | Substring | Delivered At Place → Delivered At Frontier |
| `"FCA"` | `"DAF"` | Substring | Free Carrier → Delivered At Frontier |
| `"CIF"` | `"DAF"` | Substring | Cost Insurance Freight → Delivered At Frontier |

- **Match Modes**:
    - **Exact**: Replaces entire cell content only if it matches exactly
    - **Substring**: Replaces matching text within cell content (can replace multiple occurrences)

- **Delegation**: Calls `find_and_replace()` utility function from `utils.text`.

## Replacement Engine Integration

The builder doesn't implement the replacement logic itself; it delegates to the `find_and_replace()` function from `utils.text`:

```python
from ..utils.text import find_and_replace

find_and_replace(
    workbook=self.workbook,
    rules=replacement_rules,
    limit_rows=14,
    limit_cols=14,
    invoice_data=self.invoice_data  # Optional, for data-driven rules
)
```

The replacement engine handles:
- Iterating through worksheets
- Scanning cells within specified limits
- Applying match modes (exact vs substring)
- Navigating data paths for data-driven replacements
- Date formatting

## Data Flow

```
TextReplacementBuilder.build()
        ↓
        ├─ _replace_placeholders()
        │       ├─ Define placeholder rules
        │       │   (JFINV, JFTIME, JFREF, CUSTOMER_NAME, CUSTOMER_ADDRESS)
        │       ├─ Call find_and_replace()
        │       │       ├─ Scan A1:N14 across all sheets
        │       │       ├─ For each rule:
        │       │       │   ├─ Navigate invoice_data via data_path
        │       │       │   ├─ Find matching cells
        │       │       │   └─ Replace with data value
        │       │       └─ Return
        │       └─ Return
        │
        └─ _run_daf_specific_replacement()
                ├─ Define DAF rules
                │   (Location names, shipping terms)
                ├─ Call find_and_replace()
                │       ├─ Scan 200x16 grid across all sheets
                │       ├─ For each rule:
                │       │   ├─ Find matching cells
                │       │   └─ Replace with static value
                │       └─ Return
                └─ Return
```

## Key Design Decisions

### 1. **Two-Phase Replacement Strategy**
Separating placeholder replacement and DAF replacement allows:
- Different search areas (14x14 vs 200x16)
- Different rule types (data-driven vs static)
- Independent execution (can run one without the other)

### 2. **Limited Search Areas**
Instead of scanning entire workbooks:
- Placeholder replacement: `A1:N14` (header area only)
- DAF replacement: `200 rows × 16 columns` (header and data)

This significantly improves performance by avoiding unnecessary cell scanning.

### 3. **Centralized Replacement Engine**
The builder defines **what** to replace (rules) but delegates **how** to replace to the `find_and_replace()` utility. This:
- Separates configuration from implementation
- Allows reusing the replacement engine elsewhere
- Makes rules easy to modify without changing core logic

### 4. **Data Path Navigation**
Data-driven rules use a path list (`["processed_tables_data", "1", "inv_no", 0]`) to navigate nested dictionaries, providing flexibility for complex data structures.

### 5. **Mixed Match Modes**
DAF rules use both exact and substring matching:
- **Exact**: For complete term replacements (prevents partial matches)
- **Substring**: For Incoterm replacements within phrases (e.g., "FCA BAVET" → "DAF BAVET")

### 6. **Hardcoded DAF Rules**
DAF location mappings are hardcoded rather than data-driven because:
- Rules are business logic, not data
- Changes require code review (desired for business rules)
- Rules are consistent across all invoices

## Dependencies

- **openpyxl**: Core Excel manipulation library for workbook access
- **utils.text.find_and_replace**: The replacement engine that performs actual text finding and replacement operations
- **typing**: Type hints for better code clarity

## Usage Example

### Basic Usage - Called by LayoutBuilder

```python
from invoice_generator.builders.text_replacement_builder import TextReplacementBuilder

# In LayoutBuilder.build() method:
if enable_text_replacement:
    text_replacer = TextReplacementBuilder(
        workbook=self.workbook,
        invoice_data=self.invoice_data
    )
    
    if args.DAF:
        # DAF mode: Run both replacements
        text_replacer.build()
    else:
        # Standard mode: Run only placeholder replacement
        text_replacer._replace_placeholders()
```

### Standalone Usage

```python
import openpyxl
from invoice_generator.builders.text_replacement_builder import TextReplacementBuilder

# Load workbook
workbook = openpyxl.load_workbook('invoice_template.xlsx')

# Prepare invoice data
invoice_data = {
    'processed_tables_data': {
        '1': {
            'inv_no': ['JF25046'],
            'inv_date': ['2025-01-15'],
            'inv_ref': ['REF2025001']
        }
    },
    'customer_info': {
        'name': 'ABC Trading Co., Ltd.',
        'address': '123 Business Street, Phnom Penh, Cambodia'
    }
}

# Create and run text replacement
replacer = TextReplacementBuilder(workbook=workbook, invoice_data=invoice_data)
replacer.build()  # Runs both placeholder and DAF replacements

# Save result
workbook.save('invoice_output.xlsx')
```

### Selective Replacement

```python
# Run only placeholder replacement (skip DAF)
replacer = TextReplacementBuilder(workbook=workbook, invoice_data=invoice_data)
replacer._replace_placeholders()  # Only placeholders

# Or run only DAF replacement (skip placeholders)
replacer = TextReplacementBuilder(workbook=workbook, invoice_data=invoice_data)
replacer._run_daf_specific_replacement()  # Only DAF terms
```

## Example Replacements

### Before and After - Placeholder Replacement

**Before** (Template):
```
Invoice No: JFINV
Date: JFTIME
Reference: JFREF
Customer: [[CUSTOMER_NAME]]
Address: [[CUSTOMER_ADDRESS]]
```

**After** (With Data):
```
Invoice No: JF25046
Date: 15-Jan-2025
Reference: REF2025001
Customer: ABC Trading Co., Ltd.
Address: 123 Business Street, Phnom Penh, Cambodia
```

### Before and After - DAF Replacement

**Before** (Various Formats):
```
Origin: BINH PHUOC
Destination: BAVET, SVAY RIENG
Terms: FCA BAVET,SVAYRIENG
Port: PORT KLANG
Incoterm: CIF
```

**After** (Standardized):
```
Origin: BAVET
Destination: BAVET
Terms: DAF BAVET
Port: BAVET
Incoterm: DAF
```

## Integration with LayoutBuilder

The `TextReplacementBuilder` is called early in the layout building process:

```python
class LayoutBuilder:
    def build(self):
        # Phase 1: Text Replacement (optional, early phase)
        if self.enable_text_replacement:
            text_replacer = TextReplacementBuilder(
                workbook=self.workbook,
                invoice_data=self.invoice_data
            )
            if self.args.DAF:
                text_replacer.build()  # Both replacements
            else:
                text_replacer._replace_placeholders()  # Only placeholders
        
        # Phase 2-7: Template capture, header, data, footer building...
```

**Note**: Text replacement is typically disabled for performance (`enable_text_replacement=False`) unless placeholders exist in the template.

## Performance Considerations

### Search Area Limits

The builder uses limited search areas to improve performance:

| Replacement Type | Rows | Columns | Total Cells Scanned |
|-----------------|------|---------|---------------------|
| Placeholder | 14 | 14 | 196 cells per sheet |
| DAF | 200 | 16 | 3,200 cells per sheet |

For a workbook with 3 sheets:
- Placeholder: ~600 cells scanned
- DAF: ~9,600 cells scanned

Without limits, scanning entire sheets (e.g., 1000×50) would scan 50,000+ cells per sheet.

### When to Enable Text Replacement

Enable text replacement when:
- ✅ Template contains placeholders (JFINV, [[CUSTOMER_NAME]], etc.)
- ✅ DAF mode requires location standardization
- ✅ Template has location-specific text that varies by customer

Disable text replacement when:
- ❌ Template has no placeholders
- ❌ All values are written programmatically (no pre-existing text to replace)
- ❌ Performance is critical and replacement is unnecessary

## Notes

- The builder modifies the workbook **in-place**; it does not create a copy.
- Replacement rules are executed in order; earlier rules can affect later rules if they modify the same cells.
- Exact match is case-sensitive; `"bavet"` will not match `"BAVET"`.
- Substring match replaces all occurrences within a cell; `"FCA FCA"` with rule `FCA→DAF` becomes `"DAF DAF"`.
- Date formatting for `JFTIME` is handled by the replacement engine based on the `is_date: True` flag.
- DAF location rules are specific to Southeast Asian logistics and may need adjustment for other regions.
- The builder operates on **all sheets** in the workbook, not just the active sheet.
- Column limits (14 and 16) correspond to typical invoice template widths (columns A-N and A-P).
- The 200-row limit for DAF replacement accommodates large multi-table packing lists.
- Placeholder replacement is designed for header areas, while DAF replacement covers headers and data.








