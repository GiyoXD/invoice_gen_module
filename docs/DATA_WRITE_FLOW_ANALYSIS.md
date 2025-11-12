# DataTableBuilder Write Flow - Deep Dive

## Current Data Flow Architecture

```
TableDataResolver.resolve()
    ↓ Returns resolved_data dict
    ↓
DataTableBuilder.__init__(resolved_data)
    ↓ Extracts data_rows list
    ↓
DataTableBuilder.build()
    ↓ Loops through data_rows
    ↓
worksheet.cell(row, col).value = value
    ↓
Excel file (openpyxl)
```

---

## 1. Data Format: What `data_rows` Looks Like

### Current Format (Dict-Based)

`data_rows` is a **List of Dictionaries**, where each dict represents one row:

```python
data_rows = [
    {  # Row 1
        1: "PO123",              # Column 1 (PO)
        2: "ITEM001",            # Column 2 (Item)
        3: "LEATHER TYPE A",     # Column 3 (Description)
        4: 100.50,               # Column 4 (SqFt)
        5: 5.25,                 # Column 5 (Unit Price)
        6: {                     # Column 6 (Amount) - FORMULA
            "type": "formula",
            "template": "{col_ref_1}{row}*{col_ref_0}{row}",
            "inputs": ["col_qty_sf", "col_unit_price"]
        }
    },
    {  # Row 2
        1: "PO123",
        2: "ITEM002",
        3: "LEATHER TYPE B",
        4: 200.75,
        5: 6.50,
        6: {"type": "formula", ...}
    }
]
```

**Key Characteristics**:
- ✅ **Keys are column indices** (1-based, matching Excel columns)
- ✅ **Values are mixed types**: strings, numbers, or formula dicts
- ✅ **Sparse**: Only columns with data are included (not all 20 columns present)
- ✅ **Formula support**: Special dict format with `{"type": "formula", ...}`

---

## 2. The Writing Loop (Line 88-104)

### Step-by-Step Breakdown

```python
# Line 74: Calculate starting row (after header)
data_writing_start_row = self.header_info.get('second_row_index', 0) + 1

# Line 88: Loop through each data row
for i in range(actual_rows_to_process):
    current_row_idx = data_start_row + i  # Excel row number (e.g., 22, 23, 24...)
    row_data = self.data_rows[i]          # Get dict: {col_idx: value}
    
    # Line 93: Loop through each column in this row
    for col_idx, value in row_data.items():
        # Line 95: Get the Excel cell object
        cell = self.worksheet.cell(row=current_row_idx, column=col_idx)
        
        # Line 96: Safety check - skip if cell is merged
        if not isinstance(cell, MergedCell):
            
            # Line 98: Check if value is a formula
            if isinstance(value, dict) and value.get('type') == 'formula':
                # Convert formula dict → Excel formula string
                formula_str = self._build_formula_string(value, current_row_idx)
                cell.value = formula_str  # e.g., "=B22*C22"
            else:
                # Write regular value
                cell.value = value  # String, number, etc.
```

### What Happens for Each Cell:

1. **Get cell object**: `worksheet.cell(row=22, column=4)`
2. **Check if merged**: Skip if part of merged region
3. **Check value type**:
   - If formula dict → Convert to `"=B22*C22"` string
   - If regular value → Write directly
4. **Write**: `cell.value = ...`

---

## 3. Formula Handling (Line 113-131)

### Formula Dict Structure

```python
formula_value = {
    "type": "formula",
    "template": "{col_ref_1}{row}*{col_ref_0}{row}",  # Template with placeholders
    "inputs": ["col_qty_sf", "col_unit_price"]         # Column IDs
}
```

### Conversion Process

```python
def _build_formula_string(self, formula_dict, row_num):
    # 1. Extract template and inputs
    template = "{col_ref_1}{row}*{col_ref_0}{row}"
    inputs = ["col_qty_sf", "col_unit_price"]
    
    # 2. Replace {col_ref_0}, {col_ref_1}, etc. with column letters
    # inputs[0] = "col_qty_sf" → col_idx = 4 → col_letter = "D"
    # inputs[1] = "col_unit_price" → col_idx = 5 → col_letter = "E"
    formula = template
    for i, input_id in enumerate(inputs):
        col_idx = self.col_id_map.get(input_id)  # "col_qty_sf" → 4
        col_letter = get_column_letter(col_idx)  # 4 → "D"
        formula = formula.replace(f'{{col_ref_{i}}}', col_letter)
        # After: "D{row}*E{row}"
    
    # 3. Replace {row} with actual row number
    formula = formula.replace('{row}', str(row_num))  # "D22*E22"
    
    # 4. Ensure starts with =
    if not formula.startswith('='):
        formula = '=' + formula  # "=D22*E22"
    
    return formula  # Final: "=D22*E22"
```

### Example Transformations

| Template | Inputs | Row | Result |
|----------|--------|-----|--------|
| `{col_ref_1}{row}*{col_ref_0}{row}` | `["col_qty_sf", "col_unit_price"]` | 22 | `=D22*E22` |
| `{col_ref_1}{row}/{col_ref_0}{row}` | `["col_qty_sf", "col_amount"]` | 23 | `=F23/D23` |
| `SUM({col_ref_0}{row}:{col_ref_1}{row})` | `["col_net", "col_gross"]` | 24 | `=SUM(G24:H24)` |

---

## 4. Where Data Comes From (data_preparer.py)

### The `prepare_data_rows()` Function

This is called by **TableDataResolver** and returns:

```python
(
    data_rows_prepared,      # List[Dict[int, Any]] - what we've been analyzing
    pallet_counts_for_rows,  # List[int] - pallet count per row
    dynamic_desc_used,       # bool - whether dynamic descriptions were used
    num_data_rows_from_source # int - total rows from source data
)
```

### Data Source Type Handlers

#### Aggregation (Most Common)
```python
# Input: Tuple-based aggregation
aggregation_data = {
    ('PO123', 'ITEM001', 5.25, 'LEATHER TYPE A'): {
        'sqft_sum': 100.50,
        'amount_sum': 527.625
    }
}

# Output: Dict-based rows
data_rows = [
    {1: "PO123", 2: "ITEM001", 3: "LEATHER TYPE A", 4: 100.50, 5: 5.25, 6: formula_dict}
]
```

#### Processed Tables Multi (Packing Lists)
```python
# Input: Column-array format
table_data = {
    'po': ['PO123', 'PO124'],
    'item': ['ITEM001', 'ITEM002'],
    'sqft': [100.50, 200.75],
    'pallet_count': [2, 3]
}

# Output: Dict-based rows
data_rows = [
    {1: "PO123", 2: "ITEM001", 4: 100.50},
    {1: "PO124", 2: "ITEM002", 4: 200.75}
]
```

#### DAF Aggregation (Special Shipping Term)
```python
# Input: DAF-specific aggregation
DAF_data = {
    'row1': {
        'combined_po': 'PO123',
        'combined_item': 'ITEM001',
        'total_sqft': 100.50
    }
}

# Output: Same dict-based rows format
```

---

## 5. Current Limitations & Why You Might Want to Rewrite

### Current Issues:

1. **Sparse Dict Format is Inefficient**:
   - Each row stores `{1: val, 2: val, 5: val}` 
   - Missing columns not represented (could be `None` instead)
   - Hard to validate completeness

2. **Complex Formula Dict**:
   - Formula handling embedded in write loop
   - Template system with placeholders is indirect
   - Mixing data and formulas in same structure

3. **Tight Coupling**:
   - Writer knows about formula dict structure
   - Can't easily swap formats without changing writer

4. **No Validation**:
   - No type checking before write
   - No column bounds checking
   - Errors caught at write-time, not prep-time

---

## 6. Alternative Format Ideas

### Option A: Dense List Format (Array-Based)

```python
# Each row is a list with ALL columns (even if None)
data_rows = [
    ["PO123", "ITEM001", "LEATHER", 100.50, 5.25, Formula("=D22*E22"), None, None, ...],
    ["PO124", "ITEM002", "SUEDE",   200.75, 6.50, Formula("=D23*E23"), None, None, ...]
]

# Writer becomes:
for row_idx, row_data in enumerate(data_rows):
    for col_idx, value in enumerate(row_data, start=1):
        if value is not None:
            cell = worksheet.cell(row=start_row + row_idx, column=col_idx)
            if isinstance(value, Formula):
                cell.value = value.to_excel_string(row=start_row + row_idx)
            else:
                cell.value = value
```

**Pros**:
- ✅ Simple iteration (no dict lookups)
- ✅ Clear column ordering
- ✅ Easy to validate row length
- ✅ Formula objects encapsulate logic

**Cons**:
- ❌ More memory (stores None for empty columns)
- ❌ Requires knowing total column count upfront

---

### Option B: Named Tuple / Dataclass Format (Type-Safe)

```python
from dataclasses import dataclass
from typing import Optional

@dataclass
class InvoiceRow:
    po: str
    item: str
    description: str
    sqft: float
    unit_price: float
    amount: Formula  # Custom Formula class
    net: Optional[float] = None
    gross: Optional[float] = None

data_rows = [
    InvoiceRow(
        po="PO123",
        item="ITEM001",
        description="LEATHER",
        sqft=100.50,
        unit_price=5.25,
        amount=Formula.multiply("sqft", "unit_price")
    )
]

# Writer becomes:
column_mapping = {
    'po': 1, 'item': 2, 'description': 3, 
    'sqft': 4, 'unit_price': 5, 'amount': 6
}

for row_idx, row_obj in enumerate(data_rows):
    for field_name, col_idx in column_mapping.items():
        value = getattr(row_obj, field_name)
        if value is not None:
            cell = worksheet.cell(row=start_row + row_idx, column=col_idx)
            if isinstance(value, Formula):
                cell.value = value.to_excel_string(row=start_row + row_idx, mappings=column_mapping)
            else:
                cell.value = value
```

**Pros**:
- ✅ Type safety (IDE autocomplete)
- ✅ Validation at creation time
- ✅ Clear field names (self-documenting)
- ✅ Easy to test

**Cons**:
- ❌ More boilerplate
- ❌ Less flexible for dynamic columns
- ❌ Requires schema definition

---

### Option C: Pandas DataFrame Format (Most Flexible)

```python
import pandas as pd

# Data as DataFrame
data_rows = pd.DataFrame({
    'PO': ['PO123', 'PO124'],
    'Item': ['ITEM001', 'ITEM002'],
    'SqFt': [100.50, 200.75],
    'Unit Price': [5.25, 6.50],
    'Amount': [Formula("=D{row}*E{row}"), Formula("=D{row}*E{row}")]
})

# Writer becomes:
for row_idx, row in data_rows.iterrows():
    for col_name, col_idx in column_mapping.items():
        value = row[col_name]
        if pd.notna(value):
            cell = worksheet.cell(row=start_row + row_idx, column=col_idx)
            if isinstance(value, Formula):
                cell.value = value.to_excel_string(row=start_row + row_idx)
            else:
                cell.value = value
```

**Pros**:
- ✅ Powerful data manipulation
- ✅ Built-in validation/cleaning
- ✅ Easy to debug (can print df.head())
- ✅ Integration with data analysis tools

**Cons**:
- ❌ Heavy dependency (pandas)
- ❌ Overkill for simple use case
- ❌ Memory overhead

---

## 7. Recommended Approach

### Create a Custom Formula Class First

```python
# invoice_generator/data/formula.py
class ExcelFormula:
    """Represents an Excel formula to be written to a cell."""
    
    def __init__(self, template: str, inputs: List[str]):
        self.template = template
        self.inputs = inputs
    
    def to_excel_string(self, row_num: int, col_id_map: Dict[str, int]) -> str:
        """Convert to Excel formula string like '=D22*E22'"""
        formula = self.template
        for i, input_id in enumerate(self.inputs):
            col_idx = col_id_map.get(input_id)
            if col_idx:
                col_letter = get_column_letter(col_idx)
                formula = formula.replace(f'{{col_ref_{i}}}', col_letter)
        formula = formula.replace('{row}', str(row_num))
        if not formula.startswith('='):
            formula = '=' + formula
        return formula
    
    @classmethod
    def multiply(cls, col1: str, col2: str):
        """Helper: Create multiplication formula"""
        return cls("{col_ref_0}{row}*{col_ref_1}{row}", [col1, col2])
    
    @classmethod
    def divide(cls, col1: str, col2: str):
        """Helper: Create division formula"""
        return cls("{col_ref_0}{row}/{col_ref_1}{row}", [col1, col2])
```

### Then Simplify the Writer

```python
# New simplified write loop
for i, row_data in enumerate(self.data_rows):
    current_row_idx = data_start_row + i
    
    for col_idx, value in row_data.items():
        cell = self.worksheet.cell(row=current_row_idx, column=col_idx)
        if isinstance(cell, MergedCell):
            continue
        
        # Simplified: Formula class handles its own conversion
        if isinstance(value, ExcelFormula):
            cell.value = value.to_excel_string(current_row_idx, self.col_id_map)
        else:
            cell.value = value
```

---

## 8. Migration Strategy

### Phase 1: Extract Formula Logic
1. Create `ExcelFormula` class in `invoice_generator/data/formula.py`
2. Update `data_preparer.py` to return `ExcelFormula` objects instead of dicts
3. Test with existing data

### Phase 2: Simplify Writer
1. Remove `_build_formula_string()` from DataTableBuilder
2. Update write loop to use `isinstance(value, ExcelFormula)`
3. Test all invoice types

### Phase 3: (Optional) Switch to Dense Format
1. Modify `prepare_data_rows()` to return lists instead of dicts
2. Update writer to iterate with `enumerate()`
3. Benchmark performance difference

---

## Summary

**Current Format**: `List[Dict[int, Any]]` where values can be data or formula dicts
**Current Writer**: Loops dict items, checks formula dict format, converts to Excel string
**Pain Points**: Sparse format, complex formula handling, tight coupling

**Best Next Step**: 
1. Create `ExcelFormula` class to encapsulate formula logic
2. Keep dict-based format (it works well for sparse data)
3. Simplify writer by delegating formula conversion to the Formula object

This gives you a cleaner separation without requiring a full rewrite of data preparation logic.
