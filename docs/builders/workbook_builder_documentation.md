# Builder Architecture: `workbook_builder.py`

This document explains the structure and purpose of the `WorkbookBuilder` class, which is responsible for creating clean, empty workbooks with specified sheet names, ready for content population.

## Overview

The `WorkbookBuilder` is the foundational builder that creates the initial workbook structure. It generates a new, clean workbook with properly named empty worksheets, removing the default sheet created by openpyxl. This provides a blank canvas for other builders to populate with content from templates.

- **Purpose**: To create a new, clean workbook with specified sheet names without any template artifacts or conflicts.
- **Pattern**: Builder pattern - constructs workbook structure step-by-step.
- **Key Responsibility**: Generates the workbook foundation that other builders populate with headers, data, and footers.

## Key Concepts

### Clean Slate Approach

The builder creates a **completely new workbook** rather than copying or modifying a template workbook. This approach:
- Eliminates template corruption risks
- Avoids merge conflicts from template modifications
- Provides full control over workbook structure
- Enables template state capture and selective restoration

### Template-Separate Architecture

The workbook building strategy separates concerns:
1. **Template Workbook**: Source of visual structure (read-only)
2. **Output Workbook**: Created by WorkbookBuilder (writable)
3. **TemplateStateBuilder**: Captures template state
4. **Other Builders**: Restore and populate output workbook

## `WorkbookBuilder` Class

### `__init__(...)` - The Constructor

The constructor initializes the builder with the list of sheet names to create.

- **Purpose**: To configure the builder with the sheet structure for the new workbook.
- **Parameters**:
    - `sheet_names: List[str]`: List of sheet names to create in the workbook (e.g., `["Invoice", "Packing list", "Contract"]`).
- **Instance Variables Initialized**:
    - `sheet_names`: Stores the list of sheet names to create.
    - `workbook`: Initially `None`, set to the created workbook after `build()` is called.

### `build(self) -> Workbook` - The Main Build Method

Creates a new workbook with the specified sheets and returns it.

- **Purpose**: To construct a new, clean workbook with properly named empty sheets ready for population.
- **Process**:
    1. **Create New Workbook**: Instantiate a new `Workbook` object using openpyxl.
    2. **Remove Default Sheet**: Delete the default `'Sheet'` created by openpyxl (if present).
    3. **Create Named Sheets**: For each name in `sheet_names`:
        - Create a new worksheet with the specified name
        - Log the sheet creation
    4. **Store Reference**: Store the workbook in `self.workbook` for later access.
    5. **Return Workbook**: Return the created workbook object.

- **Return Value**: `Workbook`
    - A new openpyxl `Workbook` instance containing empty worksheets with the specified names.

- **Console Output**: Prints progress messages:
    ```
    [WorkbookBuilder] Creating new workbook with 3 sheets
    [WorkbookBuilder]   Created sheet: 'Invoice'
    [WorkbookBuilder]   Created sheet: 'Packing list'
    [WorkbookBuilder]   Created sheet: 'Contract'
    [WorkbookBuilder] New workbook created successfully
    ```

### `get_worksheet(self, sheet_name: str) -> Worksheet` - Sheet Retrieval Helper

Retrieves a specific worksheet from the created workbook by name.

- **Purpose**: To provide convenient access to individual worksheets after the workbook has been created.
- **Parameters**:
    - `sheet_name: str`: The name of the sheet to retrieve (must match one of the names provided during initialization).
- **Return Value**: `Worksheet`
    - The requested openpyxl `Worksheet` object.
- **Error Handling**:
    - **RuntimeError**: If `build()` hasn't been called yet (workbook is `None`).
    - **ValueError**: If the requested sheet name doesn't exist in the workbook.

## Data Flow

```
WorkbookBuilder.__init__(sheet_names)
        ↓
Store sheet_names: ["Invoice", "Packing list", "Contract"]
        ↓
WorkbookBuilder.build()
        ↓
    ┌───────────────────────────┐
    ↓                           ↓
Create New Workbook      workbook = Workbook()
    ↓                           ↓
Remove Default Sheet     del workbook['Sheet']
    ↓
For each sheet_name:
    ├─ workbook.create_sheet(title=sheet_name)
    └─ Print creation log
        ↓
Store workbook reference
        ↓
Return workbook
        ↓
Other builders populate the sheets
        ↓
Save workbook to file
```

## Key Design Decisions

### 1. **New Workbook Creation vs Template Modification**
The builder creates a **new workbook** rather than modifying a template. This:
- **Prevents template corruption**: Original template remains pristine
- **Avoids merge conflicts**: No pre-existing merges to conflict with new ones
- **Enables parallel processing**: Multiple processes can use the same template
- **Simplifies error recovery**: If generation fails, template is unaffected

### 2. **Explicit Sheet Name List**
Requiring sheet names as a parameter (rather than discovering them from template) provides:
- **Control**: Caller decides exactly which sheets to create
- **Flexibility**: Can create subset of template sheets if needed
- **Clarity**: Sheet structure is explicit in the calling code

### 3. **Remove Default Sheet**
openpyxl always creates a default sheet named `'Sheet'`. The builder removes it to:
- **Maintain clean structure**: Only requested sheets exist
- **Match template structure**: Template doesn't have a `'Sheet'` worksheet
- **Avoid confusion**: Prevents accidental use of unexpected sheet

### 4. **Stored Workbook Reference**
The builder stores the created workbook in `self.workbook` to:
- **Enable sheet retrieval**: `get_worksheet()` can access sheets later
- **Support method chaining**: Multiple operations on the same workbook
- **Aid debugging**: Workbook state is accessible after build

### 5. **Logging for Visibility**
The builder prints progress messages to:
- **Provide feedback**: User sees workbook creation happening
- **Aid debugging**: Easy to trace which sheets were created
- **Confirm success**: Final message confirms completion

### 6. **Error-Safe Sheet Retrieval**
`get_worksheet()` validates state before access:
- **Guards against premature access**: Ensures `build()` was called
- **Validates sheet existence**: Prevents KeyError with helpful message
- **Provides clear errors**: Error messages guide the developer

## Dependencies

- **openpyxl**: Core Excel manipulation library
    - `Workbook` - Workbook creation
    - `Worksheet` - Worksheet type hint
- **typing**: Type hints for better code clarity

## Usage Example

### Basic Usage - Create and Build Workbook

```python
from invoice_generator.builders.workbook_builder import WorkbookBuilder

# Define sheet structure
sheet_names = ["Invoice", "Packing list", "Contract"]

# Create builder
builder = WorkbookBuilder(sheet_names=sheet_names)

# Build the workbook
workbook = builder.build()

# Output:
# [WorkbookBuilder] Creating new workbook with 3 sheets
# [WorkbookBuilder]   Created sheet: 'Invoice'
# [WorkbookBuilder]   Created sheet: 'Packing list'
# [WorkbookBuilder]   Created sheet: 'Contract'
# [WorkbookBuilder] New workbook created successfully

print(f"Created workbook with {len(workbook.sheetnames)} sheets")
print(f"Sheet names: {workbook.sheetnames}")
# Output:
# Created workbook with 3 sheets
# Sheet names: ['Invoice', 'Packing list', 'Contract']
```

### Retrieve Specific Worksheets

```python
# After building the workbook
builder = WorkbookBuilder(["Invoice", "Packing list"])
workbook = builder.build()

# Retrieve individual sheets
invoice_sheet = builder.get_worksheet("Invoice")
packing_sheet = builder.get_worksheet("Packing list")

print(f"Invoice sheet: {invoice_sheet.title}")
print(f"Packing sheet: {packing_sheet.title}")

# Or access directly from workbook
contract_sheet = workbook["Contract"]  # Raises KeyError if not exists
```

### Error Handling Examples

```python
builder = WorkbookBuilder(["Invoice"])

# Error: Trying to get worksheet before build()
try:
    sheet = builder.get_worksheet("Invoice")
except RuntimeError as e:
    print(f"Error: {e}")
    # Output: Error: Workbook not created yet. Call build() first.

# Build workbook first
workbook = builder.build()

# Error: Requesting non-existent sheet
try:
    sheet = builder.get_worksheet("NonExistent")
except ValueError as e:
    print(f"Error: {e}")
    # Output: Error: Sheet 'NonExistent' not found in workbook
```

### Integration with Template State Restoration

```python
from invoice_generator.builders.workbook_builder import WorkbookBuilder
from invoice_generator.builders.template_state_builder import TemplateStateBuilder
import openpyxl

# Load template
template_workbook = openpyxl.load_workbook('template.xlsx', data_only=False)
template_sheet = template_workbook['Invoice']

# Create clean output workbook
output_builder = WorkbookBuilder(sheet_names=["Invoice"])
output_workbook = output_builder.build()
output_sheet = output_builder.get_worksheet("Invoice")

# Capture template state
state_builder = TemplateStateBuilder(
    worksheet=template_sheet,
    num_header_cols=6,
    header_end_row=2,
    footer_start_row=20
)

# Restore template structure to clean workbook
state_builder.restore_header_only(target_worksheet=output_sheet)

# Now output_sheet has template header, ready for data population
print("Template structure restored to clean workbook")
```

### Complete Workflow Example

```python
from invoice_generator.builders.workbook_builder import WorkbookBuilder
from invoice_generator.builders.layout_builder import LayoutBuilder
import openpyxl

# Step 1: Load template (read-only)
template_workbook = openpyxl.load_workbook('template.xlsx', read_only=True)

# Step 2: Create clean output workbook
sheet_names = ["Invoice", "Packing list", "Contract"]
wb_builder = WorkbookBuilder(sheet_names=sheet_names)
output_workbook = wb_builder.build()

# Step 3: Process each sheet
for sheet_name in sheet_names:
    template_sheet = template_workbook[sheet_name]
    output_sheet = wb_builder.get_worksheet(sheet_name)
    
    # Prepare config bundles for LayoutBuilder
    style_config = {
        'styling_config': styling
    }
    
    context_config = {
        'sheet_name': sheet_name,
        'invoice_data': data,
        'all_sheet_configs': config,
        'args': args
    }
    
    layout_config = {
        'sheet_config': config[sheet_name]
    }
    
    # Use LayoutBuilder to populate output_sheet from template_sheet
    layout_builder = LayoutBuilder(
        workbook=output_workbook,
        worksheet=output_sheet,
        template_worksheet=template_sheet,
        style_config=style_config,
        context_config=context_config,
        layout_config=layout_config
    )
    
    success = layout_builder.build()
    if not success:
        print(f"Failed to build {sheet_name}")
        break

# Step 4: Save output workbook
output_workbook.save('invoice_output.xlsx')
template_workbook.close()
```

## Workbook Structure Comparison

### Before WorkbookBuilder (openpyxl default)
```python
wb = Workbook()
print(wb.sheetnames)
# Output: ['Sheet']  ← Unwanted default sheet
```

### After WorkbookBuilder
```python
builder = WorkbookBuilder(["Invoice", "Packing list"])
wb = builder.build()
print(wb.sheetnames)
# Output: ['Invoice', 'Packing list']  ← Only requested sheets
```

## Integration with Other Builders

The `WorkbookBuilder` is typically the **first builder** in the workflow:

```
1. WorkbookBuilder → Creates clean workbook
        ↓
2. TemplateStateBuilder → Captures template state
        ↓
3. For each sheet:
    ├─ TemplateStateBuilder → Restores template header
    ├─ HeaderBuilder → Builds header content
    ├─ DataTableBuilder → Populates data
    ├─ FooterBuilder → Creates footers
    └─ TemplateStateBuilder → Restores template footer
        ↓
4. Save workbook to file
```

## Performance Considerations

### Lightweight Operation
Creating a workbook is very fast:
- **Memory**: New workbook has minimal memory footprint
- **Time**: Sheet creation is nearly instantaneous (< 1ms per sheet)
- **Scalability**: Can create workbooks with many sheets efficiently

### No File I/O
The builder only creates in-memory structures:
- **Fast**: No disk access during build
- **Safe**: No file corruption risk
- **Flexible**: Can modify before saving

## Notes

- The builder creates **empty worksheets** with no content, formatting, or merges.
- Sheet names must follow Excel naming rules (max 31 characters, no special characters like `[]:/\?*`).
- The default openpyxl sheet `'Sheet'` is always removed if present.
- The created workbook is stored in `self.workbook` and can be accessed after `build()`.
- `get_worksheet()` is a convenience method; sheets can also be accessed directly via `workbook[sheet_name]`.
- The builder doesn't validate sheet names; invalid names will raise errors when calling `create_sheet()`.
- Sheet order matches the order in `sheet_names` list.
- The builder doesn't copy any content from templates; use `TemplateStateBuilder` for that.
- Created worksheets have default Excel dimensions (empty, no frozen panes, no print settings).
- The workbook can be saved immediately after creation, though it will be an empty workbook with named sheets.
- This builder is typically used in conjunction with `TemplateStateBuilder` to create and populate workbooks from templates.

## Common Patterns

### Pattern 1: Single-Sheet Workbook
```python
builder = WorkbookBuilder(["Report"])
workbook = builder.build()
sheet = builder.get_worksheet("Report")
# Populate sheet...
workbook.save("report.xlsx")
```

### Pattern 2: Multi-Sheet Workbook
```python
builder = WorkbookBuilder(["Sales", "Inventory", "Summary"])
workbook = builder.build()

for sheet_name in ["Sales", "Inventory", "Summary"]:
    sheet = builder.get_worksheet(sheet_name)
    # Populate each sheet...

workbook.save("report.xlsx")
```

### Pattern 3: Conditional Sheet Creation
```python
sheet_names = ["Invoice"]
if include_packing:
    sheet_names.append("Packing list")
if include_contract:
    sheet_names.append("Contract")

builder = WorkbookBuilder(sheet_names)
workbook = builder.build()
```

### Pattern 4: Direct Workbook Access
```python
builder = WorkbookBuilder(["Data"])
workbook = builder.build()

# Access sheets directly from workbook
data_sheet = workbook["Data"]
workbook.active = workbook["Data"]  # Set active sheet
```

## Comparison with Other Approaches

### Approach 1: Copy Template Workbook
```python
# ❌ Not recommended - modifies template
template = openpyxl.load_workbook('template.xlsx')
# ... modify template ...
template.save('output.xlsx')  # Template is now modified!
```

### Approach 2: Load Template and Save As
```python
# ❌ Still has template artifacts
template = openpyxl.load_workbook('template.xlsx')
# ... modify ...
template.save('output.xlsx')  # Has template data/formatting
```

### Approach 3: WorkbookBuilder + TemplateStateBuilder
```python
# ✅ Recommended - clean separation
output = WorkbookBuilder(["Invoice"]).build()
state = TemplateStateBuilder(template_sheet, ...)
state.restore_header_only(output["Invoice"])
# ... build content ...
output.save('output.xlsx')  # Clean output, template unchanged
```

## Troubleshooting

### Issue: Default 'Sheet' still present
**Cause**: Custom implementation not removing it correctly  
**Solution**: The builder handles this automatically; check openpyxl version

### Issue: Sheet names not matching template
**Cause**: Typo in sheet_names list  
**Solution**: Ensure exact name match (case-sensitive)

### Issue: RuntimeError when getting worksheet
**Cause**: Called `get_worksheet()` before `build()`  
**Solution**: Always call `build()` first

### Issue: ValueError for sheet name
**Cause**: Requesting sheet not in sheet_names list  
**Solution**: Verify sheet name matches one from initialization

## Summary

The `WorkbookBuilder` is a simple but essential builder that:
- ✅ Creates clean, empty workbooks
- ✅ Removes openpyxl's default sheet
- ✅ Creates specified named sheets
- ✅ Provides convenient sheet retrieval
- ✅ Enables template-separate architecture
- ✅ Supports any number of sheets
- ✅ Maintains workbook reference for later access

It serves as the **foundation** for the entire invoice generation system, providing the blank canvas that other builders populate with content.



