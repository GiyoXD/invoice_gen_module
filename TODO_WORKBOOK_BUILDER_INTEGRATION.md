# TODO: Integrate WorkbookBuilder into Invoice Generation Flow

## Issue: Replace new worksheet approach with new workbook approach
**Priority**: High  
**Status**: Pending  
**Created**: 2025-10-27

## Problem
Currently, `LayoutBuilder` creates a new worksheet in the existing workbook to avoid merge conflicts. However, this still has some edge cases and complexity. The better approach is to create a completely NEW workbook for each sheet processing.

## Solution
Use the newly created `WorkbookBuilder` to:
1. **Open template workbook as READ-ONLY** - No risk of accidental modification
2. Create a brand new workbook (separate from template) using WorkbookBuilder
3. Create empty sheets with the same names as template sheets
4. Use template restoration (restore_header_only → write data → write footer → restore_footer_only)
5. Read from template (read-only), write to new workbook (writable)
6. This completely avoids any template conflicts since template is never modified

## Implementation Breakdown

### 🔍 Piece 1: Verify read-only template can be captured
**Goal**: Confirm TemplateStateBuilder can capture from read-only workbook  
**Priority**: Critical (foundation for everything else)  
**Estimated effort**: 15 minutes

**Tasks**:
- [ ] Create test script to load template as read-only
- [ ] Instantiate TemplateStateBuilder with read-only worksheet
- [ ] Verify state captured correctly (header rows, footer rows, merges)
- [ ] Confirm no errors when reading from read-only workbook

**Success criteria**: 
- ✅ No errors when capturing from read-only workbook
- ✅ Header state captured with correct row count and merges
- ✅ Footer state captured with correct row count and merges

**Risk**: Low - openpyxl read-only mode supports reading cells/merges

---

### 🏗️ Piece 2: Create WorkbookBuilder integration point
**Goal**: Modify generate_invoice.py to create output workbook with WorkbookBuilder  
**Priority**: High  
**Estimated effort**: 30 minutes

**Tasks**:
- [ ] Open template workbook as **read-only**: `load_workbook(template_path, read_only=True)`
- [ ] Collect all sheet names from template workbook
- [ ] Create `WorkbookBuilder` with those sheet names
- [ ] Call `workbook_builder.build()` to get a new clean workbook
- [ ] Store both template_workbook (read-only) and output_workbook (writable)

**Files to modify**:
- `invoice_generator/generate_invoice.py` (main function, around line 180-200)

**Success criteria**: 
- ✅ Template loaded as read-only
- ✅ New workbook created with correct sheet names
- ✅ Both workbooks available for next steps

**Risk**: Low - simple API usage

---

### 🔌 Piece 3: Update processor interface
**Goal**: Processors accept both template_workbook and output_workbook  
**Priority**: High  
**Estimated effort**: 45 minutes

**Tasks**:
- [ ] Modify `BaseProcessor.__init__()` to accept `template_workbook` and `output_workbook` parameters
- [ ] Update `SingleTableProcessor.__init__()` to pass both workbooks
- [ ] Update `MultiTableProcessor.__init__()` to pass both workbooks
- [ ] Update processor instantiation in `generate_invoice.py` to pass both workbooks
- [ ] Store references to both template and output worksheets

**Files to modify**:
- `invoice_generator/processors/base_processor.py`
- `invoice_generator/processors/single_table_processor.py`
- `invoice_generator/processors/multi_table_processor.py`
- `invoice_generator/generate_invoice.py` (processor instantiation)

**Success criteria**: 
- ✅ Processors instantiate without errors
- ✅ Both template and output worksheets accessible
- ✅ No breaking changes to existing processor logic

**Risk**: Medium - interface change affects multiple files

---

### 🎨 Piece 4: Update LayoutBuilder
**Goal**: LayoutBuilder works with separate template/output worksheets  
**Priority**: High  
**Estimated effort**: 1 hour

**Tasks**:
- [ ] Add `template_worksheet` parameter to `__init__()` 
- [ ] Update `build()` to use `template_worksheet` for capture, `self.worksheet` for output
- [ ] Remove internal workbook creation code (lines ~51-60)
- [ ] Remove `self.new_workbook` storage
- [ ] Remove cleanup section (lines ~258-270) that deletes/renames sheets
- [ ] Update TemplateStateBuilder to capture from template_worksheet
- [ ] Ensure all builders write to self.worksheet (output)
- [ ] Update print statements to clarify template vs output

**Files to modify**:
- `invoice_generator/builders/layout_builder.py`

**Success criteria**: 
- ✅ LayoutBuilder accepts separate template/output worksheets
- ✅ Captures from template, builds to output
- ✅ No internal workbook creation
- ✅ No cleanup code needed

**Risk**: Medium - needs careful parameter passing

---

### ✅ Piece 5: Test single sheet generation
**Goal**: Verify one sheet (Invoice) generates correctly  
**Priority**: High  
**Estimated effort**: 30 minutes

**Tasks**:
- [ ] Run invoice generation with new approach on single sheet
- [ ] Verify output file is created
- [ ] Open Excel file and check for errors
- [ ] Verify no merge conflict warnings
- [ ] Check header content and formatting
- [ ] Check data rows are correct
- [ ] Check footer content (both dynamic and template)
- [ ] Verify row heights and column widths

**Test command**:
```powershell
python -m invoice_generator.generate_invoice invoice_generator\JF.json -t invoice_generator\template -c invoice_generator\config
```

**Success criteria**: 
- ✅ Excel file opens without repair warnings
- ✅ No merge conflicts
- ✅ Content correct and formatted
- ✅ Template footer present after data footer

**Risk**: Low - integration test

---

### 🎯 Piece 6: Test multi-sheet generation
**Goal**: Verify all sheets generate correctly  
**Priority**: Medium  
**Estimated effort**: 30 minutes

**Tasks**:
- [ ] Run invoice generation with full multi-sheet data
- [ ] Verify all sheets present in output
- [ ] Check Invoice sheet
- [ ] Check Contract sheet
- [ ] Check Packing list sheet (multi-table)
- [ ] Verify no errors in any sheet
- [ ] Check merge cells in all sheets
- [ ] Verify formatting preserved

**Test command**:
```powershell
python -m invoice_generator.generate_invoice invoice_generator\CLW.json -t invoice_generator\template -c invoice_generator\config
```

**Success criteria**: 
- ✅ All sheets present and correct
- ✅ No merge conflicts in any sheet
- ✅ Multi-table sheet (Packing list) works correctly
- ✅ All formatting preserved

**Risk**: Low - should work if Piece 5 works

---

## Execution Order
1️⃣ Piece 1 (Verify read-only) → 2️⃣ Piece 2 (WorkbookBuilder integration) → 3️⃣ Piece 3 (Processor interface) → 4️⃣ Piece 4 (LayoutBuilder) → 5️⃣ Piece 5 (Single sheet test) → 6️⃣ Piece 6 (Multi-sheet test)

**Recommended approach**: Complete each piece fully before moving to the next. Each piece builds on the previous one.

## Benefits
✅ **No merge conflicts** - New workbook starts completely clean  
✅ **Template protected** - Read-only mode prevents accidental modification  
✅ **No cleanup needed** - No sheets to delete/rename/move  
✅ **Clearer separation** - Template stays untouched, output is separate  
✅ **Simpler logic** - Just read from one, write to another  
✅ **Better architecture** - Each builder has a single clear responsibility  
✅ **Performance** - Read-only mode can be faster for large templates  

## Current State
- [x] WorkbookBuilder created and ready
- [ ] Integration pending
- [ ] Needs testing after integration

## Files to Modify
1. `invoice_generator/generate_invoice.py` - Main orchestration
2. `invoice_generator/builders/layout_builder.py` - Remove internal workbook creation
3. `invoice_generator/processors/single_table_processor.py` - Accept template + output workbooks
4. `invoice_generator/processors/multi_table_processor.py` - Accept template + output workbooks
5. `invoice_generator/processors/base_processor.py` - Update interface if needed

## Notes
- This is a cleaner architecture than the current "new worksheet" approach
- **Template workbook opened as read-only** - cannot be modified even accidentally
- All writes go to the new clean workbook created by WorkbookBuilder
- No risk of merge conflicts since new workbook has no existing merges
- The template file on disk stays pristine and can be reused infinitely
