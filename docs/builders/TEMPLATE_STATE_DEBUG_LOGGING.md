# TemplateStateBuilder Debug Logging Enhancement

## Overview

Added comprehensive debug logging to `TemplateStateBuilder` to provide visibility into what template values and styles are being captured during the state capture process.

## What Was Added

### 1. New Helper Method: `_format_cell_style_info()`

Located after `_is_default_style()`, this method formats cell styling information into readable debug output.

**Formats:**
- **Font**: name, size, bold, italic, color (RGB)
- **Fill**: fill_type, start_color (RGB)
- **Border**: left, right, top, bottom styles (e.g., "thin", "medium")
- **Alignment**: horizontal, vertical, wrap_text

**Example Output:**
```
A5: value='Invoice Number', font(name=Calibri, size=11, bold), fill(solid, FFFF00), border(B:thin)
```

### 2. Enhanced `_capture_header()` Method

**Before:**
```
>>> DEBUG >>> Capturing header up to row 21
>>> DEBUG >>> Header captured: 5 rows, 2 merges
```

**After:**
```
>>> DEBUG >>> [template_state_builder.py:198 in _capture_header()] === CAPTURING HEADER (rows 1 to 21) ===
>>> DEBUG >>> [template_state_builder.py:207 in _capture_header()]   Header starts at row 1, ends at row 21
>>> DEBUG >>> [template_state_builder.py:208 in _capture_header()]   Max columns: 13
>>> DEBUG >>> [template_state_builder.py:234 in _capture_header()]   Row 5: A5='SELLER:', B5='Company Name Ltd.', G5='Invoice No.'
>>> DEBUG >>> [template_state_builder.py:240 in _capture_header()]     Style: A5: value='SELLER:', font(name=Calibri, size=11, bold)
>>> DEBUG >>> [template_state_builder.py:240 in _capture_header()]     Style: B5: value='Company Name Ltd.', font(name=Calibri, size=11), border(B:thin)
>>> DEBUG >>> [template_state_builder.py:248 in _capture_header()]   Captured 2 merged cells: A5:A6, G5:H5
>>> DEBUG >>> [template_state_builder.py:254 in _capture_header()]   [OK]OK] Header capture complete: 21 rows, 2 merges
```

**What's Logged:**
- Header row range (start/end)
- Max columns being scanned
- **Each row with content**: Shows cell coordinates and values (first 5 non-empty cells)
- **Styling details**: First 2 styled cells per row with font/fill/border/alignment info
- **Merged cells**: List of merged cell ranges (first 3, with count if more)
- **Summary**: Total rows captured and merge count

### 3. Enhanced `_capture_footer()` Method

**Before:**
```
>>> DEBUG >>> Capturing footer from row 42 to 180
>>> DEBUG >>> Footer ends at row 47 (found 6 footer rows)
>>> DEBUG >>> Footer captured: 6 rows, 1 merges, start row: 42
```

**After:**
```
>>> DEBUG >>> [template_state_builder.py:259 in _capture_footer()] === CAPTURING FOOTER (starting from row 42) ===
>>> DEBUG >>> [template_state_builder.py:260 in _capture_footer()]   Max columns: 13, max search row: 180
>>> DEBUG >>> [template_state_builder.py:283 in _capture_footer()]   Footer ends at row 47 (6 footer rows)
>>> DEBUG >>> [template_state_builder.py:308 in _capture_footer()]   Row 42: A42='Total Net Weight (KGS)', H42='1,234.56'
>>> DEBUG >>> [template_state_builder.py:314 in _capture_footer()]     Style: A42: value='Total Net Weight (KGS)', font(name=Calibri, size=10, bold), border(T:thin, B:thin)
>>> DEBUG >>> [template_state_builder.py:314 in _capture_footer()]     Style: H42: value='1,234.56', font(name=Calibri, size=10), align(h=right)
>>> DEBUG >>> [template_state_builder.py:325 in _capture_footer()]   Captured 1 merged cells: A42:G42
>>> DEBUG >>> [template_state_builder.py:333 in _capture_footer()]   [OK]OK]OK] Footer capture complete: 6 rows, 1 merges, template footer start: 42
```

**What's Logged:**
- Footer search parameters (start row, max search row, max columns)
- Footer detection result (end row, total rows found)
- **Each row with content**: Cell coordinates and values (first 5 non-empty cells)
- **Styling details**: First 2 styled cells per row with font/fill/border/alignment info
- **Merged cells**: List of merged cell ranges (first 3, with count if more)
- **Summary**: Total rows captured, merge count, template footer start row

## How to Use

### Running with Debug Logging

```powershell
# Enable DEBUG level logging
python -m invoice_generator.generate_invoice CLW.json --output result.xlsx --log-level DEBUG
```

### Expected Output Structure

```
=== CAPTURING HEADER (rows 1 to 21) ===
  Header starts at row 1, ends at row 21
  Max columns: 13
  Row 1: A1='COMMERCIAL INVOICE'
    Style: A1: value='COMMERCIAL INVOICE', font(name=Arial, size=14, bold), fill(solid, FF4472C4), align(h=center)
  Row 5: A5='SELLER:', B5='Company XYZ'
    Style: A5: value='SELLER:', font(name=Calibri, size=11, bold)
    Style: B5: value='Company XYZ', font(name=Calibri, size=11)
  Captured 2 merged cells: A1:H1, A5:A6
  [OK]OK]OK] Header capture complete: 21 rows, 2 merges

=== CAPTURING FOOTER (starting from row 42) ===
  Max columns: 13, max search row: 180
  Footer ends at row 47 (6 footer rows)
  Row 42: A42='Total:', H42='$12,345.67'
    Style: A42: value='Total:', font(name=Calibri, size=10, bold), border(T:thin, B:double)
  [OK]OK]OK] Footer capture complete: 6 rows, 1 merges, template footer start: 42
```

## Why This Matters

### Before Enhancement
- Only saw counts: "Header captured: 5 rows, 2 merges"
- **No visibility** into what values were captured
- **No visibility** into what styling was preserved
- Debugging template restoration issues required manual Excel inspection

### After Enhancement
- See **actual cell values** being captured
- See **specific styling** (fonts, colors, borders, alignment)
- See **merged cell ranges** explicitly
- Can diagnose template restoration issues from logs alone

### Real-World Use Cases

1. **Template Restoration Debugging**
   - User: "Why is my invoice header bold not showing?"
   - Debug log shows: `A5: value='SELLER:', font(name=Calibri, size=11)` (no bold!)
   - Root cause: Template cell wasn't bold to begin with

2. **Merge Cell Issues**
   - User: "Footer text is split across cells instead of merged"
   - Debug log shows: `Captured 0 merged cells` in footer
   - Root cause: Template footer didn't have merged cells

3. **Missing Border Lines**
   - User: "Invoice table bottom border disappeared"
   - Debug log shows: `A21: value='Total', border(L:thin, R:thin)` (no bottom!)
   - Root cause: Template footer first row missing top border

## Integration with Logging System

Uses the project's level-based logging format:
- **DEBUG level** shows `[filename:lineno in funcName()]` prefix
- **INFO level** would be clean (but template capture only logs at DEBUG)

Example:
```
>>> DEBUG >>> [template_state_builder.py:198 in _capture_header()] === CAPTURING HEADER ===
```

## Performance Considerations

- **Minimal overhead**: Style formatting only happens for cells with styling (not every cell)
- **Log spam prevention**: Limited to first 5 values per row, first 2 styled cells per row, first 3 merges
- **Conditional execution**: All debug logging skipped when log level > DEBUG

## Testing

To verify the enhanced logging works:

```powershell
# Run test invoice generation with DEBUG logging
cd c:\Users\JPZ031127\Desktop\refactor
python -m invoice_generator.generate_invoice CLW.json --output test_debug.xlsx --log-level DEBUG

# Should see detailed template capture logs in console
```

Expected output should show:
- Header capture section with cell values and styles
- Footer capture section with cell values and styles
- No errors or exceptions

## Related Documentation

- **ENHANCED_LOGGING_GUIDE.md**: Overall logging configuration and best practices
- **template_state_builder_documentation.md**: Full TemplateStateBuilder API reference
- **.github/copilot-instructions.md**: Project conventions including "NO print statements" rule
