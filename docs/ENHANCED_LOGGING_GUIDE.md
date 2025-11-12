# Enhanced Logging Configuration Guide

## Current vs Enhanced Logging Output

### What You Have Now (Minimal):
```
>>> ERROR >>> FooterBuilder failed for sheet 'Invoice'
>>> ERROR >>> Invalid next_row_after_footer=-1 - HALTING EXECUTION
```

### What You Can Have (Detailed):
```
>>> ERROR >>> [layout_builder.py:245 in build()] FooterBuilder failed for sheet 'Invoice'
>>> ERROR >>> [layout_builder.py:247 in build()] Invalid next_row_after_footer=-1 - HALTING EXECUTION
Traceback (most recent call last):
  File "invoice_generator/builders/layout_builder.py", line 240, in build
    success, footer_row, _, _, _ = footer_builder.build()
  File "invoice_generator/builders/footer_builder.py", line 156, in build
    return -1
ValueError: Invalid footer configuration
```

---

## 🔧 How to Enhance Your Logging

### Option 1: Add File/Line Info to Format (Recommended)

**Current format**:
```python
formatter = ColoredFormatter('>>> %(levelname)s >>> %(message)s')
```

**Enhanced format**:
```python
# Shows: filename, line number, function name
formatter = ColoredFormatter(
    '>>> %(levelname)s >>> [%(filename)s:%(lineno)d in %(funcName)s()] %(message)s'
)
```

**Output**:
```
>>> ERROR >>> [footer_builder.py:156 in build()] FooterBuilder failed for sheet 'Invoice'
```

---

### Option 2: Add Timestamps (Know When Errors Happen)

```python
formatter = ColoredFormatter(
    '%(asctime)s >>> %(levelname)s >>> [%(filename)s:%(lineno)d] %(message)s',
    datefmt='%H:%M:%S'
)
```

**Output**:
```
14:32:15 >>> ERROR >>> [footer_builder.py:156] FooterBuilder failed for sheet 'Invoice'
```

---

### Option 3: Full Debug Format (Maximum Detail)

```python
formatter = ColoredFormatter(
    '%(asctime)s [%(name)s] %(levelname)s [%(pathname)s:%(lineno)d in %(funcName)s()]\n'
    '  └─ %(message)s'
)
```

**Output**:
```
14:32:15 [invoice_generator.builders.footer_builder] ERROR [C:\...\footer_builder.py:156 in build()]
  └─ FooterBuilder failed for sheet 'Invoice'
```

---

## 🔥 How to Show Stack Traces

### Current Issue:
Your code logs errors but **doesn't show the stack trace**:

```python
# Current (no stack trace)
logger.error("FooterBuilder failed")
```

### Solution 1: Use `exc_info=True`

```python
# Shows full stack trace
try:
    footer_builder.build()
except Exception as e:
    logger.error(f"FooterBuilder failed: {e}", exc_info=True)
```

### Solution 2: Use `logger.exception()`

```python
# Automatically includes stack trace
try:
    footer_builder.build()
except Exception as e:
    logger.exception(f"FooterBuilder failed for sheet '{sheet_name}'")
```

**Output with stack trace**:
```
>>> ERROR >>> [layout_builder.py:245] FooterBuilder failed for sheet 'Invoice'
Traceback (most recent call last):
  File "invoice_generator/builders/layout_builder.py", line 240, in build
    success, footer_row, _, _, _ = footer_builder.build()
  File "invoice_generator/builders/footer_builder.py", line 156, in build
    if not self.footer_config:
TypeError: 'NoneType' object is not subscriptable
```

Now you can see **EXACTLY where and why** it failed!

---

## 📝 Recommended Logging Format

### For Development (Most Detail):

```python
# In generate_invoice.py around line 250
formatter = ColoredFormatter(
    '>>> %(levelname)s >>> [%(filename)s:%(lineno)d in %(funcName)s()] %(message)s'
)
```

### For Production (Clean but Informative):

```python
formatter = ColoredFormatter(
    '%(asctime)s >>> %(levelname)s >>> [%(filename)s:%(lineno)d] %(message)s',
    datefmt='%H:%M:%S'
)
```

---

## 🎯 How to Fix Your Current Error Logging

### Find and Update Error Logging in LayoutBuilder

Look for code like this:

```python
# Current (line ~245-247 in layout_builder.py)
if not success or next_row_after_footer < 0:
    logger.error(f"FooterBuilder failed for sheet '{self.sheet_name}'")
    logger.error(f"Invalid next_row_after_footer={next_row_after_footer} - HALTING EXECUTION")
    return False
```

**Change to**:

```python
# Enhanced (shows WHY it failed)
if not success or next_row_after_footer < 0:
    logger.error(
        f"FooterBuilder failed for sheet '{self.sheet_name}': "
        f"success={success}, next_row_after_footer={next_row_after_footer}"
    )
    
    # Log the footer builder's internal state for debugging
    logger.debug(f"  footer_config: {layout_config.get('footer_config')}")
    logger.debug(f"  footer_row_num: {footer_row_position}")
    
    return False
```

---

## 🚀 Complete Enhanced Setup

Here's the complete code to replace in `generate_invoice.py`:

```python
# Around line 230-255 in generate_invoice.py

# Custom formatter with ANSI color codes AND location info
class ColoredFormatter(logging.Formatter):
    COLORS = {
        'DEBUG': '\033[36m',      # Cyan
        'INFO': '\033[0m',         # White (reset)
        'WARNING': '\033[33m',     # Yellow
        'ERROR': '\033[31m',       # Red
        'CRITICAL': '\033[35m',    # Magenta
        'RESET': '\033[0m'         # Reset
    }
    
    def format(self, record):
        color = self.COLORS.get(record.levelname, self.COLORS['RESET'])
        record.levelname = f"{color}{record.levelname}{self.COLORS['RESET']}"
        return super().format(record)

# Enhanced formatter with file/line/function info
formatter = ColoredFormatter(
    '>>> %(levelname)s >>> [%(filename)s:%(lineno)d in %(funcName)s()] %(message)s'
)

# Rest stays the same...
stdout_handler = logging.StreamHandler(sys.stdout)
stdout_handler.setLevel(logging.DEBUG)
stdout_handler.setFormatter(formatter)
stdout_handler.addFilter(lambda record: record.levelno <= logging.INFO)
logger.addHandler(stdout_handler)

stderr_handler = logging.StreamHandler(sys.stderr)
stderr_handler.setLevel(logging.WARNING)
stderr_handler.setFormatter(formatter)
logger.addHandler(stderr_handler)
```

---

## 📊 Before & After Comparison

### Before (Your Current Output):
```
>>> INFO >>> Restoring header from template to output worksheet
>>> INFO >>> Skipping header builder
>>> INFO >>> Entering data table builder block
>>> INFO >>> Using resolver-provided resolved_data
>>> ERROR >>> FooterBuilder failed for sheet 'Invoice'
>>> ERROR >>> Invalid next_row_after_footer=-1 - HALTING EXECUTION
```

❌ **Can't tell**:
- Which file has the error?
- What line number?
- What function failed?
- Why did it fail?

### After (With Enhanced Logging):
```
>>> INFO >>> [layout_builder.py:180 in build()] Restoring header from template
>>> INFO >>> [layout_builder.py:195 in build()] Skipping header builder (skip_header_builder=True)
>>> INFO >>> [layout_builder.py:210 in build()] Entering data table builder block
>>> INFO >>> [layout_builder.py:215 in build()] Using resolver-provided resolved_data (modern approach)
>>> ERROR >>> [layout_builder.py:245 in build()] FooterBuilder failed for sheet 'Invoice': success=False, next_row_after_footer=-1
>>> DEBUG >>> [layout_builder.py:248 in build()]   footer_config: None
>>> DEBUG >>> [layout_builder.py:249 in build()]   footer_row_num: 25
>>> ERROR >>> [layout_builder.py:251 in build()] HALTING EXECUTION due to footer failure
```

✅ **Now you know**:
- **File**: `layout_builder.py`
- **Line**: 245
- **Function**: `build()`
- **Why**: `footer_config` is `None`, `next_row_after_footer=-1`

---

## 🎓 Logging Best Practices

### 1. INFO Level - Keep it Brief (Never Mind Being Descriptive)
```python
# ✅ Good (concise, just the facts)
logger.info("Processing sheet 'Invoice'")
logger.info("Creating workbook")
logger.info("Template loaded")
logger.info("Data resolved")

# ❌ Too verbose for INFO
logger.info("Processing sheet 'Invoice' using SingleTableProcessor with aggregation data source type")
logger.info("Creating new workbook with 3 sheets: Invoice, Packing list, Contract with template from JF.xlsx")
```

### 2. DEBUG/WARNING/ERROR - Be Descriptive (Include Details)
```python
# ✅ Good (detailed when debugging/troubleshooting)
logger.debug(f"Resolver state: sheet={sheet_name}, pallets={pallets}, data_type={data_source_type}")
logger.warning(f"Missing 'col_po' in column_id_map, using fallback 'N/A'")
logger.error(f"FooterBuilder failed: success={success}, next_row={next_row}, config={footer_config}")
logger.exception(f"Exception in LayoutBuilder for sheet '{sheet_name}': {e}")
```

### 3. Always Include Context in Errors
```python
# ❌ Bad (not descriptive enough for errors)
logger.error("Failed to build")

# ✅ Good (descriptive - shows what, why, and values)
logger.error(f"Failed to build footer for sheet '{sheet_name}': success={success}, next_row={next_row}")

# ✅ Better (with stack trace)
logger.error(f"Failed to build footer: {e}", exc_info=True)
```

### 4. Use Correct Log Levels
```python
logger.debug(f"Variable x = {x}, y = {y}")           # Detailed diagnostics (verbose OK)
logger.info(f"Processing sheet {name}")              # Brief progress updates (concise)
logger.warning(f"Missing config, using default")    # Unexpected but recoverable (explain)
logger.error(f"Build failed: {reason}")              # Something broke (explain with details)
logger.exception(f"Critical failure: {e}")           # Fatal error (with full stack trace)
```

---

## 🛠️ Quick Fix for Your Current Issue

Want to immediately see where the FooterBuilder error comes from? Add this:

```python
# In layout_builder.py, find the FooterBuilder error section
# Add exc_info=True or use logger.exception()

try:
    success, next_row_after_footer, _, _, _ = footer_builder.build()
    if not success or next_row_after_footer < 0:
        logger.error(
            f"FooterBuilder failed for sheet '{self.sheet_name}': "
            f"success={success}, next_row={next_row_after_footer}"
        )
        # Add this to see what went wrong:
        logger.debug(f"FooterBuilder state: config={footer_builder.footer_config}")
        return False
except Exception as e:
    logger.exception(f"Exception in FooterBuilder for sheet '{self.sheet_name}'")
    return False
```

This will show you the full stack trace and why the footer builder returned -1.
