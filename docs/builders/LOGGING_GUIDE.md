# Python Logging Guide for Invoice Generator

## Logging Levels (from lowest to highest severity)

### 1. **DEBUG** - `logger.debug()`
**When to use:** Detailed diagnostic information for developers debugging the code
- Variable values during execution
- Step-by-step flow tracking
- Internal state information
- "Entering function X", "Variable Y = Z"

**Example:**
```python
logger.debug(f"skip_data_table_builder = {self.skip_data_table_builder}")
logger.debug(f"Calling FooterBuilder.build() with footer_row_position={footer_row_position}")
```

### 2. **INFO** - `logger.info()`
**When to use:** General informational messages about normal program execution
- Major steps completed successfully
- What the program is doing at a high level
- Confirmation of expected behavior
- "Starting process X", "Completed task Y"

**Example:**
```python
logger.info(f"Building layout for sheet '{self.sheet_name}'")
logger.info(f"Restoring template footer after row {write_pointer_row}")
logger.info(f"Layout built successfully for sheet '{self.sheet_name}'")
```

### 3. **WARNING** - `logger.warning()`
**When to use:** Something unexpected happened, but program continues
- Deprecated functionality being used
- Recoverable errors
- Configuration issues that don't prevent execution
- "Using fallback value", "Deprecated method called"

**Example:**
```python
logger.warning(f"Using legacy data source resolution. Consider using BuilderConfigResolver instead.")
logger.warning(f"Column width config missing, using default values")
```

### 4. **ERROR** - `logger.error()`
**When to use:** Serious problem that prevented a function from completing
- Builder failures that require halting execution
- Missing required data or configuration
- Operations that failed but program can potentially recover
- "Failed to X", "Cannot complete Y because Z"

**Example:**
```python
logger.error(f"HeaderBuilder failed for sheet '{self.sheet_name}'")
logger.error(f"Invalid next_row_after_footer={self.next_row_after_footer} - HALTING EXECUTION")
logger.error(f"Cannot restore template footer - invalid write_pointer_row={write_pointer_row}")
```

### 5. **CRITICAL** - `logger.critical()`
**When to use:** Very serious error that may cause program termination
- System-level failures
- Unrecoverable errors
- Data corruption
- "Database connection lost", "Cannot write to disk"

**Example:**
```python
logger.critical(f"Failed to open workbook: {file_path}")
logger.critical(f"Out of memory while processing large dataset")
```

---

## How to Configure Logging

### Basic Configuration (in main entry point)

Add this to `generate_invoice.py` or your main script:

```python
import logging

# Configure logging at the start of your program
logging.basicConfig(
    level=logging.INFO,  # Set the minimum level to display
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# To see DEBUG messages, change level to:
# level=logging.DEBUG
```

### Output Format Examples

With the above configuration:
```
2025-11-11 16:30:45 - invoice_generator.builders.layout_builder - INFO - Building layout for sheet 'Invoice'
2025-11-11 16:30:45 - invoice_generator.builders.layout_builder - DEBUG - skip_data_table_builder = False
2025-11-11 16:30:46 - invoice_generator.builders.layout_builder - ERROR - FooterBuilder failed for sheet 'Invoice'
```

### Controlling Log Levels

**Show only WARNING and above (hide INFO and DEBUG):**
```python
logging.basicConfig(level=logging.WARNING)
```

**Show everything including DEBUG:**
```python
logging.basicConfig(level=logging.DEBUG)
```

**Show only ERROR and CRITICAL:**
```python
logging.basicConfig(level=logging.ERROR)
```

---

## Quick Reference

| Level | Use When | Example |
|-------|----------|---------|
| DEBUG | Detailed diagnostics for developers | `logger.debug(f"Variable x = {x}")` |
| INFO | Normal operation milestones | `logger.info("Processing started")` |
| WARNING | Unexpected but recoverable | `logger.warning("Using fallback value")` |
| ERROR | Function failed, halting operation | `logger.error("Builder failed")` |
| CRITICAL | System-level failure | `logger.critical("Cannot save file")` |

---

## Best Practices

1. **Use DEBUG for development, INFO for production** - Set level to INFO in production to avoid excessive logs
2. **ERROR should always indicate a problem** - If you log an ERROR, something failed
3. **Include context in messages** - Add relevant variable values and identifiers
4. **Use consistent format** - Follow the patterns established in the codebase
5. **Don't log sensitive data** - Avoid logging passwords, API keys, personal information

---

## Implementation Status

✅ **layout_builder.py** - Logging implemented with proper levels
- ERROR: Builder failures and halting conditions
- INFO: Major steps and successful completions  
- DEBUG: Internal state and variable tracking

🔲 **Other builders** - Can be migrated to logging as needed
- header_builder.py
- data_table_builder.py
- footer_builder.py
- template_state_builder.py
