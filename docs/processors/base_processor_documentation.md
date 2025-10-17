# Processor Architecture: `base_processor.py`

This document explains the structure and purpose of the `SheetProcessor` abstract base class, which is the foundation of our Strategy/Processor design pattern.

## `SheetProcessor` Abstract Base Class (ABC)

The `SheetProcessor` class defines a mandatory contract or interface that all concrete sheet processors (like `SingleTableProcessor` and `MultiTableProcessor`) must follow. It ensures that the main orchestrator can interact with any type of processor in a consistent way, without needing to know the specific details of its implementation.

- **Purpose**: To enforce a standard structure for all sheet processing logic.
- **Technology**: It uses Python's `abc` module (`ABC` and `abstractmethod`) to define the abstract interface.

### `__init__(...)` - The Constructor

The constructor is responsible for initializing the processor with all the data and configuration it needs to do its job. It acts as a dependency injection mechanism, ensuring that every processor instance has access to the workbook, its specific worksheet, configuration, and the complete dataset.

- **Purpose**: To equip a new processor instance with its required context.
- **Parameters**:
    - `workbook: Workbook`: The active `openpyxl` Workbook object. This allows the processor to interact with other sheets if necessary.
    - `worksheet: Worksheet`: The specific `openpyxl` Worksheet object that this processor instance is responsible for modifying.
    - `sheet_name: str`: The name of the worksheet.
    - `sheet_config: Dict`: The portion of the configuration file that applies specifically to this sheet (e.g., `data_mapping["Invoice"]`).
    - `data_mapping_config: Dict`: The entire `data_mapping` section from the main configuration.
    - `data_source_indicator: str`: A key (e.g., `"processed_tables_multi"`) that tells the processor which part of the main `invoice_data` to use as its primary data source.
    - `invoice_data: Dict`: The complete, raw data dictionary loaded from the input file.
    - `cli_args: argparse.Namespace`: The parsed command-line arguments, allowing the processor to respond to user-input flags (like `--custom`).
    - `final_grand_total_pallets: int`: A pre-calculated value that can be used by any processor that needs to display this global total.

### `process(self) -> bool` - The Abstract Method

This is the core of the abstract class. It defines the one method that all subclasses *must* implement.

- **Purpose**: To serve as the main entry point for the processor's logic. The orchestrator calls this single method to kick off all processing for the sheet.
- **Implementation**: As an `@abstractmethod`, the `SheetProcessor` base class provides no implementation for `process()`. It is the sole responsibility of the concrete subclasses (e.g., `SingleTableProcessor`) to provide the actual logic.
- **Return Value**:
    - `True`: Indicates that the sheet was processed successfully.
    - `False`: Indicates that a critical error occurred, signaling to the orchestrator that it should stop the entire generation process.
