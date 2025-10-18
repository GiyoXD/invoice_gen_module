# Gemini Context: Invoice Generator Project

## Project Overview

This project is a Python-based invoice generator designed to create Excel-based invoices and packing lists from structured data (JSON or pickle files) and Excel templates. The system is undergoing a significant refactoring from a monolithic script-based approach to a modular, component-based architecture.

The new architecture leverages the **Strategy** and **Builder** design patterns to create a flexible and extensible generation pipeline. The core idea is to use a declarative approach where JSON configuration files act as a "recipe" to define the structure, content, and styling of the final document.

### Key Components:

*   **`processors` (Strategy Pattern):** Defines the high-level algorithm for generating a specific document type (e.g., `SingleTableProcessor` for standard invoices, `MultiTableProcessor` for packing lists).
*   **`builders` (Builder Pattern):** Responsible for constructing the different visual parts of the Excel document, such as the header, data tables, and footer. The `LayoutBuilder` acts as a Director, orchestrating the other builders.
*   **`config`:** Handles loading, validation, and modeling of the JSON configuration files using Pydantic models.
*   **`styling`:** Centralizes all styling and formatting logic, acting as a theme engine.
*   **`data`:** Manages data transformation and preparation before it's written to the sheet.
*   **`utils`:** Contains low-level, reusable helper functions for Excel operations.

The main entry point is `invoice_generator/generate_invoice.py`, which orchestrates the entire process.

## Building and Running

### Running the Application

The invoice generator is a command-line application. To run it, you need to provide the path to an input data file and specify the directories for templates and configurations.

**Command:**

```bash
python -m invoice_generator.generate_invoice <path_to_data_file> --output <output_path.xlsx> --templatedir <template_directory> --configdir <config_directory>
```

**Example:**

```bash
python -m invoice_generator.generate_invoice "invoice_gen/data/CLW.json" --output "result.xlsx" --templatedir "invoice_gen/TEMPLATE" --configdir "invoice_gen/config"
```

### Running Tests

The project uses Python's built-in `unittest` framework. The tests are located in the `tests/` directory and mirror the structure of the main `invoice_generator` package.

The primary test suite is an end-to-end integration test that runs the main script with sample data.

**Command to run tests:**

```bash
python -m unittest tests/test_invoice_generation.py
```

## Development Conventions

*   **Modularity:** All new logic should be placed in the appropriate package (`builders`, `processors`, `styling`, etc.) as outlined in the `REFACTORING.md` document.
*   **Configuration over Code:** When adding support for new invoice formats or variations, prefer extending the JSON configuration rather than writing new code.
    *   For minor layout changes, add flags to the configuration that the `Builders` can read.
    *   For different high-level document structures, create a new `Processor` strategy.
*   **Testing:** Every new component (e.g., a new builder) should have a corresponding test file in the `tests` directory (e.g., `tests/builders/test_new_builder.py`). The testing strategy is bottom-up, starting with unit tests for utilities and data handlers, followed by integration tests for builders and processors.
*   **File Naming:**
    *   Data files are expected to follow a pattern like `{customer_name}.json`.
    *   Template files: `{customer_name}.xlsx`.
    *   Configuration files: `{customer_name}_config.json`.
    The `derive_paths` function in `generate_invoice.py` handles this logic.

## Gemini Specific Rules

*   **No Modification of Config/Template Directories:** Gemini must *not* modify any files within the `config` or `TEMPLATE` directories. Only source code files can be modified.
*   **Post-Issue Resolution Invoice Generation:** After resolving a GitHub issue, Gemini *must* generate a new invoice using the updated source code before asking the user for confirmation or marking the issue as fully resolved.