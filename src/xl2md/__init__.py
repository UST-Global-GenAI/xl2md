# src/xl2md/__init__.py

"""
xl2md: A utility to convert Excel sheets to Markdown tables.
"""

__version__ = "0.1.0"

# Expose the main components for easy import
# e.g., from xl2md import ExcelToMarkdownConverter
from .converter import (
    ConverterOptions,
    ExcelToMarkdownConverter,
    ExcelToMarkdownError,
    SheetConversionError,
    WorkbookReadError,
)