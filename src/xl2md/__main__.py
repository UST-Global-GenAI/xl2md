# src/xl2md/__main__.py

import argparse
import logging
import sys

from .converter import ConverterOptions, ExcelToMarkdownConverter


def main():
    """Main function for the command-line interface."""
    parser = argparse.ArgumentParser(
        description="Convert all sheets in an Excel workbook to Markdown files."
    )
    parser.add_argument(
        "excel_path",
        help="Path to the input Excel file (.xlsx, .xls)."
    )
    parser.add_argument(
        "--out-dir",
        default="./markdown_sheets",
        help="Directory to save the output Markdown files (default: ./markdown_sheets)."
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite existing Markdown files if they exist."
    )
    parser.add_argument(
        "--no-safe-filenames",
        dest="safe_filenames",
        action="store_false",
        help="Use original sheet names for filenames instead of 'slugified' versions."
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_const",
        dest="log_level",
        const=logging.DEBUG,
        default=logging.INFO,
        help="Enable verbose (debug) logging output."
    )
    
    args = parser.parse_args()

    # Create options from command-line arguments
    options = ConverterOptions(
        out_dir=args.out_dir,
        overwrite=args.overwrite,
        safe_filenames=args.safe_filenames,
        log_level=args.log_level,
    )

    try:
        # Initialize and run the converter
        converter = ExcelToMarkdownConverter(excel_path=args.excel_path, options=options)
        converter.convert()
    except Exception as e:
        # Use a basic logger for errors if initialization fails
        logging.basicConfig(level=logging.ERROR)
        logging.error("A critical error occurred: %s", e, exc_info=False)
        sys.exit(1)


if __name__ == "__main__":
    main()