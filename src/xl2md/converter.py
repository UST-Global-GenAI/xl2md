# src/xl2md/converter.py

from __future__ import annotations

import logging
import math
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Sequence

import pandas as pd

# (ConverterOptions and Exception classes remain unchanged)
@dataclass
class ConverterOptions:
    out_dir: str = "./markdown_sheets"
    include_index: bool = False
    index_label: Optional[str] = None
    header: bool = True
    skip_empty_sheets: bool = True
    engine: str = "openpyxl"
    safe_filenames: bool = True
    title_prefix: str = ""
    log_level: int = logging.INFO
    overwrite: bool = True
    sheet_name_allowlist: Sequence[str] = field(default_factory=list)
    sheet_name_blocklist: Sequence[str] = field(default_factory=list)


class ExcelToMarkdownError(Exception):
    """Base class for conversion errors."""


class WorkbookReadError(ExcelToMarkdownError):
    """Raised when the Excel file cannot be read."""


class SheetConversionError(ExcelToMarkdownError):
    """Raised when a single sheet fails to convert."""


class ExcelToMarkdownConverter:
    def __init__(self, excel_path: str, options: Optional[ConverterOptions] = None, logger: Optional[logging.Logger] = None):
        self.excel_path = Path(excel_path)
        self.options = options or ConverterOptions()
        self.logger = logger or logging.getLogger(self.__class__.__name__)
        self._configure_logger()

        if not self.excel_path.exists():
            raise WorkbookReadError(f"File not found: {self.excel_path}")

        if not self.options.title_prefix:
            self.options.title_prefix = self.excel_path.stem

        self.logger.debug(
            "Initialized ExcelToMarkdownConverter with path=%s, options=%s",
            self.excel_path, self.options
        )

    # --------- Public API ---------

    def convert(self) -> List[str]:
        """
        Convert all sheets in the input file (Excel or CSV) to Markdown files.
        Returns a list of written file paths.
        """
        self.logger.info("Starting conversion for: %s", self.excel_path)
        
        # --- MODIFIED SECTION ---
        # Detect file type and dispatch to the correct handler.
        file_suffix = self.excel_path.suffix.lower()
        if file_suffix == '.csv':
            return self._convert_csv()
        elif file_suffix in ['.xlsx', '.xls']:
            return self._convert_excel()
        else:
            raise WorkbookReadError(
                f"Unsupported file type: '{file_suffix}'. Only .xlsx, .xls, and .csv are supported."
            )
        # --- END MODIFIED SECTION ---

    # --------- Internals ---------

    # --- NEW METHOD ---
    def _convert_csv(self) -> List[str]:
        """Handles the conversion of a single CSV file."""
        self.logger.debug("Processing as CSV file.")
        out_dir = Path(self.options.out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)

        try:
            df = pd.read_csv(self.excel_path)
        except Exception as e:
            self.logger.exception("Failed reading CSV file '%s'", self.excel_path)
            raise SheetConversionError(f"Failed reading CSV file '{self.excel_path}': {e}") from e

        if self.options.skip_empty_sheets and df.empty:
            self.logger.info("Skipping empty CSV file: %s", self.excel_path)
            return []

        # A CSV has no "sheet name", so we use the file's name (without extension)
        sheet_name = self.excel_path.stem
        md_text = self._sheet_to_markdown(sheet_name, df)
        filename = self._compose_filename(sheet_name)
        out_path = out_dir / filename

        if out_path.exists() and not self.options.overwrite:
            self.logger.warning("File exists and overwrite=False: %s (skipping write)", out_path)
            return []
        
        out_path.write_text(md_text, encoding="utf-8")
        self.logger.info("Wrote: %s", out_path)
        self.logger.info("Conversion complete. 1 file written.")
        return [str(out_path)]
    # --- END NEW METHOD ---

    # --- NEW METHOD (refactored from original `convert`) ---
    def _convert_excel(self) -> List[str]:
        """Handles the conversion of an Excel workbook."""
        self.logger.debug("Processing as Excel file.")
        xls = self._open_workbook()
        out_dir = Path(self.options.out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        
        written_paths: List[str] = []

        for sheet_name in xls.sheet_names:
            if not self._sheet_allowed(sheet_name):
                self.logger.debug("Skipping sheet due to allow/block rules: %s", sheet_name)
                continue

            try:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
            except Exception as e:
                self.logger.exception("Failed reading sheet '%s'", sheet_name)
                raise SheetConversionError(f"Failed reading sheet '{sheet_name}': {e}") from e

            if self.options.skip_empty_sheets and df.empty:
                self.logger.info("Skipping empty sheet: %s", sheet_name)
                continue

            md_text = self._sheet_to_markdown(sheet_name, df)
            filename = self._compose_filename(sheet_name)
            out_path = out_dir / filename

            if out_path.exists() and not self.options.overwrite:
                self.logger.warning("File exists and overwrite=False: %s (skipping write)", out_path)
            else:
                out_path.write_text(md_text, encoding="utf-8")
                self.logger.info("Wrote: %s", out_path)
                written_paths.append(str(out_path))

        if not written_paths:
            self.logger.warning("No sheets were converted. Check filters and file content.")
        else:
            self.logger.info("Conversion complete. %d file(s) written.", len(written_paths))

        return written_paths
    # --- END NEW METHOD ---


    def _open_workbook(self) -> pd.ExcelFile:
        # This method is now only called for Excel files.
        try:
            xls = pd.ExcelFile(str(self.excel_path), engine=self.options.engine)
            self.logger.debug("Opened workbook. Sheets found: %s", xls.sheet_names)
            return xls
        except ImportError as e:
            msg = f"Missing engine '{self.options.engine}'. Try: pip install openpyxl"
            self.logger.exception(msg)
            raise WorkbookReadError(msg) from e
        except Exception as e:
            msg = f"Failed to open workbook: {self.excel_path} ({e})"
            self.logger.exception(msg)
            raise WorkbookReadError(msg) from e

    # (All other methods like _sheet_allowed, _sheet_to_markdown, etc., remain unchanged)
    def _sheet_allowed(self, sheet_name: str) -> bool:
        for pat in self.options.sheet_name_blocklist:
            if re.search(pat, sheet_name, flags=re.I):
                return False
        if self.options.sheet_name_allowlist:
            for pat in self.options.sheet_name_allowlist:
                if re.search(pat, sheet_name, flags=re.I):
                    return True
            return False
        return True

    def _sheet_to_markdown(self, sheet_name: str, df: pd.DataFrame) -> str:
        self.logger.debug("Converting sheet to markdown: %s (shape=%s)", sheet_name, df.shape)
        title = f"# {self.options.title_prefix} — {sheet_name}\n\n"
        table_md = self._df_to_markdown_table(
            df,
            include_index=self.options.include_index,
            index_label=self.options.index_label,
            header=self.options.header
        )
        return title + table_md + "\n"

    @staticmethod
    def _is_nan(x) -> bool:
        try:
            return pd.isna(x)
        except Exception:
            try:
                return x is None or (isinstance(x, float) and math.isnan(x))
            except Exception:
                return False

    @classmethod
    def _to_str(cls, x) -> str:
        if cls._is_nan(x):
            return ""
        s = str(x)
        s = s.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "<br>")
        s = s.replace("\\", "\\\\").replace("|", r"\|")
        return s.strip()

    @classmethod
    def _clean_header(cls, col) -> str:
        if isinstance(col, str) and col.lower().startswith("unnamed:"):
            return ""
        return cls._to_str(col)

    @classmethod
    def _df_to_markdown_table(
        cls,
        df: pd.DataFrame,
        include_index: bool = False,
        index_label: Optional[str] = None,
        header: bool = True,
    ) -> str:
        if not isinstance(df, pd.DataFrame):
            raise TypeError("df must be a pandas DataFrame")
        col_labels: List[str] = []
        if include_index:
            col_labels.append(cls._to_str(index_label) if index_label is not None else "")
        for c in df.columns:
            col_labels.append(cls._clean_header(c))
        if len(col_labels) == 0:
            col_labels = [""]
        lines: List[str] = []
        if header:
            lines.append("| " + " | ".join(col_labels) + " |")
            lines.append("| " + " | ".join(["---"] * len(col_labels)) + " |")
        if include_index:
            for idx, row in df.iterrows():
                cells = [cls._to_str(idx)] + [cls._to_str(v) for v in row.tolist()]
                lines.append("| " + " | ".join(cells) + " |")
        else:
            for _, row in df.iterrows():
                cells = [cls._to_str(v) for v in row.tolist()]
                if len(cells) < len(df.columns):
                    cells += [""] * (len(df.columns) - len(cells))
                lines.append("| " + " | ".join(cells) + " |")
        if df.shape[0] == 0 and header and len(lines) == 2:
            lines.append("| " + " | ".join([""] * len(col_labels)) + " |")
        return "\n".join(lines)

    @staticmethod
    def _slug(s: str) -> str:
        s = s.strip()
        s = re.sub(r"[^\w\-]+", "-", s)
        s = re.sub(r"-{2,}", "-", s).strip("-")
        return s or "sheet"

    def _compose_filename(self, sheet_name: str) -> str:
        if self.options.safe_filenames:
            wb = self._slug(self.excel_path.stem)
            sh = self._slug(sheet_name)
            return f"{wb}_{sh}.md"
        else:
            return f"{self.excel_path.stem}_{sheet_name}.md"

    def _configure_logger(self):
        if not self.logger.handlers:
            handler = logging.StreamHandler()
            fmt = logging.Formatter(
                "%(asctime)s | %(name)s | %(levelname)s | %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S",
            )
            handler.setFormatter(fmt)
            self.logger.addHandler(handler)
        self.logger.setLevel(self.options.log_level)