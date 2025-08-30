"""
src/xl2md/converter.py

A small utility to convert every sheet in an Excel workbook into
a Markdown table that mirrors the sheet's structure as closely
as pandas represents it.
"""

from __future__ import annotations

import logging
import math
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Sequence

import pandas as pd


# -----------------------------
# Options & Exceptions
# -----------------------------

@dataclass
class ConverterOptions:
    out_dir: str = "./markdown_sheets"
    include_index: bool = False
    index_label: Optional[str] = None
    header: bool = True                     # include a header row in markdown from df.columns
    skip_empty_sheets: bool = True          # skip sheets that are empty after read
    engine: str = "openpyxl"                # Excel reader engine
    safe_filenames: bool = True             # slugify sheet names for filenames
    title_prefix: str = ""                  # optional prefix in the H1 title (e.g., workbook name)
    log_level: int = logging.INFO           # default log level
    overwrite: bool = True                  # overwrite existing .md files
    # Advanced: if set, only convert sheets whose names match any of these regex patterns
    sheet_name_allowlist: Sequence[str] = field(default_factory=list)
    # Advanced: skip sheets whose names match any of these regex patterns
    sheet_name_blocklist: Sequence[str] = field(default_factory=list)


class ExcelToMarkdownError(Exception):
    """Base class for conversion errors."""


class WorkbookReadError(ExcelToMarkdownError):
    """Raised when the Excel file cannot be read."""


class SheetConversionError(ExcelToMarkdownError):
    """Raised when a single sheet fails to convert."""


# -----------------------------
# Converter Class
# -----------------------------

class ExcelToMarkdownConverter:
    def __init__(self, excel_path: str, options: Optional[ConverterOptions] = None, logger: Optional[logging.Logger] = None):
        self.excel_path = Path(excel_path)
        self.options = options or ConverterOptions()
        self.logger = logger or logging.getLogger(self.__class__.__name__)
        self._configure_logger()

        if not self.excel_path.exists():
            raise WorkbookReadError(f"Excel file not found: {self.excel_path}")

        # If no title_prefix provided, default to workbook name for nicer H1 titles
        if not self.options.title_prefix:
            self.options.title_prefix = self.excel_path.stem

        self.logger.debug(
            "Initialized ExcelToMarkdownConverter with path=%s, options=%s",
            self.excel_path, self.options
        )

    # --------- Public API ---------

    def convert(self) -> List[str]:
        """
        Convert all sheets in the Excel file to Markdown files.
        Returns a list of written file paths.
        """
        self.logger.info("Starting conversion for: %s", self.excel_path)
        xls = self._open_workbook()
        out_dir = Path(self.options.out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)

        written_paths: List[str] = []

        for sheet_name in xls.sheet_names:
            if not self._sheet_allowed(sheet_name):
                self.logger.debug("Skipping sheet due to allow/block rules: %s", sheet_name)
                continue

            try:
                # header=0 is typical; if your file uses None, adjust as needed
                df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
            except Exception as e:
                self.logger.exception("Failed reading sheet '%s'", sheet_name)
                raise SheetConversionError(f"Failed reading sheet '{sheet_name}': {e}") from e

            if self.options.skip_empty_sheets and df.empty:
                self.logger.info("Skipping empty sheet: %s", sheet_name)
                continue

            md_text = self._sheet_to_markdown(sheet_name, df)

            # Save as workbookname_sheetname.md (slugified if safe_filenames = True)
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

    # --------- Internals ---------

    def _open_workbook(self) -> pd.ExcelFile:
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

    def _sheet_allowed(self, sheet_name: str) -> bool:
        # Blocklist check
        for pat in self.options.sheet_name_blocklist:
            if re.search(pat, sheet_name, flags=re.I):
                return False
        # Allowlist check (if provided, require a match)
        if self.options.sheet_name_allowlist:
            for pat in self.options.sheet_name_allowlist:
                if re.search(pat, sheet_name, flags=re.I):
                    return True
            return False
        return True

    def _sheet_to_markdown(self, sheet_name: str, df: pd.DataFrame) -> str:
        self.logger.debug("Converting sheet to markdown: %s (shape=%s)", sheet_name, df.shape)

        # Title: "<workbook> — <sheet>"
        title = f"# {self.options.title_prefix} — {sheet_name}\n\n"

        # Table
        table_md = self._df_to_markdown_table(
            df,
            include_index=self.options.include_index,
            index_label=self.options.index_label,
            header=self.options.header
        )

        return title + table_md + "\n"

    # --- Rendering helpers (structure-preserving) ---

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
        # Preserve line breaks inside Markdown table cells
        s = s.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "<br>")
        # Escape pipes and backslashes so tables render correctly
        s = s.replace("\\", "\\\\").replace("|", r"\|")
        return s.strip()

    @classmethod
    def _clean_header(cls, col) -> str:
        # Render 'Unnamed: n' like a blank header (as Excel visually shows)
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

        # Header labels
        col_labels: List[str] = []
        if include_index:
            col_labels.append(cls._to_str(index_label) if index_label is not None else "")
        for c in df.columns:
            col_labels.append(cls._clean_header(c))

        # Guarantee at least one column for Markdown table validity
        if len(col_labels) == 0:
            col_labels = [""]

        lines: List[str] = []

        if header:
            lines.append("| " + " | ".join(col_labels) + " |")
            lines.append("| " + " | ".join(["---"] * len(col_labels)) + " |")

        # Body rows
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

        # Edge case: no data rows
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
        """
        Compose output filename as 'workbookname_sheetname.md'.
        Uses slugified names if safe_filenames=True.
        """
        if self.options.safe_filenames:
            wb = self._slug(self.excel_path.stem)
            sh = self._slug(sheet_name)
            return f"{wb}_{sh}.md"
        else:
            return f"{self.excel_path.stem}_{sheet_name}.md"

    def _configure_logger(self):
        # Only configure if the logger has no handlers (avoid duplicate logs when reused)
        if not self.logger.handlers:
            handler = logging.StreamHandler()
            fmt = logging.Formatter(
                "%(asctime)s | %(name)s | %(levelname)s | %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S",
            )
            handler.setFormatter(fmt)
            self.logger.addHandler(handler)
        self.logger.setLevel(self.options.log_level)
