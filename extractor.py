"""
Excel extraction logic.
Handles reading large .xlsx files (up to 200k rows) and filtering by keywords.
"""

import pandas as pd
from typing import Callable, Optional


def load_headers(filepath: str) -> list[str]:
    """Read only the header row from an Excel file."""
    df = pd.read_excel(filepath, nrows=0, engine="openpyxl")
    return list(df.columns)


def extract_rows(
    filepath: str,
    column: str,
    keywords: list[str],
    progress_callback: Optional[Callable[[int, int], None]] = None,
    chunk_size: int = 10000,
) -> pd.DataFrame:
    """
    Extract rows where the specified column contains any of the given keywords.

    - Partial match (substring search)
    - OR search across keywords
    - Skips empty cells and cells containing only digits

    Args:
        filepath: Path to the .xlsx file.
        column: Column name to search in.
        keywords: List of keywords (OR condition).
        progress_callback: Called with (rows_processed, total_rows).
        chunk_size: Number of rows per progress update batch.

    Returns:
        DataFrame of matched rows with the original column structure.
    """
    if not keywords:
        raise ValueError("At least one keyword is required.")

    df = pd.read_excel(filepath, engine="openpyxl", dtype=str)
    total = len(df)

    if progress_callback:
        progress_callback(0, total)

    col_data = df[column]

    # Skip empty and digit-only cells
    is_valid = col_data.notna() & col_data.str.strip().ne("") & \
               ~col_data.str.strip().str.fullmatch(r"\d+")

    # OR match across all keywords (partial match)
    pattern = "|".join(map(_escape_keyword, keywords))
    is_match = col_data.str.contains(pattern, na=False, regex=True)

    if progress_callback:
        progress_callback(total, total)

    return df[is_valid & is_match].reset_index(drop=True)


def get_total_rows(filepath: str) -> int:
    """Return the total number of data rows (excluding header)."""
    # openpyxl load_workbook is faster for just getting the row count
    import openpyxl
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active
    count = ws.max_row - 1  # subtract header
    wb.close()
    return max(count, 0)


def save_result(df: pd.DataFrame, output_path: str) -> None:
    """Save the result DataFrame to an .xlsx file."""
    df.to_excel(output_path, index=False, engine="openpyxl")


def _escape_keyword(kw: str) -> str:
    """Escape special regex characters in a keyword."""
    import re
    return re.escape(kw.strip())
