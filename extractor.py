"""
Excel extraction logic.
Handles reading large .xlsx files (up to 200k rows) and filtering by keywords.
"""

import pandas as pd
from typing import Callable


def load_headers(filepath: str) -> list[str]:
    """Read only the header row from an Excel file."""
    df = pd.read_excel(filepath, nrows=0, engine="openpyxl")
    return list(df.columns)


def extract_rows(
    filepath: str,
    column: str,
    keywords: list[str],
    progress_callback: Callable[[int, int], None] | None = None,
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
        chunk_size: Number of rows to read per chunk.

    Returns:
        DataFrame of matched rows with the original column structure.
    """
    if not keywords:
        raise ValueError("At least one keyword is required.")

    # Get total row count for progress reporting (subtract 1 for header)
    total_rows = sum(1 for _ in pd.read_excel(
        filepath, usecols=[column], engine="openpyxl", chunksize=chunk_size
    ))
    # Re-read with actual chunk iteration
    reader = pd.read_excel(
        filepath, engine="openpyxl", chunksize=chunk_size
    )

    results: list[pd.DataFrame] = []
    processed = 0

    for chunk in reader:
        processed += len(chunk)

        col_data = chunk[column]

        # Skip empty and digit-only cells
        is_valid = col_data.notna() & col_data.astype(str).str.strip().ne("") & \
                   ~col_data.astype(str).str.strip().str.fullmatch(r"\d+")

        # OR match across all keywords (partial match)
        pattern = "|".join(map(_escape_keyword, keywords))
        is_match = col_data.astype(str).str.contains(pattern, na=False, regex=True)

        matched = chunk[is_valid & is_match]
        if not matched.empty:
            results.append(matched)

        if progress_callback:
            progress_callback(processed, processed)  # total unknown at chunk level

    if not results:
        return pd.DataFrame(columns=pd.read_excel(filepath, nrows=0, engine="openpyxl").columns)

    return pd.concat(results, ignore_index=True)


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
