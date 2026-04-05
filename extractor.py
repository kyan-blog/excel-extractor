"""
Excel extraction logic.
Handles reading large .xlsx files (up to 200k rows) and filtering by keywords.
"""

import pandas as pd
from typing import Callable, Optional


def load_headers(filepath: str, header_row: int = 0) -> list[str]:
    """Read only the header row from an Excel file.

    Args:
        filepath: Path to the .xlsx file.
        header_row: 0-indexed row number of the header (0 = 1行目).
    """
    try:
        df = pd.read_excel(filepath, nrows=0, header=header_row, engine="openpyxl")
        return list(df.columns)
    except Exception as e:
        raise RuntimeError(
            f"ファイルのヘッダー読み込みに失敗しました。\n"
            f"ファイルが壊れているか、対応していない形式の可能性があります。\n\n詳細: {e}"
        ) from e


def get_sample_values(filepath: str, column: str, header_row: int = 0, n: int = 10) -> list[str]:
    """Return the first n non-null values from the specified column (for debugging)."""
    try:
        df = pd.read_excel(filepath, header=header_row, engine="openpyxl", dtype=str)
        values = df[column].dropna().head(n).tolist()
        return [str(v) for v in values]
    except Exception as e:
        raise RuntimeError(f"サンプル値の取得に失敗しました。\n\n詳細: {e}") from e


def extract_rows(
    filepath: str,
    column: str,
    keywords: list[str],
    header_row: int = 0,
    progress_callback: Optional[Callable[[int, int], None]] = None,
    chunk_size: int = 10000,
) -> pd.DataFrame:
    """
    Extract rows where the specified column contains any of the given keywords.

    - Partial match (substring search)
    - OR search across keywords (case-insensitive)
    - Skips empty cells and cells containing only digits

    Args:
        filepath: Path to the .xlsx file.
        column: Column name to search in.
        keywords: List of keywords (OR condition).
        header_row: 0-indexed row number of the header.
        progress_callback: Called with (rows_processed, total_rows).
        chunk_size: Number of rows per progress update batch.

    Returns:
        DataFrame of matched rows with the original column structure.
    """
    if not keywords:
        raise ValueError("キーワードを1つ以上入力してください。")

    try:
        df = pd.read_excel(filepath, header=header_row, engine="openpyxl", dtype=str)
    except MemoryError:
        raise MemoryError(
            "メモリ不足のため、ファイルを読み込めませんでした。\n"
            "他のアプリを閉じてから再試行してください。"
        )
    except Exception as e:
        raise RuntimeError(
            f"ファイルの読み込みに失敗しました。\n"
            f"ファイルが壊れているか、別のアプリで開いている可能性があります。\n\n詳細: {e}"
        ) from e

    total = len(df)

    if progress_callback:
        progress_callback(0, total)

    try:
        # strip whitespace before comparison
        col_data = df[column].str.strip()

        # Skip empty and digit-only cells
        is_valid = col_data.notna() & col_data.ne("") & \
                   ~col_data.str.fullmatch(r"\d+")

        # OR match across all keywords (partial match, case-insensitive)
        pattern = "|".join(map(_escape_keyword, keywords))
        is_match = col_data.str.contains(pattern, na=False, regex=True, case=False)

    except MemoryError:
        raise MemoryError(
            "メモリ不足のため、抽出処理を完了できませんでした。\n"
            "他のアプリを閉じてから再試行してください。"
        )
    except Exception as e:
        raise RuntimeError(f"抽出処理中にエラーが発生しました。\n\n詳細: {e}") from e

    if progress_callback:
        progress_callback(total, total)

    return df[is_valid & is_match].reset_index(drop=True)


def get_total_rows(filepath: str) -> int:
    """Return the total number of data rows (excluding header)."""
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
