"""
Excel Extractor - Main GUI entry point.
Requires Python 3.11+ with tkinter (standard library).
"""

import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


def _check_dependencies() -> None:
    """Verify required libraries are importable before launching the GUI."""
    missing = []
    for lib in ("numpy", "pandas", "openpyxl"):
        try:
            __import__(lib)
        except ImportError:
            missing.append(lib)
    if missing:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "起動エラー",
            "以下のライブラリが見つかりません。アプリを再インストールしてください。\n\n"
            + "\n".join(f"  ・{m}" for m in missing),
        )
        root.destroy()
        raise SystemExit(1)


import extractor


class ExcelExtractorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel 抽出ツール")
        self.resizable(False, False)
        self.geometry("520x460")

        self._filepath: str = ""
        self._headers: list[str] = []
        self._result_df = None

        self._build_ui()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        # --- File selection ---
        file_frame = ttk.LabelFrame(self, text="① Excelファイル選択", padding=8)
        file_frame.pack(fill="x", **pad)

        self._file_label = ttk.Label(file_frame, text="（未選択）", foreground="gray")
        self._file_label.pack(side="left", fill="x", expand=True)

        ttk.Button(file_frame, text="ファイルを開く", command=self._open_file).pack(side="right")

        # --- Header row selection ---
        header_frame = ttk.LabelFrame(self, text="② ヘッダー行の選択", padding=8)
        header_frame.pack(fill="x", **pad)

        ttk.Label(header_frame, text="ヘッダー行は何行目ですか？").pack(side="left")
        self._header_row_var = tk.StringVar(value="1行目")
        header_combo = ttk.Combobox(
            header_frame,
            textvariable=self._header_row_var,
            values=["1行目", "2行目", "3行目", "4行目", "5行目"],
            state="readonly",
            width=8,
        )
        header_combo.pack(side="left", padx=(8, 0))
        header_combo.bind("<<ComboboxSelected>>", self._on_header_row_changed)

        # --- Column selection ---
        col_frame = ttk.LabelFrame(self, text="③ 検索対象の列を選択", padding=8)
        col_frame.pack(fill="x", **pad)

        col_inner = ttk.Frame(col_frame)
        col_inner.pack(fill="x")

        self._col_var = tk.StringVar()
        self._col_combo = ttk.Combobox(
            col_inner, textvariable=self._col_var, state="disabled", width=32
        )
        self._col_combo.pack(side="left", fill="x", expand=True)

        self._debug_btn = ttk.Button(
            col_inner, text="先頭10件を確認", command=self._show_sample, state="disabled"
        )
        self._debug_btn.pack(side="right", padx=(6, 0))

        # --- Keyword input ---
        kw_frame = ttk.LabelFrame(self, text="④ キーワード入力（カンマ区切りでOR検索）", padding=8)
        kw_frame.pack(fill="x", **pad)

        self._kw_entry = ttk.Entry(kw_frame, width=50)
        self._kw_entry.pack(fill="x")
        ttk.Label(
            kw_frame,
            text="例: ｱｲｳ,ｴｵ　　※部分一致・大文字小文字区別なし・空欄/数字のみのセルは除外",
            foreground="gray",
        ).pack(anchor="w")

        # --- Run button ---
        run_frame = ttk.Frame(self)
        run_frame.pack(fill="x", **pad)

        self._run_btn = ttk.Button(
            run_frame, text="抽出実行", command=self._run_extraction, state="disabled"
        )
        self._run_btn.pack(side="right")

        # --- Progress ---
        prog_frame = ttk.Frame(self)
        prog_frame.pack(fill="x", padx=10, pady=2)

        self._progress_var = tk.DoubleVar(value=0)
        self._progress_bar = ttk.Progressbar(
            prog_frame, variable=self._progress_var, mode="indeterminate", length=500
        )
        self._progress_bar.pack(fill="x")

        self._status_label = ttk.Label(self, text="", foreground="gray")
        self._status_label.pack(anchor="w", padx=10)

        # --- Result info ---
        self._result_label = ttk.Label(self, text="", font=("", 10, "bold"))
        self._result_label.pack(anchor="w", padx=10, pady=4)

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _get_header_row(self) -> int:
        """Return the selected header row as a 0-indexed integer."""
        label = self._header_row_var.get()  # e.g. "2行目"
        return int(label.replace("行目", "")) - 1

    def _reload_headers(self):
        """Re-read headers using the currently selected header row."""
        if not self._filepath:
            return
        self._status_label.config(text="ヘッダーを読み込み中...")
        self.update_idletasks()
        try:
            self._headers = extractor.load_headers(self._filepath, header_row=self._get_header_row())
        except Exception as e:
            messagebox.showerror("エラー", str(e))
            return
        self._col_combo.config(values=self._headers, state="readonly")
        if self._headers:
            self._col_combo.current(0)
        self._run_btn.config(state="normal")
        self._debug_btn.config(state="normal")
        self._status_label.config(text=f"列数: {len(self._headers)}")

    # ------------------------------------------------------------------
    # Event handlers
    # ------------------------------------------------------------------

    def _open_file(self):
        path = filedialog.askopenfilename(
            title="Excelファイルを選択",
            filetypes=[("Excel ファイル", "*.xlsx"), ("すべてのファイル", "*.*")],
        )
        if not path:
            return

        self._filepath = path
        short = path if len(path) <= 60 else "…" + path[-57:]
        self._file_label.config(text=short, foreground="black")
        self._result_label.config(text="")
        self._reload_headers()

    def _on_header_row_changed(self, _event=None):
        """Re-load headers when the user changes the header row selection."""
        self._reload_headers()

    def _show_sample(self):
        """Show the first 10 values of the selected column for debugging."""
        column = self._col_var.get()
        if not column:
            messagebox.showwarning("警告", "列を選択してください。")
            return
        try:
            values = extractor.get_sample_values(
                self._filepath, column, header_row=self._get_header_row()
            )
        except Exception as e:
            messagebox.showerror("エラー", str(e))
            return

        if not values:
            messagebox.showinfo("先頭10件", "データが見つかりませんでした。")
            return

        lines = "\n".join(f"  {i+1}: {repr(v)}" for i, v in enumerate(values))
        messagebox.showinfo(
            f"「{column}」列の先頭{len(values)}件",
            f"列名: {column}\n\n{lines}\n\n"
            "※ repr()表示のため、半角スペースや特殊文字も確認できます。",
        )

    def _run_extraction(self):
        if not self._filepath:
            messagebox.showwarning("警告", "ファイルを選択してください。")
            return

        column = self._col_var.get()
        if not column:
            messagebox.showwarning("警告", "検索対象の列を選択してください。")
            return

        raw_kw = self._kw_entry.get().strip()
        if not raw_kw:
            messagebox.showwarning("警告", "キーワードを入力してください。")
            return

        keywords = [k.strip() for k in raw_kw.split(",") if k.strip()]
        if not keywords:
            messagebox.showwarning("警告", "有効なキーワードがありません。")
            return

        self._run_btn.config(state="disabled")
        self._result_label.config(text="")
        self._progress_bar.start(10)
        self._status_label.config(text="抽出中...")

        threading.Thread(
            target=self._extraction_worker,
            args=(self._filepath, column, keywords, self._get_header_row()),
            daemon=True,
        ).start()

    def _extraction_worker(self, filepath: str, column: str, keywords: list[str], header_row: int):
        try:
            df = extractor.extract_rows(
                filepath=filepath,
                column=column,
                keywords=keywords,
                header_row=header_row,
                progress_callback=self._on_progress,
            )
            self._result_df = df
            self.after(0, self._on_extraction_done, len(df))
        except Exception as e:
            self.after(0, self._on_extraction_error, str(e))

    def _on_progress(self, processed: int, _total: int):
        self.after(0, lambda: self._status_label.config(text=f"処理中: {processed:,} 行"))

    def _on_extraction_done(self, count: int):
        self._progress_bar.stop()
        self._progress_var.set(100)
        self._run_btn.config(state="normal")
        self._status_label.config(text="抽出完了")
        self._result_label.config(text=f"抽出件数: {count:,} 行")

        if count == 0:
            messagebox.showinfo("結果", "条件に一致する行が見つかりませんでした。")
            return

        save_path = filedialog.asksaveasfilename(
            title="保存先を選択",
            defaultextension=".xlsx",
            filetypes=[("Excel ファイル", "*.xlsx")],
        )
        if not save_path:
            return

        try:
            extractor.save_result(self._result_df, save_path)
            messagebox.showinfo("保存完了", f"ファイルを保存しました:\n{save_path}")
        except Exception as e:
            messagebox.showerror("保存エラー", f"保存に失敗しました:\n{e}")

    def _on_extraction_error(self, msg: str):
        self._progress_bar.stop()
        self._run_btn.config(state="normal")
        self._status_label.config(text="エラーが発生しました")
        messagebox.showerror("エラー", f"抽出中にエラーが発生しました:\n{msg}")


def main():
    _check_dependencies()
    app = ExcelExtractorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
