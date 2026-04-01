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
        # Show error without a full Tk window
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
        self.geometry("520x400")

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

        # --- Column selection ---
        col_frame = ttk.LabelFrame(self, text="② 検索対象の列を選択", padding=8)
        col_frame.pack(fill="x", **pad)

        self._col_var = tk.StringVar()
        self._col_combo = ttk.Combobox(
            col_frame, textvariable=self._col_var, state="disabled", width=40
        )
        self._col_combo.pack(fill="x")

        # --- Keyword input ---
        kw_frame = ttk.LabelFrame(self, text="③ キーワード入力（カンマ区切りでOR検索）", padding=8)
        kw_frame.pack(fill="x", **pad)

        self._kw_entry = ttk.Entry(kw_frame, width=50)
        self._kw_entry.pack(fill="x")
        ttk.Label(
            kw_frame,
            text="例: ｱｲｳ,ｴｵ　　※部分一致・空欄/数字のみのセルは除外",
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
        self._status_label.config(text="ヘッダーを読み込み中...")
        self.update_idletasks()

        try:
            self._headers = extractor.load_headers(path)
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの読み込みに失敗しました:\n{e}")
            return

        self._col_combo.config(values=self._headers, state="readonly")
        if self._headers:
            self._col_combo.current(0)

        self._run_btn.config(state="normal")
        self._status_label.config(text=f"列数: {len(self._headers)}")
        self._result_label.config(text="")

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
            args=(self._filepath, column, keywords),
            daemon=True,
        ).start()

    def _extraction_worker(self, filepath: str, column: str, keywords: list[str]):
        try:
            df = extractor.extract_rows(
                filepath=filepath,
                column=column,
                keywords=keywords,
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
