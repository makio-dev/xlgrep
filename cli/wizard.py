"""
ウィザードモジュール（tkinter GUI版）
ステップ式のGUIウィザードでExcelファイル検索を実行する。
"""
import re
import shutil
import sys
import threading
import time
from pathlib import Path
from tkinter import (
    BooleanVar, END, IntVar, StringVar, Text, Tk, Toplevel,
    filedialog, messagebox, ttk
)
import tkinter as tk
from typing import List, Optional

from core.file_handler import (
    collect_excel_files_from_csv,
    collect_excel_files_from_folder,
    collect_excel_files_from_text,
    validate_excel_files,
)
from core.logger import CallbackLogger
from core.searcher import ExcelSearcher
from core.exporter import export_results


# ─── カラーパレット ────────────────────────────────────────────────────────
COLORS = {
    "bg": "#1e1e2e",
    "surface": "#2a2a3e",
    "surface2": "#313145",
    "accent": "#7c6af7",
    "accent_hover": "#9b8cf9",
    "success": "#50fa7b",
    "warning": "#f1fa8c",
    "error": "#ff5555",
    "text": "#cdd6f4",
    "text_dim": "#7f849c",
    "border": "#3d3d56",
    "entry_bg": "#1a1a2a",
    "btn_bg": "#7c6af7",
    "btn_fg": "#ffffff",
    "btn_hover": "#9b8cf9",
}

FONT_FAMILY = "Helvetica"


# ─── カスタムウィジェット ──────────────────────────────────────────────────

class StyledButton(tk.Frame):
    """macOS対応・視認性重視のクリッカブルボタン（Frame+Label実装）"""

    def __init__(self, parent, text, command=None, style="primary", **kwargs):
        self._command = command
        self._style = style

        if style == "primary":
            self._bg       = COLORS["accent"]
            self._hover_bg = COLORS["accent_hover"]
            self._fg       = "#ffffff"
            border_color   = COLORS["accent_hover"]
        else:
            self._bg       = COLORS["surface2"]
            self._hover_bg = COLORS["border"]
            self._fg       = COLORS["text"]
            border_color   = COLORS["border"]

        # 1px ボーダー用の外枠フレーム
        super().__init__(parent, bg=border_color, padx=1, pady=1,
                         cursor="hand2")

        font = kwargs.get("font", (FONT_FAMILY, 11, "bold"))
        padx = kwargs.get("padx", 18)
        pady = kwargs.get("pady", 8)

        self._label = tk.Label(
            self, text=text,
            font=font,
            bg=self._bg, fg=self._fg,
            padx=padx, pady=pady,
            cursor="hand2",
        )
        self._label.pack(fill="both", expand=True)

        # イベントバインド
        for w in (self, self._label):
            w.bind("<Enter>",       self._on_enter)
            w.bind("<Leave>",       self._on_leave)
            w.bind("<Button-1>",    self._on_click)
            w.bind("<ButtonRelease-1>", self._on_release)

        self._enabled = True

    def _on_enter(self, _=None):
        if self._enabled:
            self._label.config(bg=self._hover_bg)

    def _on_leave(self, _=None):
        if self._enabled:
            self._label.config(bg=self._bg)

    def _on_click(self, _=None):
        if self._enabled:
            self._label.config(bg=COLORS["accent"] if self._style == "primary" else COLORS["surface"])

    def _on_release(self, _=None):
        if self._enabled:
            self._label.config(bg=self._hover_bg)
            if self._command:
                self._command()

    def config(self, **kwargs):
        state = kwargs.pop("state", None)
        text  = kwargs.pop("text", None)
        if state == "disabled":
            self._enabled = False
            self._label.config(bg=COLORS["border"], fg=COLORS["text_dim"])
        elif state == "normal":
            self._enabled = True
            self._label.config(bg=self._bg, fg=self._fg)
        if text is not None:
            self._label.config(text=text)
        if kwargs:
            super().config(**kwargs)

    # pack/grid/place の state= に対応するため
    configure = config


class StyledEntry(tk.Entry):
    """スタイル付きテキスト入力"""
    def __init__(self, parent, **kwargs):
        kwargs.setdefault("font", (FONT_FAMILY, 11))
        kwargs.setdefault("bg", COLORS["entry_bg"])
        kwargs.setdefault("fg", COLORS["text"])
        kwargs.setdefault("insertbackground", COLORS["text"])
        kwargs.setdefault("relief", "flat")
        kwargs.setdefault("bd", 0)
        super().__init__(parent, **kwargs)
        # ボーダー風の枠
        self.config(highlightthickness=1, highlightbackground=COLORS["border"],
                    highlightcolor=COLORS["accent"])


class SectionLabel(tk.Label):
    """セクション見出しラベル"""
    def __init__(self, parent, text, **kwargs):
        kwargs.setdefault("font", (FONT_FAMILY, 12, "bold"))
        kwargs.setdefault("fg", COLORS["accent"])
        kwargs.setdefault("bg", COLORS["bg"])
        kwargs.setdefault("anchor", "w")
        super().__init__(parent, text=text, **kwargs)


class HintLabel(tk.Label):
    """補足説明ラベル"""
    def __init__(self, parent, text, **kwargs):
        kwargs.setdefault("font", (FONT_FAMILY, 9))
        kwargs.setdefault("fg", COLORS["text_dim"])
        kwargs.setdefault("bg", COLORS["bg"])
        kwargs.setdefault("anchor", "w")
        super().__init__(parent, text=text, **kwargs)


# ─── ウィザードアプリ本体 ──────────────────────────────────────────────────

class ExcelGrepWizard(tk.Tk):
    """Excel Grep ウィザード GUI メインウィンドウ"""

    STEPS = [
        "① 検索対象",
        "② キーワード",
        "③ オプション",
        "④ 検索実行",
        "⑤ エクスポート",
    ]

    def __init__(self):
        super().__init__()
        self.title("Excel Grep Tool")
        self.geometry("720x580")
        self.minsize(680, 520)
        self.configure(bg=COLORS["bg"])
        self.resizable(True, True)
        try:
            self.tk.call("tk", "scaling", 1.3)
        except Exception:
            pass

        # 状態変数
        self.current_step = 0
        self.search_mode = StringVar(value="folder")   # "folder" / "filelist"
        self.folder_path = StringVar()
        self.filelist_mode = StringVar(value="csv")    # "csv" / "text"
        self.csv_path = StringVar()
        self.text_paths = None
        self.use_regex = BooleanVar(value=False)
        self.keyword_vars: List[StringVar] = [StringVar() for _ in range(10)]
        self.output_path = StringVar()
        self.output_format = StringVar(value="csv")
        self.file_list: List[Path] = []
        self.search_results = None

        self._build_ui()
        self._show_step(0)

    # ─── UI構築 ──────────────────────────────────────────────────────────

    def _build_ui(self):
        """UIを組み立てる"""
        # ─ タイトルバー
        title_frame = tk.Frame(self, bg=COLORS["surface"], pady=14)
        title_frame.pack(fill="x")
        tk.Label(
            title_frame, text="🔍  Excel Grep Tool",
            font=(FONT_FAMILY, 16, "bold"),
            fg=COLORS["text"], bg=COLORS["surface"],
        ).pack(side="left", padx=24)

        # ─ ステップインジケーター
        self.step_frame = tk.Frame(self, bg=COLORS["bg"], pady=8)
        self.step_frame.pack(fill="x", padx=24)
        self.step_labels = []
        for i, name in enumerate(self.STEPS):
            lbl = tk.Label(
                self.step_frame, text=name,
                font=(FONT_FAMILY, 9),
                fg=COLORS["text_dim"], bg=COLORS["bg"],
                padx=4,
            )
            lbl.pack(side="left")
            self.step_labels.append(lbl)
            if i < len(self.STEPS) - 1:
                tk.Label(self.step_frame, text=" › ", fg=COLORS["border"],
                         bg=COLORS["bg"], font=(FONT_FAMILY, 9)).pack(side="left")

        # 区切り線
        tk.Frame(self, bg=COLORS["border"], height=1).pack(fill="x")

        # ─ コンテンツエリア（各ステップのフレームを切り替え）
        self.content = tk.Frame(self, bg=COLORS["bg"])
        self.content.pack(fill="both", expand=True, padx=28, pady=16)


        # ─ ナビゲーションバー（上部に区切り線を追加してボタンを目立たせる）
        tk.Frame(self, bg=COLORS["accent"], height=2).pack(fill="x", side="bottom")
        nav_frame = tk.Frame(self, bg="#111120", pady=12)
        nav_frame.pack(fill="x", side="bottom")
        self.btn_back = StyledButton(nav_frame, "← 戻る", command=self._go_back, style="secondary")
        self.btn_back.pack(side="left", padx=20, pady=2)
        self.btn_next = StyledButton(nav_frame, "次へ →", command=self._go_next)
        self.btn_next.pack(side="right", padx=20, pady=2)


        # ─ 各ステップのフレームを生成
        self.step_frames = [
            self._build_step_target(),
            self._build_step_keywords(),
            self._build_step_options(),
            self._build_step_search(),
            self._build_step_export(),
        ]

    # ─── ステップ0: 検索対象 ─────────────────────────────────────────────

    def _build_step_target(self) -> tk.Frame:
        f = tk.Frame(self.content, bg=COLORS["bg"])

        SectionLabel(f, "検索対象の選択").pack(fill="x", pady=(0, 12))

        # モード選択
        mode_frame = tk.Frame(f, bg=COLORS["surface"], pady=12, padx=16)
        mode_frame.pack(fill="x", pady=(0, 16))

        for val, label, hint in [
            ("folder", "フォルダーを指定する", "指定フォルダー配下のExcelを再帰的に検索"),
            ("filelist", "ファイルリストを指定する", "CSVまたはテキストでファイルパスを列挙"),
        ]:
            rb = tk.Radiobutton(
                mode_frame, text=f"  {label}", variable=self.search_mode, value=val,
                font=(FONT_FAMILY, 11), fg=COLORS["text"], bg=COLORS["surface"],
                selectcolor=COLORS["accent"], activebackground=COLORS["surface"],
                activeforeground=COLORS["text"], command=self._on_mode_change,
            )
            rb.pack(anchor="w")
            tk.Label(mode_frame, text=f"     {hint}", font=(FONT_FAMILY, 9),
                     fg=COLORS["text_dim"], bg=COLORS["surface"]).pack(anchor="w")
            tk.Frame(mode_frame, bg=COLORS["border"], height=1).pack(fill="x", pady=6)

        # ─ フォルダー入力エリア
        self.folder_area = tk.Frame(f, bg=COLORS["bg"])
        self.folder_area.pack(fill="x")
        SectionLabel(self.folder_area, "フォルダーパス").pack(fill="x")
        folder_row = tk.Frame(self.folder_area, bg=COLORS["bg"])
        folder_row.pack(fill="x", pady=4)
        self.folder_entry = StyledEntry(folder_row, textvariable=self.folder_path)
        self.folder_entry.pack(side="left", fill="x", expand=True)
        StyledButton(folder_row, "参照…", command=self._browse_folder, style="secondary").pack(side="left", padx=(6, 0))

        # ─ ファイルリスト入力エリア
        self.filelist_area = tk.Frame(f, bg=COLORS["bg"])

        # サブモード選択
        submode_frame = tk.Frame(self.filelist_area, bg=COLORS["bg"])
        submode_frame.pack(fill="x", pady=(0, 8))
        SectionLabel(submode_frame, "入力方法").pack(fill="x")
        for val, label in [("csv", "CSVファイルを指定（推奨）"), ("text", "パスをテキストで直接入力")]:
            tk.Radiobutton(
                submode_frame, text=f"  {label}", variable=self.filelist_mode, value=val,
                font=(FONT_FAMILY, 10), fg=COLORS["text"], bg=COLORS["bg"],
                selectcolor=COLORS["accent"], activebackground=COLORS["bg"],
                command=self._on_filelist_mode_change,
            ).pack(anchor="w")

        # CSV選択
        self.csv_area = tk.Frame(self.filelist_area, bg=COLORS["bg"])
        self.csv_area.pack(fill="x", pady=4)
        SectionLabel(self.csv_area, "CSVファイルパス").pack(fill="x")
        csv_row = tk.Frame(self.csv_area, bg=COLORS["bg"])
        csv_row.pack(fill="x")
        self.csv_entry = StyledEntry(csv_row, textvariable=self.csv_path)
        self.csv_entry.pack(side="left", fill="x", expand=True)
        StyledButton(csv_row, "参照…", command=self._browse_csv, style="secondary").pack(side="left", padx=(6, 0))
        HintLabel(self.csv_area, "CSVは 'filepath' 列にファイルパスを1行ずつ記載してください").pack(anchor="w", pady=(2, 0))
        StyledButton(self.csv_area, "CSVひな形をダウンロード", command=self._download_template, style="secondary").pack(anchor="w", pady=(6, 0))

        # テキスト入力
        self.text_area_frame = tk.Frame(self.filelist_area, bg=COLORS["bg"])
        SectionLabel(self.text_area_frame, "ファイルパスを1行ずつ入力").pack(fill="x")
        HintLabel(self.text_area_frame, 'ダブルクォートで囲まれていても問題ありません').pack(anchor="w")
        self.text_import = Text(
            self.text_area_frame, height=6,
            bg=COLORS["entry_bg"], fg=COLORS["text"], insertbackground=COLORS["text"],
            font=(FONT_FAMILY, 10), relief="flat", bd=0,
            highlightthickness=1, highlightbackground=COLORS["border"],
            highlightcolor=COLORS["accent"],
        )
        self.text_import.pack(fill="both", expand=True, pady=4)

        self._on_mode_change()
        return f

    def _on_mode_change(self, *_):
        mode = self.search_mode.get()
        if mode == "folder":
            self.filelist_area.pack_forget()
            self.folder_area.pack(fill="x")
        else:
            self.folder_area.pack_forget()
            self.filelist_area.pack(fill="both", expand=True)
        self._on_filelist_mode_change()

    def _on_filelist_mode_change(self, *_):
        submode = self.filelist_mode.get()
        if submode == "csv":
            self.text_area_frame.pack_forget()
            self.csv_area.pack(fill="x", pady=4)
        else:
            self.csv_area.pack_forget()
            self.text_area_frame.pack(fill="both", expand=True, pady=4)

    def _browse_folder(self):
        path = filedialog.askdirectory(title="検索対象フォルダーを選択")
        if path:
            self.folder_path.set(path)

    def _browse_csv(self):
        path = filedialog.askopenfilename(
            title="CSVファイルを選択",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            self.csv_path.set(path)

    def _download_template(self):
        save_path = filedialog.asksaveasfilename(
            title="CSVひな形の保存先",
            defaultextension=".csv",
            initialfile="filelist_template.csv",
            filetypes=[("CSV files", "*.csv")],
        )
        if not save_path:
            return
        template_src = Path(__file__).parent.parent / "templates" / "filelist_template.csv"
        try:
            if template_src.exists():
                shutil.copy(template_src, save_path)
            else:
                with open(save_path, "w", encoding="utf-8-sig", newline="") as f:
                    f.write("filepath\n")
                    f.write('"C:\\path\\to\\file1.xlsx"\n')
                    f.write('"C:\\path\\to\\file2.xlsx"\n')
            messagebox.showinfo("完了", f"CSVひな形を保存しました:\n{save_path}")
        except Exception as e:
            messagebox.showerror("エラー", f"保存に失敗しました:\n{e}")

    # ─── ステップ1: キーワード ────────────────────────────────────────────

    def _build_step_keywords(self) -> tk.Frame:
        f = tk.Frame(self.content, bg=COLORS["bg"])
        SectionLabel(f, "検索キーワードの入力（最大10個）").pack(fill="x", pady=(0, 4))
        HintLabel(f, "空白のフィールドはスキップされます").pack(fill="x", pady=(0, 12))

        scroll_frame = tk.Frame(f, bg=COLORS["bg"])
        scroll_frame.pack(fill="both", expand=True)

        for i, var in enumerate(self.keyword_vars):
            row = tk.Frame(scroll_frame, bg=COLORS["bg"])
            row.pack(fill="x", pady=3)
            tk.Label(
                row, text=f"{i+1:>2}.", font=(FONT_FAMILY, 11),
                fg=COLORS["text_dim"], bg=COLORS["bg"], width=3,
            ).pack(side="left")
            entry = StyledEntry(row, textvariable=var)
            entry.pack(side="left", fill="x", expand=True)

        return f

    # ─── ステップ2: オプション ────────────────────────────────────────────

    def _build_step_options(self) -> tk.Frame:
        f = tk.Frame(self.content, bg=COLORS["bg"])

        SectionLabel(f, "検索オプション").pack(fill="x", pady=(0, 12))

        # ─ 正規表現
        regex_box = tk.Frame(f, bg=COLORS["surface"], padx=16, pady=12)
        regex_box.pack(fill="x")
        tk.Label(regex_box, text="正規表現モード",
                 font=(FONT_FAMILY, 11, "bold"), fg=COLORS["text"], bg=COLORS["surface"]).pack(anchor="w")
        tk.Frame(regex_box, bg=COLORS["border"], height=1).pack(fill="x", pady=6)
        for val, label, hint in [
            (False, "通常の文字列検索（推奨）", "入力したキーワードをそのまま文字列として検索します"),
            (True,  "正規表現を使用する",       "メタ文字（^ $ . * + ? など）が有効になります"),
        ]:
            row = tk.Frame(regex_box, bg=COLORS["surface"])
            row.pack(fill="x", pady=2)
            tk.Radiobutton(
                row, text=f"  {label}", variable=self.use_regex, value=val,
                font=(FONT_FAMILY, 11), fg=COLORS["text"], bg=COLORS["surface"],
                selectcolor=COLORS["accent"], activebackground=COLORS["surface"],
            ).pack(anchor="w")
            tk.Label(row, text=f"     {hint}", font=(FONT_FAMILY, 9),
                     fg=COLORS["text_dim"], bg=COLORS["surface"]).pack(anchor="w")

        # ─ 置換設定
        tk.Frame(f, bg=COLORS["border"], height=1).pack(fill="x", pady=10)
        SectionLabel(f, "置換設定（オプション）").pack(fill="x")

        # 置換を行うかのトグル
        self.do_replace = BooleanVar(value=False)
        toggle_frame = tk.Frame(f, bg=COLORS["bg"])
        toggle_frame.pack(fill="x", pady=(4, 0))
        tk.Checkbutton(
            toggle_frame, text="  置換を行う（検索マッチしたセルを書き換える）",
            variable=self.do_replace, command=self._on_replace_toggle,
            font=(FONT_FAMILY, 11), fg=COLORS["text"], bg=COLORS["bg"],
            selectcolor=COLORS["accent"], activebackground=COLORS["bg"],
        ).pack(anchor="w")

        # 置換設定エリア（トグルON時に表示）
        self.replace_config_frame = tk.Frame(f, bg=COLORS["surface"], padx=14, pady=10)

        # 置換文字列入力欄（キーワードと対応）
        SectionLabel(self.replace_config_frame, "置換文字列").pack(fill="x")
        HintLabel(self.replace_config_frame,
                  "各キーワードに対応する置換後の文字列を入力してください").pack(anchor="w", pady=(0, 6))
        self.replacement_frame = tk.Frame(self.replace_config_frame, bg=COLORS["surface"])
        self.replacement_frame.pack(fill="x")
        self.replacement_vars: List[StringVar] = [StringVar() for _ in range(10)]
        self._replacement_rows: list = []

        # バックアップ・ドライラン
        opt_row = tk.Frame(self.replace_config_frame, bg=COLORS["surface"])
        opt_row.pack(fill="x", pady=(10, 0))
        self.backup_var = BooleanVar(value=True)
        self.dry_run_var = BooleanVar(value=False)
        tk.Checkbutton(
            opt_row, text="  置換前にバックアップを作成（.bak）",
            variable=self.backup_var,
            font=(FONT_FAMILY, 10), fg=COLORS["text"], bg=COLORS["surface"],
            selectcolor=COLORS["accent"], activebackground=COLORS["surface"],
        ).pack(anchor="w")
        tk.Checkbutton(
            opt_row, text="  ドライラン（プレビューのみ・ファイルを変更しない）",
            variable=self.dry_run_var,
            font=(FONT_FAMILY, 10), fg=COLORS["warning"], bg=COLORS["surface"],
            selectcolor=COLORS["accent"], activebackground=COLORS["surface"],
        ).pack(anchor="w")

        return f

    def _on_replace_toggle(self):
        """置換トグルのON/OFFで設定エリアを表示/非表示"""
        if self.do_replace.get():
            self.replace_config_frame.pack(fill="x", pady=(6, 0))
            self._refresh_replacement_rows()
        else:
            self.replace_config_frame.pack_forget()

    def _refresh_replacement_rows(self):
        """入力済みキーワードに合わせて置換文字列入力欄を更新"""
        for w in self._replacement_rows:
            w.destroy()
        self._replacement_rows.clear()

        keywords = [v.get().strip() for v in self.keyword_vars if v.get().strip()]
        for i, kw in enumerate(keywords):
            row = tk.Frame(self.replacement_frame, bg=COLORS["surface"])
            row.pack(fill="x", pady=2)
            tk.Label(
                row, text=f'  "{kw}"  →',
                font=(FONT_FAMILY, 10), fg=COLORS["text_dim"], bg=COLORS["surface"],
                width=20, anchor="w",
            ).pack(side="left")
            entry = StyledEntry(row, textvariable=self.replacement_vars[i])
            entry.pack(side="left", fill="x", expand=True)
            self._replacement_rows.append(row)



    # ─── ステップ3: 検索実行 ──────────────────────────────────────────────

    def _build_step_search(self) -> tk.Frame:
        f = tk.Frame(self.content, bg=COLORS["bg"])

        # サマリー
        self.search_summary_var = StringVar(value="")
        tk.Label(
            f, textvariable=self.search_summary_var,
            font=(FONT_FAMILY, 10), fg=COLORS["text_dim"], bg=COLORS["bg"],
            anchor="w", justify="left", wraplength=620,
        ).pack(fill="x", pady=(0, 10))

        # プログレスバー
        self.progress_var = tk.DoubleVar(value=0)
        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Accent.Horizontal.TProgressbar",
            troughcolor=COLORS["surface"],
            background=COLORS["accent"],
            thickness=14,
        )
        self.progress_bar = ttk.Progressbar(
            f, variable=self.progress_var, maximum=100,
            style="Accent.Horizontal.TProgressbar",
        )
        self.progress_bar.pack(fill="x", pady=(0, 4))
        self.progress_label = tk.Label(
            f, text="", font=(FONT_FAMILY, 9),
            fg=COLORS["text_dim"], bg=COLORS["bg"], anchor="w",
        )
        self.progress_label.pack(fill="x", pady=(0, 8))

        # ログ表示
        SectionLabel(f, "ログ").pack(fill="x")
        self.log_text = Text(
            f, height=14,
            bg=COLORS["entry_bg"], fg=COLORS["text"],
            font=("Courier", 9), relief="flat", bd=0, state="disabled",
            highlightthickness=1, highlightbackground=COLORS["border"],
        )
        self.log_text.pack(fill="both", expand=True, pady=(4, 0))
        # タグ設定
        self.log_text.tag_config("INFO", foreground=COLORS["text"])
        self.log_text.tag_config("WARNING", foreground=COLORS["warning"])
        self.log_text.tag_config("ERROR", foreground=COLORS["error"])
        self.log_text.tag_config("SUCCESS", foreground=COLORS["success"])

        return f

    # ─── ステップ4: エクスポート ──────────────────────────────────────────

    def _build_step_export(self) -> tk.Frame:
        f = tk.Frame(self.content, bg=COLORS["bg"])

        # 検索結果サマリー
        self.result_summary_var = StringVar(value="")
        tk.Label(
            f, textvariable=self.result_summary_var,
            font=(FONT_FAMILY, 13, "bold"), fg=COLORS["success"], bg=COLORS["bg"],
            anchor="w",
        ).pack(fill="x", pady=(0, 16))

        # エクスポート設定
        SectionLabel(f, "エクスポート設定").pack(fill="x")
        export_box = tk.Frame(f, bg=COLORS["surface"], padx=16, pady=14)
        export_box.pack(fill="x", pady=(4, 0))

        # 形式
        fmt_row = tk.Frame(export_box, bg=COLORS["surface"])
        fmt_row.pack(fill="x", pady=(0, 10))
        tk.Label(fmt_row, text="形式:", font=(FONT_FAMILY, 11), fg=COLORS["text"],
                 bg=COLORS["surface"], width=8, anchor="w").pack(side="left")
        for val, label in [("csv", "CSV"), ("json", "JSON"), ("txt", "TXT")]:
            tk.Radiobutton(
                fmt_row, text=label, variable=self.output_format, value=val,
                font=(FONT_FAMILY, 11), fg=COLORS["text"], bg=COLORS["surface"],
                selectcolor=COLORS["accent"], activebackground=COLORS["surface"],
            ).pack(side="left", padx=8)

        # 保存先
        out_row = tk.Frame(export_box, bg=COLORS["surface"])
        out_row.pack(fill="x")
        tk.Label(out_row, text="保存先:", font=(FONT_FAMILY, 11), fg=COLORS["text"],
                 bg=COLORS["surface"], width=8, anchor="w").pack(side="left")
        self.output_entry = StyledEntry(out_row, textvariable=self.output_path)
        self.output_entry.pack(side="left", fill="x", expand=True)
        StyledButton(out_row, "参照…", command=self._browse_output, style="secondary").pack(side="left", padx=(6, 0))

        # エクスポートボタン
        self.export_btn = StyledButton(f, "💾  エクスポート実行", command=self._do_export)
        self.export_btn.pack(pady=16)

        self.export_status = tk.Label(f, text="", font=(FONT_FAMILY, 10),
                                      fg=COLORS["success"], bg=COLORS["bg"])
        self.export_status.pack()

        return f

    def _browse_output(self):
        fmt = self.output_format.get()
        ext_map = {"csv": [("CSV", "*.csv")], "json": [("JSON", "*.json")], "txt": [("Text", "*.txt")]}
        path = filedialog.asksaveasfilename(
            title="保存先を選択",
            defaultextension=f".{fmt}",
            filetypes=ext_map.get(fmt, [("All", "*.*")]),
            initialfile=f"results.{fmt}",
        )
        if path:
            self.output_path.set(path)

    # ─── ナビゲーション ───────────────────────────────────────────────────

    def _show_step(self, step: int):
        for f in self.step_frames:
            f.pack_forget()
        self.step_frames[step].pack(fill="both", expand=True)

        # ステップインジケーター更新
        for i, lbl in enumerate(self.step_labels):
            if i == step:
                lbl.config(fg=COLORS["accent"], font=(FONT_FAMILY, 9, "bold"))
            elif i < step:
                lbl.config(fg=COLORS["success"], font=(FONT_FAMILY, 9))
            else:
                lbl.config(fg=COLORS["text_dim"], font=(FONT_FAMILY, 9))

        self.btn_back.config(state="normal" if step > 0 else "disabled")

        # 最終ステップでは次へボタン無効
        if step == len(self.STEPS) - 1:
            self.btn_next.config(state="disabled")
        else:
            self.btn_next.config(state="normal", text="次へ →")

    def _go_back(self):
        if self.current_step > 0:
            self.current_step -= 1
            self._show_step(self.current_step)

    def _go_next(self):
        if not self._validate_current_step():
            return
        if self.current_step < len(self.STEPS) - 1:
            self.current_step += 1
            self._show_step(self.current_step)
            if self.current_step == 3:  # 検索実行ステップ
                self._start_search()

    def _validate_current_step(self) -> bool:
        step = self.current_step

        if step == 0:
            return self._validate_target()
        elif step == 1:
            return self._validate_keywords()
        elif step == 2:
            return self._validate_options()
        return True

    # ─── バリデーション ───────────────────────────────────────────────────

    def _validate_target(self) -> bool:
        mode = self.search_mode.get()
        if mode == "folder":
            path = self.folder_path.get().strip()
            if not path:
                messagebox.showerror("入力エラー", "フォルダーパスを入力してください")
                return False
            if not Path(path).is_dir():
                messagebox.showerror("入力エラー", f"フォルダーが存在しません:\n{path}")
                return False
        else:
            submode = self.filelist_mode.get()
            if submode == "csv":
                csv = self.csv_path.get().strip()
                if not csv:
                    messagebox.showerror("入力エラー", "CSVファイルパスを入力してください")
                    return False
                if not Path(csv).exists():
                    messagebox.showerror("入力エラー", f"CSVファイルが存在しません:\n{csv}")
                    return False
            else:
                text = self.text_import.get("1.0", END).strip()
                if not text:
                    messagebox.showerror("入力エラー", "ファイルパスを入力してください")
                    return False
        return True

    def _validate_keywords(self) -> bool:
        keywords = [v.get().strip() for v in self.keyword_vars if v.get().strip()]
        if not keywords:
            messagebox.showerror("入力エラー", "キーワードを1個以上入力してください")
            return False
        return True

    def _validate_options(self) -> bool:
        if self.use_regex.get():
            keywords = [v.get().strip() for v in self.keyword_vars if v.get().strip()]
            for kw in keywords:
                try:
                    re.compile(kw)
                except re.error as e:
                    messagebox.showerror(
                        "正規表現エラー",
                        f"不正なパターン: '{kw}'\n\nエラー: {e}\n\n正規表現を無効にするか、パターンを修正してください。"
                    )
                    return False
        return True

    # ─── 検索実行 ─────────────────────────────────────────────────────────

    def _start_search(self):
        """検索を別スレッドで実行する"""
        keywords = [v.get().strip() for v in self.keyword_vars if v.get().strip()]
        use_regex = self.use_regex.get()
        mode_str = "正規表現" if use_regex else "通常検索"

        # ファイルリスト収集
        try:
            mode = self.search_mode.get()
            if mode == "folder":
                raw_files = collect_excel_files_from_folder(self.folder_path.get())
            elif self.filelist_mode.get() == "csv":
                raw_files = collect_excel_files_from_csv(self.csv_path.get())
            else:
                text = self.text_import.get("1.0", END)
                raw_files = collect_excel_files_from_text(text)
        except Exception as e:
            messagebox.showerror("エラー", str(e))
            return

        valid_files, invalid_files = validate_excel_files(raw_files)
        if not valid_files:
            messagebox.showerror("エラー", "有効なExcelファイルが見つかりませんでした")
            return

        self.file_list = valid_files

        # サマリー更新
        self.search_summary_var.set(
            f"ファイル数: {len(valid_files)}件  │  キーワード: {', '.join(keywords)}  │  モード: {mode_str}"
        )
        self.btn_next.config(state="disabled")
        self.btn_back.config(state="disabled")
        self._append_log("INFO", "─" * 50)
        self._append_log("INFO", f"検索開始: {len(valid_files)}ファイル")
        if invalid_files:
            self._append_log("WARNING", f"スキップ: {len(invalid_files)}件（存在しないファイル等）")

        def run():
            log_messages = []

            def log_cb(level, msg):
                self.after(0, lambda l=level, m=msg: self._append_log(l, m))

            logger = CallbackLogger(callback=log_cb, log_dir="logs", quiet=True)
            total = len(valid_files)

            def progress_cb(current, _total, file_path, total_matches):
                pct = current / total * 100
                self.after(0, lambda p=pct, fp=file_path, tm=total_matches: self._update_progress(p, fp, tm))

            searcher = ExcelSearcher(
                keywords=keywords,
                use_regex=use_regex,
                logger=logger,
                progress_callback=progress_cb,
            )

            start = time.time()
            results = searcher.search(valid_files)
            elapsed = time.time() - start

            self.after(0, lambda: self._on_search_done(results, elapsed, logger))

        threading.Thread(target=run, daemon=True).start()

    def _append_log(self, level: str, msg: str):
        """ログテキストウィジェットにメッセージを追記する"""
        import datetime
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] [{level}] {msg}\n"
        self.log_text.config(state="normal")
        self.log_text.insert(END, line, level)
        self.log_text.see(END)
        self.log_text.config(state="disabled")

    def _update_progress(self, pct: float, file_path: str, total_matches: int):
        self.progress_var.set(pct)
        fname = Path(file_path).name if file_path else ""
        self.progress_label.config(
            text=f"{pct:.0f}%  処理中: {fname}  マッチ: {total_matches}件"
        )

    def _on_search_done(self, results, elapsed: float, logger):
        """検索完了後の処理"""
        self.search_results = results
        total_matches = sum(r.match_count for r in results if r and not r.skipped)
        errors = sum(1 for r in results if r and r.skipped)

        self.progress_var.set(100)
        self._append_log("SUCCESS", "─" * 50)
        self._append_log("SUCCESS", f"検索完了: 総マッチ={total_matches}件, エラー={errors}件, 処理時間={elapsed:.1f}秒")
        self._append_log("INFO", f"ログファイル: {logger.get_log_file_path()}")
        self.progress_label.config(text=f"完了  総マッチ: {total_matches}件  処理時間: {elapsed:.1f}秒")

        # ─ 置換処理
        if hasattr(self, "do_replace") and self.do_replace.get():
            keywords = [v.get().strip() for v in self.keyword_vars if v.get().strip()]
            replacements = [self.replacement_vars[i].get() for i in range(len(keywords))]
            replace_map = dict(zip(keywords, replacements))
            dry_run = self.dry_run_var.get()
            backup = self.backup_var.get()
            mode_label = "[DRY-RUN] " if dry_run else ""

            self._append_log("INFO", "─" * 50)
            self._append_log("INFO", f"{mode_label}置換開始: {len(self.file_list)}ファイル")
            for k, v in replace_map.items():
                self._append_log("INFO", f'  "{k}"  →  "{v}"')

            def do_replace_thread():
                from core.replacer import replace_files

                def rep_cb(level, msg):
                    self.after(0, lambda l=level, m=msg: self._append_log(l, m))

                rep_logger = CallbackLogger(callback=rep_cb, log_dir="logs", quiet=True)
                rep_results = replace_files(
                    self.file_list,
                    replace_map,
                    use_regex=self.use_regex.get(),
                    backup=backup,
                    dry_run=dry_run,
                    logger=rep_logger,
                )
                total_replaced = sum(r.replace_count for r in rep_results if not r.skipped)
                rep_errors = sum(1 for r in rep_results if r.skipped)

                def on_done():
                    if dry_run:
                        self._append_log("WARNING", f"[DRY-RUN] {total_replaced}件を置換する予定です（ファイルは変更されていません）")
                        for rr in rep_results:
                            for rec in rr.records:
                                self._append_log("WARNING", f'  [{rec.sheet_name}] {rec.cell_address}: "{rec.before}" → "{rec.after}"')
                    else:
                        self._append_log("SUCCESS", f"置換完了: {total_replaced}件置換, エラー: {rep_errors}件")
                    self._finish_search_step(total_matches, elapsed)

                self.after(0, on_done)

            threading.Thread(target=do_replace_thread, daemon=True).start()
        else:
            self._finish_search_step(total_matches, elapsed)

    def _finish_search_step(self, total_matches: int, elapsed: float):
        """検索・置換の全処理完了後にUIを更新して次のステップへ"""
        # エクスポートステップの結果表示を更新
        replace_label = "  │  置換あり" if hasattr(self, "do_replace") and self.do_replace.get() else ""
        self.result_summary_var.set(
            f"✅  完了  ─  {total_matches}件マッチ  │  処理時間: {elapsed:.1f}秒{replace_label}  │  エクスポートしますか？"
        )

        # デフォルト出力ファイル名
        if not self.output_path.get():
            ts = time.strftime("%Y%m%d_%H%M%S")
            self.output_path.set(f"results_{ts}.{self.output_format.get()}")

        # 次のステップへ
        self.btn_next.config(state="normal", text="次へ →")
        self.btn_back.config(state="normal")
        self.current_step += 1
        self._show_step(self.current_step)


    # ─── エクスポート ─────────────────────────────────────────────────────

    def _do_export(self):
        if not self.search_results:
            messagebox.showerror("エラー", "検索結果がありません")
            return
        out = self.output_path.get().strip()
        if not out:
            messagebox.showerror("エラー", "保存先ファイルを指定してください")
            return
        fmt = self.output_format.get()
        try:
            count = export_results(self.search_results, out, fmt=fmt)
            self.export_status.config(
                text=f"✅  {count}件をエクスポートしました: {out}",
                fg=COLORS["success"],
            )
        except Exception as e:
            self.export_status.config(text=f"❌  エラー: {e}", fg=COLORS["error"])


# ─── エントリポイント ─────────────────────────────────────────────────────

def run_wizard():
    """ウィザードGUIを起動する"""
    app = ExcelGrepWizard()
    app.mainloop()
