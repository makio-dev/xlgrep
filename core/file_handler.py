"""
ファイル操作モジュール
フォルダー再帰検索、CSVからのファイルリスト読み込み、テキストパースを提供する。
"""
import csv
import re
import sys
from pathlib import Path
from typing import List, Optional


EXCEL_EXTENSIONS = {".xlsx", ".xls"}


def collect_excel_files_from_folder(folder_path: str) -> List[Path]:
    """
    指定フォルダー配下の全Excelファイルを再帰的に収集する。

    Args:
        folder_path: 検索対象フォルダーパス

    Returns:
        Excelファイルのパスリスト

    Raises:
        ValueError: フォルダーが存在しない場合
    """
    path = Path(folder_path)
    if not path.exists():
        raise ValueError(f"フォルダーが存在しません: {folder_path}")
    if not path.is_dir():
        raise ValueError(f"フォルダーパスが正しくありません: {folder_path}")

    files = []
    for ext in EXCEL_EXTENSIONS:
        files.extend(path.rglob(f"*{ext}"))

    # 重複除外・ソート
    return sorted(set(files))


def collect_excel_files_from_csv(csv_path: str) -> List[Path]:
    """
    CSVファイルからExcelファイルパスリストを読み込む。
    CSVは 'filepath' 列を持つ形式（またはヘッダーなし1列目をパスとして扱う）。

    Args:
        csv_path: CSVファイルパス

    Returns:
        Excelファイルのパスリスト

    Raises:
        ValueError: CSVファイルが存在しない場合
        ValueError: CSVファイルの形式が不正な場合
    """
    path = Path(csv_path)
    if not path.exists():
        raise ValueError(f"CSVファイルが存在しません: {csv_path}")

    files = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames

        if fieldnames and "filepath" in fieldnames:
            # 'filepath' 列から読み込み
            for row in reader:
                raw = row["filepath"].strip().strip('"')
                if raw:
                    files.append(Path(raw))
        elif fieldnames:
            # 最初の列をパスとして扱う
            first_col = fieldnames[0]
            for row in reader:
                raw = row[first_col].strip().strip('"')
                if raw:
                    files.append(Path(raw))
        else:
            raise ValueError("CSVファイルが空または形式が不正です")

    return files


def collect_excel_files_from_text(text_or_path: str, is_file: bool = False) -> List[Path]:
    """
    テキスト（または テキストファイル）からExcelファイルパスリストをパースする。
    各行がダブルクォーテーションで囲まれたフルパスを想定。

    Args:
        text_or_path: テキスト内容またはファイルパス
        is_file: Trueの場合、text_or_pathをファイルパスとして扱う

    Returns:
        Excelファイルのパスリスト
    """
    if is_file:
        path = Path(text_or_path)
        if not path.exists():
            raise ValueError(f"ファイルが存在しません: {text_or_path}")
        with open(path, encoding="utf-8-sig") as f:
            content = f.read()
    else:
        content = text_or_path

    files = []
    for line in content.splitlines():
        line = line.strip()
        if not line:
            continue
        # ダブルクォート除去
        line = line.strip('"')
        if line:
            files.append(Path(line))

    return files


def validate_excel_files(file_list: List[Path], logger=None) -> tuple[List[Path], List[Path]]:
    """
    ファイルリストを検証し、存在するファイルと存在しないファイルに分類する。

    Args:
        file_list: 検証するファイルパスリスト
        logger: ロガーインスタンス（Noneの場合はprintを使用）

    Returns:
        (valid_files, invalid_files) のタプル
    """
    valid = []
    invalid = []

    for f in file_list:
        if not f.exists():
            msg = f"ファイルが存在しません（スキップ）: {f}"
            if logger:
                logger.warning(msg)
            else:
                print(f"⚠ {msg}", file=sys.stderr)
            invalid.append(f)
        elif f.suffix.lower() not in EXCEL_EXTENSIONS:
            msg = f"Excelファイルではありません（スキップ）: {f}"
            if logger:
                logger.warning(msg)
            else:
                print(f"⚠ {msg}", file=sys.stderr)
            invalid.append(f)
        else:
            valid.append(f)

    return valid, invalid
