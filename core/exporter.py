"""
結果出力モジュール
検索結果をCSV、JSON、TXT形式でエクスポートする。
"""
import csv
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from core.searcher import MatchResult, SearchResult


def _flatten_results(results: List[SearchResult]) -> List[MatchResult]:
    """SearchResultリストからMatchResultのフラットリストを生成する"""
    matches = []
    for r in results:
        if r and not r.skipped:
            matches.extend(r.matches)
    return matches


def export_csv(results: List[SearchResult], output_path: str) -> int:
    """
    検索結果をCSV形式でエクスポートする。

    Args:
        results: SearchResultリスト
        output_path: 出力先ファイルパス

    Returns:
        出力したレコード数
    """
    matches = _flatten_results(results)
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    fieldnames = [
        "file_path", "sheet_name", "cell_address", "row", "col",
        "matched_keyword", "cell_value", "search_mode",
    ]

    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for m in matches:
            writer.writerow(m.to_dict())

    return len(matches)


def export_json(results: List[SearchResult], output_path: str) -> int:
    """
    検索結果をJSON形式でエクスポートする。

    Args:
        results: SearchResultリスト
        output_path: 出力先ファイルパス

    Returns:
        出力したレコード数
    """
    matches = _flatten_results(results)
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    data = {
        "exported_at": datetime.now().isoformat(),
        "total_matches": len(matches),
        "results": [m.to_dict() for m in matches],
    }

    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    return len(matches)


def export_txt(results: List[SearchResult], output_path: str) -> int:
    """
    検索結果を人間可読テキスト形式でエクスポートする。

    Args:
        results: SearchResultリスト
        output_path: 出力先ファイルパス

    Returns:
        出力したレコード数
    """
    matches = _flatten_results(results)
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    with open(path, "w", encoding="utf-8") as f:
        f.write(f"=== Excel Grep 検索結果 ===\n")
        f.write(f"出力日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"総マッチ件数: {len(matches)}件\n")
        f.write("=" * 50 + "\n\n")

        # ファイル別にグループ化
        current_file = None
        for m in matches:
            if m.file_path != current_file:
                current_file = m.file_path
                f.write(f"\n📄 {m.file_path}\n")
                f.write("-" * 40 + "\n")
            f.write(
                f"  [{m.sheet_name}] {m.cell_address} "
                f"キーワード={m.matched_keyword!r} "
                f"値={m.cell_value!r}\n"
            )

        f.write("\n" + "=" * 50 + "\n")
        f.write(f"合計: {len(matches)}件\n")

    return len(matches)


def export_results(
    results: List[SearchResult],
    output_path: str,
    fmt: Optional[str] = None,
    logger=None,
) -> int:
    """
    検索結果をファイルにエクスポートする（形式は拡張子または fmt で決定）。

    Args:
        results: SearchResultリスト
        output_path: 出力先ファイルパス
        fmt: 出力形式 ("csv" / "json" / "txt")。Noneの場合は拡張子から判定。
        logger: ロガーインスタンス

    Returns:
        出力したレコード数
    """
    path = Path(output_path)

    if fmt is None:
        ext = path.suffix.lower()
        fmt = ext.lstrip(".")

    fmt = fmt.lower()

    if fmt == "csv":
        count = export_csv(results, output_path)
    elif fmt == "json":
        count = export_json(results, output_path)
    elif fmt in ("txt", "text"):
        count = export_txt(results, output_path)
    else:
        raise ValueError(f"サポートされていない出力形式です: {fmt}（csv/json/txt を指定してください）")

    if logger:
        logger.info(f"エクスポート完了: {output_path} ({count}件, 形式={fmt})")
    else:
        print(f"✓ エクスポート完了: {output_path} ({count}件)")

    return count


def print_summary(
    results: List[SearchResult],
    elapsed: float,
    logger=None,
    quiet: bool = False,
):
    """
    検索結果サマリーをコンソールに表示する。

    Args:
        results: SearchResultリスト
        elapsed: 処理時間（秒）
        logger: ロガーインスタンス
        quiet: 最小限の出力フラグ
    """
    total_files = len(results)
    error_files = sum(1 for r in results if r and r.skipped)
    total_matches = sum(r.match_count for r in results if r and not r.skipped)

    summary_lines = [
        "=" * 40,
        "検索完了",
        f"総マッチ数: {total_matches}件",
        f"処理ファイル: {total_files}件",
        f"エラー: {error_files}件",
        f"処理時間: {elapsed:.1f}秒",
        "=" * 40,
    ]

    for line in summary_lines:
        if logger:
            logger.info(line)
        elif not quiet:
            print(line)
