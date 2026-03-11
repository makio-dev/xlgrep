#!/usr/bin/env python3
"""
Excel Grep Tool - メインエントリポイント
ExcelファイルのテキストをCLI/ウィザードで検索するツール。
"""
import sys
import time
from pathlib import Path

# プロジェクトルートをパスに追加（EXE化時のパス解決対応）
project_root = Path(__file__).parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

try:
    from colorama import Fore, Style, init as colorama_init
    colorama_init(autoreset=True)
    COLORAMA = True
except ImportError:
    COLORAMA = False


def _color(text: str, color: str) -> str:
    if not COLORAMA:
        return text
    colors = {
        "green": Fore.GREEN,
        "yellow": Fore.YELLOW,
        "red": Fore.RED,
        "cyan": Fore.CYAN,
        "bold": Style.BRIGHT,
    }
    return colors.get(color, "") + text + Style.RESET_ALL


def _print_banner(verbose: bool = False):
    if not verbose:
        return
    print(_color("=" * 44, "cyan"))
    print(_color("  Excel Grep Tool v1.0", "bold"))
    print(_color("=" * 44, "cyan"))
    print()


def run_cli(args) -> int:
    """
    CLIモードで検索を実行する。

    Returns:
        終了コード（0=成功、1=エラー）
    """
    from core.logger import ExcelGrepLogger
    from core.file_handler import (
        collect_excel_files_from_folder,
        collect_excel_files_from_csv,
        collect_excel_files_from_text,
        validate_excel_files,
    )
    from core.searcher import ExcelSearcher
    from core.exporter import export_results, print_summary

    # ロガー初期化
    logger = ExcelGrepLogger(
        log_dir=args.log_dir,
        log_file=args.log_file,
        verbose=args.verbose,
        quiet=args.quiet,
    )

    _print_banner(args.verbose)

    if not args.quiet:
        mode_str = "フォルダー検索" if args.mode == "folder" else "ファイルリスト検索"
        regex_str = "有効" if args.use_regex else "無効"
        logger.info(f"検索開始: モード={mode_str}")
        logger.info(f"キーワード: {args.keywords}")
        logger.info(f"正規表現: {regex_str}")

    # ─── ファイルリスト収集 ───────────────────────────────────────
    try:
        if args.mode == "folder":
            logger.info(f"対象パス: {args.path}")
            file_list = collect_excel_files_from_folder(args.path)
        elif args.mode == "filelist":
            if args.csv:
                logger.info(f"CSVファイル: {args.csv}")
                file_list = collect_excel_files_from_csv(args.csv)
            else:
                logger.info(f"テキストファイル: {args.input_file}")
                file_list = collect_excel_files_from_text(args.input_file, is_file=True)
    except ValueError as e:
        logger.error(str(e))
        return 1

    # ファイル検証
    valid_files, invalid_files = validate_excel_files(file_list, logger=logger)

    if not valid_files:
        logger.error("有効なExcelファイルがありません。終了します。")
        return 1

    logger.info(f"{len(valid_files)}件のExcelファイルを処理します")

    # ─── プログレスコールバック（tqdm対応） ────────────────────────
    try:
        from tqdm import tqdm
        pbar = tqdm(total=len(valid_files), desc="検索中", unit="files", disable=args.quiet)

        def progress_callback(current: int, total: int, file_path: str, total_matches: int):
            pbar.n = current
            pbar.set_postfix({"マッチ": total_matches, "現在": Path(file_path).name})
            pbar.refresh()

    except ImportError:
        pbar = None
        progress_callback = None

    # ─── 検索実行 ─────────────────────────────────────────────────
    try:
        searcher = ExcelSearcher(
            keywords=args.keywords,
            use_regex=args.use_regex,
            logger=logger,
            progress_callback=progress_callback,
        )
    except ValueError as e:
        logger.error(str(e))
        return 1

    start = time.time()
    results = searcher.search(valid_files)
    elapsed = time.time() - start

    if pbar:
        pbar.close()

    # ─── サマリー表示 ─────────────────────────────────────────────
    print_summary(results, elapsed, logger=logger if args.verbose else None, quiet=args.quiet)

    if not args.quiet:
        logger.info(f"ログファイル: {logger.get_log_file_path()}")

    # ─── 結果エクスポート ──────────────────────────────────────────
    if args.output:
        try:
            fmt = args.output_format  # Noneの場合は拡張子から判定
            export_results(results, args.output, fmt=fmt, logger=logger)
        except Exception as e:
            logger.error(f"エクスポートエラー: {e}")
            return 1

    # ─── 置換処理 ─────────────────────────────────────────────────
    if getattr(args, "replacements", None):
        from core.replacer import replace_files

        replace_map = dict(zip(args.keywords, args.replacements))
        dry_run = getattr(args, "dry_run", False)
        backup = getattr(args, "backup", True)

        logger.info("─" * 44)
        if dry_run:
            logger.info("[DRY-RUN] プレビューモード - ファイルは変更されません")
        else:
            logger.info(f"置換開始: backup={backup}")

        for k, v in replace_map.items():
            logger.info(f'  "{k}"  →  "{v}"')

        replace_results = replace_files(
            valid_files,
            replace_map,
            use_regex=args.use_regex,
            backup=backup,
            dry_run=dry_run,
            logger=logger,
        )

        total_replaced = sum(r.replace_count for r in replace_results if not r.skipped)
        replace_errors = sum(1 for r in replace_results if r.skipped)

        if dry_run:
            logger.info("─" * 44)
            logger.info(f"[DRY-RUN] {total_replaced}件を置換する予定")
            for rr in replace_results:
                for rec in rr.records:
                    logger.info(f'  [{rec.sheet_name}] {rec.cell_address}  "{rec.before}"  →  "{rec.after}"')
        else:
            logger.info(f"置換完了: {total_replaced}件置換, エラー: {replace_errors}件")

    return 0



def main():
    """メイン関数"""
    from cli.parser import parse_args

    args = parse_args()

    if args.wizard:
        from cli.wizard import run_wizard
        run_wizard()
        sys.exit(0)
    else:
        exit_code = run_cli(args)
        sys.exit(exit_code)


if __name__ == "__main__":
    main()
