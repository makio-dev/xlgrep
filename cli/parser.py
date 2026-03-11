"""
CLI引数解析モジュール
argparse を使ってコマンドライン引数を定義・検証する。
"""
import argparse
import sys
from pathlib import Path


def build_parser() -> argparse.ArgumentParser:
    """ArgumentParser を構築して返す"""
    parser = argparse.ArgumentParser(
        prog="excel_grep",
        description="Excel Grep Tool - Excelファイル内テキスト検索ツール",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  # フォルダー指定（正規表現使用）
  python excel_grep.py --mode folder --path "C:\\data" --keywords "keyword1" "keyword2" --use-regex --verbose

  # CSVファイルリスト指定
  python excel_grep.py --mode filelist --csv files.csv --keywords "エラー" "警告" --no-regex --output results.csv

  # テキストファイル指定
  python excel_grep.py --mode filelist --input-file paths.txt --keywords "keyword" --quiet

  # ウィザードモード
  python excel_grep.py --wizard
        """,
    )

    # ウィザードモード
    parser.add_argument(
        "--wizard",
        action="store_true",
        help="インタラクティブウィザードモードで実行する",
    )

    # 検索モード
    mode_group = parser.add_argument_group("検索モード")
    mode_group.add_argument(
        "--mode",
        choices=["folder", "filelist"],
        help="検索モード: folder（フォルダー指定）/ filelist（ファイルリスト指定）",
    )
    mode_group.add_argument(
        "--path",
        metavar="FOLDER_PATH",
        help="検索対象フォルダーパス（--mode folder 時に必須）",
    )
    mode_group.add_argument(
        "--csv",
        metavar="CSV_PATH",
        help="ファイルリストCSVのパス（--mode filelist 時に使用）",
    )
    mode_group.add_argument(
        "--input-file",
        metavar="FILE_PATH",
        dest="input_file",
        help="ファイルリストのテキストファイルパス（--mode filelist 時に使用）",
    )

    # キーワード
    kw_group = parser.add_argument_group("キーワード設定")
    kw_group.add_argument(
        "--keywords",
        nargs="+",
        metavar="KEYWORD",
        help="検索キーワード（最大10個、スペース区切りで複数指定可）",
    )

    # 正規表現
    regex_group = kw_group.add_mutually_exclusive_group()
    regex_group.add_argument(
        "--use-regex",
        action="store_true",
        default=False,
        dest="use_regex",
        help="正規表現モードを有効化（デフォルト: 無効）",
    )
    regex_group.add_argument(
        "--no-regex",
        action="store_true",
        default=False,
        dest="no_regex",
        help="正規表現モードを明示的に無効化",
    )

    # 置換設定
    rep_group = parser.add_argument_group("置換設定")
    rep_group.add_argument(
        "--replacements",
        nargs="+",
        metavar="REPLACEMENT",
        help="置換文字列（--keywords と同数指定。指定時は検索後に置換を実行する）",
    )
    rep_group.add_argument(
        "--dry-run",
        action="store_true",
        dest="dry_run",
        help="置換のプレビューのみ実行（ファイルは書き換えない）",
    )
    rep_group.add_argument(
        "--backup",
        action="store_true",
        default=True,
        dest="backup",
        help="置換前にバックアップ (.bak) を作成する（デフォルト: 有効）",
    )
    rep_group.add_argument(
        "--no-backup",
        action="store_false",
        dest="backup",
        help="バックアップを作成しない",
    )

    # 出力設定
    out_group = parser.add_argument_group("出力設定")
    out_group.add_argument(
        "--output",
        metavar="OUTPUT_PATH",
        help="結果出力先ファイルパス（拡張子: .csv / .json / .txt）",
    )
    out_group.add_argument(
        "--output-format",
        choices=["csv", "json", "txt"],
        dest="output_format",
        help="出力形式を明示的に指定（--output の拡張子より優先）",
    )

    # ログ設定
    log_group = parser.add_argument_group("ログ設定")
    log_group.add_argument(
        "--verbose",
        action="store_true",
        help="詳細ログを表示する",
    )
    log_group.add_argument(
        "--quiet",
        action="store_true",
        help="最小限の出力のみ表示する",
    )
    log_group.add_argument(
        "--log-file",
        metavar="LOG_PATH",
        dest="log_file",
        help="ログファイル出力先を指定する（デフォルト: logs/excel_grep_YYYYMMDD_HHMMSS.log）",
    )
    log_group.add_argument(
        "--log-dir",
        metavar="LOG_DIR",
        dest="log_dir",
        default="logs",
        help="ログ出力ディレクトリ（デフォルト: logs）",
    )

    # バージョン
    parser.add_argument(
        "--version",
        action="version",
        version="Excel Grep Tool v1.0",
    )

    return parser


def validate_args(args: argparse.Namespace) -> list[str]:
    """
    引数の論理検証を行い、エラーメッセージリストを返す。
    エラーがなければ空リストを返す。
    """
    errors = []

    if args.wizard:
        # ウィザードモードは他の引数バリデーション不要
        return errors

    # モードは必須
    if not args.mode:
        errors.append("--mode を指定してください（folder または filelist）")
        return errors

    if args.mode == "folder":
        if not args.path:
            errors.append("--mode folder の場合は --path を指定してください")
    elif args.mode == "filelist":
        if not args.csv and not args.input_file:
            errors.append("--mode filelist の場合は --csv または --input-file を指定してください")
        if args.csv and args.input_file:
            errors.append("--csv と --input-file は同時に指定できません")

    # キーワードは必須
    if not args.keywords:
        errors.append("--keywords を1個以上指定してください")
    elif len(args.keywords) > 10:
        errors.append(f"--keywords は最大10個までです（現在: {len(args.keywords)}個）")

    # 置換文字列の個数チェック
    if args.replacements:
        if len(args.replacements) != len(args.keywords or []):
            errors.append(
                f"--replacements の個数 ({len(args.replacements)}) は "
                f"--keywords の個数 ({len(args.keywords or [])}) と一致させてください"
            )

    # dry-run は --replacements がある場合のみ意味を持つ（警告は出さないが注記）
    if args.dry_run and not args.replacements:
        errors.append("--dry-run は --replacements と合わせて指定してください")

    # verbose と quiet は同時指定不可
    if args.verbose and args.quiet:
        errors.append("--verbose と --quiet は同時に指定できません")

    return errors


def parse_args(argv=None) -> argparse.Namespace:
    """
    コマンドライン引数をパースして Namespace を返す。
    バリデーションエラーがあれば表示して終了する。
    """
    parser = build_parser()
    args = parser.parse_args(argv)

    errors = validate_args(args)
    if errors:
        print("エラー:", file=sys.stderr)
        for e in errors:
            print(f"  - {e}", file=sys.stderr)
        print("\n使い方の詳細は --help を参照してください。", file=sys.stderr)
        sys.exit(1)

    return args
