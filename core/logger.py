"""
ログ管理モジュール
ファイルログ出力、コンソールログ出力、ログレベル管理を提供する。
"""
import logging
import sys
from datetime import datetime
from pathlib import Path


class ExcelGrepLogger:
    """Excel Grep ツール用ロガー"""

    def __init__(self, log_dir: str = "logs", log_file: str = None, verbose: bool = False, quiet: bool = False):
        """
        Args:
            log_dir: ログ出力ディレクトリ
            log_file: ログファイルパス（Noneの場合は自動生成）
            verbose: 詳細ログ表示フラグ
            quiet: 最小限の出力フラグ
        """
        self.verbose = verbose
        self.quiet = quiet
        self.logger = logging.getLogger("excel_grep")
        self.logger.setLevel(logging.DEBUG)
        self.logger.handlers.clear()

        # ログディレクトリ作成
        log_path = Path(log_dir)
        log_path.mkdir(parents=True, exist_ok=True)

        # ログファイルパス決定
        if log_file is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            log_file = log_path / f"excel_grep_{timestamp}.log"
        self.log_file_path = str(log_file)

        # フォーマット定義
        log_format = "[%(asctime)s] [%(levelname)s] %(message)s"
        date_format = "%Y-%m-%d %H:%M:%S"
        formatter = logging.Formatter(log_format, datefmt=date_format)

        # ファイルハンドラ（常にDEBUGレベル）
        try:
            file_handler = logging.FileHandler(self.log_file_path, encoding="utf-8")
            file_handler.setLevel(logging.DEBUG)
            file_handler.setFormatter(formatter)
            self.logger.addHandler(file_handler)
        except Exception as e:
            print(f"警告: ログファイルを作成できませんでした: {e}", file=sys.stderr)

        # コンソールハンドラ
        if not quiet:
            console_handler = logging.StreamHandler(sys.stdout)
            if verbose:
                console_handler.setLevel(logging.DEBUG)
            else:
                console_handler.setLevel(logging.INFO)

            # コンソール用シンプルフォーマット
            console_format = "[%(asctime)s] %(message)s"
            console_formatter = logging.Formatter(console_format, datefmt="%H:%M:%S")
            console_handler.setFormatter(console_formatter)
            self.logger.addHandler(console_handler)

    def debug(self, msg: str):
        self.logger.debug(msg)

    def info(self, msg: str):
        self.logger.info(msg)

    def warning(self, msg: str):
        self.logger.warning(msg)

    def error(self, msg: str):
        self.logger.error(msg)

    def critical(self, msg: str):
        self.logger.critical(msg)

    def get_log_file_path(self) -> str:
        return self.log_file_path


# グローバルロガーインスタンス（ウィザードモード用コールバック対応版）
class CallbackLogger(ExcelGrepLogger):
    """ウィザードモード向け：ログメッセージをコールバックで通知するロガー"""

    def __init__(self, callback=None, **kwargs):
        """
        Args:
            callback: ログメッセージを受け取るコールバック関数 callback(level, message)
        """
        super().__init__(**kwargs)
        self.callback = callback

    def _notify(self, level: str, msg: str):
        if self.callback:
            self.callback(level, msg)

    def info(self, msg: str):
        super().info(msg)
        self._notify("INFO", msg)

    def warning(self, msg: str):
        super().warning(msg)
        self._notify("WARNING", msg)

    def error(self, msg: str):
        super().error(msg)
        self._notify("ERROR", msg)
