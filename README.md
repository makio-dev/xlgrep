# Excel Grep Tool

Excelファイル内のテキストを正規表現対応で検索できるPythonツールです。
CLIとインタラクティブウィザードの両方で実行できます。

---

## 動作環境

- Python 3.10 以上
- macOS / Linux / Windows

---

## セットアップ

```bash
# 1. プロジェクトディレクトリへ移動
cd xlgrep

# 2. 仮想環境を作成
python3 -m venv .venv

# 3. 仮想環境を有効化
source .venv/bin/activate            # macOS / Linux
# .venv\Scripts\Activate.ps1        # Windows PowerShell
# .venv\Scripts\activate.bat        # Windows コマンドプロンプト（cmd）

# 4. 依存ライブラリをインストール
pip install -r requirements.txt
```

> **Note**: 以降のコマンドはすべて仮想環境を有効化した状態（`(.venv)` がプロンプトに表示されている状態）で実行してください。

---

## 使い方

### ウィザードモード（推奨）

```bash
# 仮想環境を有効化してから実行
source .venv/bin/activate
python excel_grep.py --wizard
```

対話型のプロンプトに従うだけで検索できます。

### CLIモード

#### フォルダー内の全Excelファイルを検索

```bash
python excel_grep.py --mode folder --path "/path/to/folder" --keywords "エラー" "警告"
```

#### CSVファイルで対象ファイルを指定して検索

```bash
python excel_grep.py --mode filelist --csv files.csv --keywords "エラー" "警告"
```

#### テキストファイルでパスリストを指定して検索

```bash
python excel_grep.py --mode filelist --input-file paths.txt --keywords "keyword"
```

---

## オプション一覧

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--wizard` | ウィザードモードで起動 | - |
| `--mode folder` | フォルダー指定モード | - |
| `--mode filelist` | ファイルリスト指定モード | - |
| `--path FOLDER` | 検索対象フォルダーパス | - |
| `--csv CSV_PATH` | ファイルリストCSVパス | - |
| `--input-file FILE` | ファイルリストのテキストファイルパス | - |
| `--keywords KW...` | 検索キーワード（最大10個、スペース区切り） | - |
| `--use-regex` | 正規表現モードを有効化 | 無効 |
| `--no-regex` | 正規表現モードを明示的に無効化 | - |
| `--output OUTPUT` | 結果の出力先ファイル（.csv / .json / .txt） | - |
| `--verbose` | 詳細ログを表示 | - |
| `--quiet` | 最小限の出力のみ表示 | - |
| `--log-file LOG` | ログファイルの出力先を指定 | 自動生成 |

---

## 実行例

```bash
# まず仮想環境を有効化
source .venv/bin/activate

# フォルダー検索（通常検索、CSV出力、詳細ログあり）
python excel_grep.py --mode folder --path "/data/logs" --keywords "エラー" "警告" --no-regex --output results.csv --verbose

# 正規表現で日付形式を検索
python excel_grep.py --mode folder --path "/data" --keywords "^\d{4}-\d{2}-\d{2}$" --use-regex --output results.json

# CSVリストを使って静音モードで検索
python excel_grep.py --mode filelist --csv files.csv --keywords "keyword" --quiet --output results.txt
```

---

## CSVファイルリストの形式

`--csv` に渡すCSVは以下の形式で作成してください（`filepath` 列が必須）。

```csv
filepath
/path/to/file1.xlsx
/path/to/file2.xlsx
```

ひな形ファイルは `templates/filelist_template.csv` にあります。
ウィザードモードの「CSVひな形をダウンロードする」でカレントディレクトリにコピーされます。

---

## 出力形式

### CSV（`--output results.csv`）

```
file_path, sheet_name, cell_address, row, col, matched_keyword, cell_value, search_mode
```

### JSON（`--output results.json`）

```json
{
  "exported_at": "2026-03-11T11:04:23",
  "total_matches": 23,
  "results": [...]
}
```

### TXT（`--output results.txt`）

ファイル別にグループ化した人間が読みやすい形式。

---

## ログ

実行のたびに `logs/excel_grep_YYYYMMDD_HHMMSS.log` が自動生成されます。

```
[2026-03-11 11:04:23] [INFO] 検索開始: フォルダーモード
[2026-03-11 11:04:24] [INFO] report1.xlsx: 2件マッチ
[2026-03-11 11:04:26] [WARNING] 読み込みエラー: corrupted.xlsx - スキップ
[2026-03-11 11:04:30] [INFO] 検索完了: 総マッチ=23件, エラー=1件
```

---

## EXE化（スタンドアロン実行ファイルの作成）

Pythonがインストールされていない環境へ配布する場合は、PyInstallerで単一の実行ファイルを生成します。

> **Note**: EXE化は配布先と**同じOS**の環境で行う必要があります。
> Windows配布用EXEはWindows上で、Mac配布用バイナリはMac上でビルドしてください。

### Windows の場合

```batch
build_exe.bat
```

`dist\excel_grep.exe` が生成されます。

```batch
REM 実行例
dist\excel_grep.exe --wizard
dist\excel_grep.exe --mode folder --path "C:\data" --keywords "エラー" --verbose
```

### macOS / Linux の場合

```bash
chmod +x build_mac.sh
./build_mac.sh
```

`dist/excel_grep` が生成されます。

```bash
# 実行例
./dist/excel_grep --wizard
./dist/excel_grep --mode folder --path "/data" --keywords "エラー" --verbose
```

### 手動でビルドする場合

```bash
# venv を有効化した状態で
source .venv/bin/activate

pip install pyinstaller
pyinstaller excel_grep.spec --clean
```

### 配布パッケージの構成

```
配布フォルダ/
├── excel_grep.exe（または excel_grep）  # 実行ファイル
├── templates/
│   └── filelist_template.csv            # CSVひな形
└── logs/                                # 自動生成
```

### 注意事項

- ウイルス対策ソフトが誤検知する場合があります（初回実行時に許可が必要なことがあります）
- 特定フォルダーへのアクセスには管理者権限が必要な場合があります
- `logs/` ディレクトリは実行ファイルと同じ階層に自動生成されます

---

## プロジェクト構造

```
xlgrep/
├── excel_grep.py          # メインスクリプト
├── core/
│   ├── searcher.py        # 検索エンジン（並列処理・正規表現対応）
│   ├── file_handler.py    # ファイル収集（フォルダー/CSV/テキスト）
│   ├── exporter.py        # 結果出力（CSV/JSON/TXT）
│   └── logger.py          # ログ管理
├── cli/
│   ├── parser.py          # CLI引数解析
│   └── wizard.py          # ウィザードUI
├── templates/
│   └── filelist_template.csv  # CSVひな形
├── logs/                  # ログ出力ディレクトリ（自動生成）
├── requirements.txt
└── README.md
```
