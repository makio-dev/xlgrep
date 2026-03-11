#!/bin/bash
set -e

echo "========================================"
echo "  Excel Grep Tool - Build Script (Mac)"
echo "========================================"
echo

# Python確認
echo "[1/5] Checking Python environment..."
python3 --version || { echo "ERROR: Python3 not found!"; exit 1; }

# 依存関係インストール
echo
echo "[2/5] Installing dependencies..."
if [ -d ".venv" ]; then
    .venv/bin/pip install -r requirements.txt
else
    python3 -m venv .venv
    .venv/bin/pip install -r requirements.txt
fi

# PyInstaller確認
echo
echo "[3/5] Checking PyInstaller..."
.venv/bin/pip show pyinstaller > /dev/null 2>&1 || .venv/bin/pip install pyinstaller

# クリーンアップ
echo
echo "[4/5] Cleaning previous build..."
rm -rf build dist

# ビルド実行
echo
echo "[5/5] Building executable..."
.venv/bin/pyinstaller excel_grep.spec --clean

echo
echo "========================================"
echo "  Build completed!"
echo "  Executable: dist/excel_grep"
echo "========================================"
echo
echo "実行方法:"
echo "  ./dist/excel_grep --wizard"
echo "  ./dist/excel_grep --mode folder --path /your/data --keywords 'keyword'"
