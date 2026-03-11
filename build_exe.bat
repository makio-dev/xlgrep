@echo off
chcp 65001 > nul
echo ========================================
echo   Excel Grep Tool - EXE Build Script
echo ========================================
echo.

REM 仮想環境の確認
echo [1/5] Checking Python environment...
python --version
if %errorlevel% neq 0 (
    echo ERROR: Python not found! Please install Python 3.10+.
    pause
    exit /b 1
)

REM 依存関係のインストール
echo.
echo [2/5] Installing dependencies...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies!
    pause
    exit /b 1
)

REM PyInstallerのインストール確認
echo.
echo [3/5] Checking PyInstaller...
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing PyInstaller...
    pip install pyinstaller
    if %errorlevel% neq 0 (
        echo ERROR: Failed to install PyInstaller!
        pause
        exit /b 1
    )
)
echo PyInstaller OK.

REM 既存のビルドファイル削除
echo.
echo [4/5] Cleaning previous build...
if exist build (
    rmdir /s /q build
    echo Removed: build/
)
if exist dist (
    rmdir /s /q dist
    echo Removed: dist/
)

REM EXEビルド実行
echo.
echo [5/5] Building executable...
pyinstaller excel_grep.spec --clean
if %errorlevel% neq 0 (
    echo ERROR: Build failed!
    pause
    exit /b 1
)

echo.
echo ========================================
echo   Build completed successfully!
echo   Executable: dist\excel_grep.exe
echo ========================================
echo.
echo 使用方法:
echo   dist\excel_grep.exe --wizard
echo   dist\excel_grep.exe --mode folder --path "C:\data" --keywords "keyword"
echo.
pause
