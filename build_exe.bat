@echo off
chcp 65001 > nul
echo ========================================
echo   Excel Grep Tool - EXE Build Script
echo ========================================
echo.

REM Python確認
echo [1/5] Checking Python environment...
python --version
if %errorlevel% neq 0 (
    echo ERROR: Python not found! Please install Python 3.10+.
    pause
    exit /b 1
)

REM venv作成（なければ）
echo.
echo [2/5] Setting up virtual environment...
if not exist .venv (
    python -m venv .venv
    if %errorlevel% neq 0 (
        echo ERROR: Failed to create virtual environment!
        pause
        exit /b 1
    )
    echo Created: .venv
) else (
    echo Already exists: .venv
)

REM 依存関係インストール（venv内）
echo.
echo [3/5] Installing dependencies...
.venv\Scripts\python.exe -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies!
    pause
    exit /b 1
)

REM PyInstaller確認・インストール（venv内）
echo.
echo [4/5] Checking PyInstaller...
.venv\Scripts\python.exe -m pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing PyInstaller...
    .venv\Scripts\python.exe -m pip install pyinstaller
    if %errorlevel% neq 0 (
        echo ERROR: Failed to install PyInstaller!
        pause
        exit /b 1
    )
)
echo PyInstaller OK.

REM 既存ビルドファイル削除
echo.
echo [5/5] Building executable...
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist

REM ビルド実行（venv内のpyinstaller）
.venv\Scripts\pyinstaller.exe excel_grep.spec --clean
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
echo 実行方法:
echo   dist\excel_grep.exe --wizard
echo   dist\excel_grep.exe --mode folder --path "C:\data" --keywords "keyword"
echo.
pause
