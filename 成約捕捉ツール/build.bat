@echo off
chcp 65001 >nul
setlocal

echo ============================================================
echo  Build Script / サンレモ成約捕捉 ビルドスクリプト (Tkinter版)
echo ============================================================
echo.

REM --- Python check --- [cite: 2]
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python が見つかりません。
    pause
    exit /b 1
)

REM --- PyInstaller check ---
python -c "import PyInstaller" >nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] PyInstaller をインストールします...
    pip install pyinstaller
)

echo [INFO] 必要なライブラリを確認中...
REM NiceGUI関連を削除し、tkinterアプリに必要なものに限定
pip install xlwings pandas openpyxl pyyaml msoffcrypto-tool
echo.

echo [INFO] ビルド開始... [cite: 3]
echo.

REM specファイルを使用してビルド
pyinstaller sanremo.spec --clean

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] ビルド失敗。エラーを確認してください。
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  ビルド完了! [cite: 4]
echo  出力先: dist\sanremo\
echo ============================================================
echo.

set DIST_DIR=dist\sanremo [cite: 5]

if exist "config.yml" (
    copy /Y "config.yml" "%DIST_DIR%\config.yml"
    echo 設定ファイルをコピーしました: config.yml [cite: 5]
) else (
    echo [WARNING] config.yml が見つかりません。
)

echo.
echo [INFO] FMT等のExcelファイルを %DIST_DIR%\ へコピーしてください。 [cite: 6]
echo.

pause