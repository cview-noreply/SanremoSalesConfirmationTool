@echo off
rem chcp 65001
setlocal

set SRC=%~dp0成約捕捉ツール
set DST=C:\成約捕捉ツール

if exist "%DST%" (
    echo 既存フォルダを削除中...
    rd /s /q "%DST%"
    
    rem 削除できたか確認
    if exist "%DST%" (
        echo 既存フォルダの削除に失敗しました
        pause
        exit /b 1
    )
)

echo フォルダをコピー中...
robocopy "%SRC%" "%DST%" /E /R:1 /W:1 /NFL /NDL /NJH /NJS /NP
if errorlevel 8 (
    echo コピーに失敗しました
    pause
    exit /b 1
)

if not exist "%DST%\サンレモ成約捕捉ツール.exe" (
    echo 実行ファイルが見つかりません
    pause
    exit /b 1
)

echo ツールを起動します...
start "" /d "%DST%" "%DST%\サンレモ成約捕捉ツール.exe"

exit /b
