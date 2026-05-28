@echo off
rem chcp 65001 #本ファイルはShift JISで開く
setlocal

set SRC=%~dp0成約捕捉ツール
set DST=C:\成約捕捉ツール

echo フォルダを同期中...
robocopy "%SRC%" "%DST%" /E /XO /MT:8 /R:1 /W:1 /FFT /NFL /NDL /NJH /NJS /NP
if errorlevel 8 (
    echo 同期に失敗しました
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

timeout /t 3 /nobreak >nul
exit /b
