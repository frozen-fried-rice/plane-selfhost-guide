@echo off
setlocal
rem xlsx -> CSV 変換（先頭2行削除）ダブルクリックで実行
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0xlsx2csv.ps1" %*
echo.
echo 続行するには何かキーを押してください . . .
pause >nul
