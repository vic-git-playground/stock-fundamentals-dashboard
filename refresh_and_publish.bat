@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

rem 用於：你自己已經在 Excel 裡手動更新完 CMoney 資料（沒有用 update_from_cmoney.bat），
rem 現在只要「重新匯出網站資料 + 推送上網」，不會再呼叫一次 CMoney 更新。
rem 執行前請先把 Excel 檔案存檔關閉，避免檔案被鎖住讀不到。

set "PYTHON_EXE="

where python >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_EXE=python"
    goto :found
)

for %%P in (
    "%USERPROFILE%\anaconda3\python.exe"
    "%USERPROFILE%\miniconda3\python.exe"
    "%LOCALAPPDATA%\anaconda3\python.exe"
    "%LOCALAPPDATA%\miniconda3\python.exe"
    "%LOCALAPPDATA%\Continuum\anaconda3\python.exe"
    "C:\ProgramData\anaconda3\python.exe"
    "C:\ProgramData\miniconda3\python.exe"
    "C:\Anaconda3\python.exe"
    "C:\Miniconda3\python.exe"
) do (
    if exist %%P (
        set "PYTHON_EXE=%%~P"
        goto :found
    )
)

echo 找不到 python.exe，請打開這個檔案，把 PYTHON_EXE 設成你的 python.exe 完整路徑，例如：
echo     set "PYTHON_EXE=C:\Users\yourname\anaconda3\python.exe"
pause
exit /b 1

:found
echo 使用 Python: %PYTHON_EXE%
echo.
echo 重新匯出網站資料（讀取 Excel 目前內容，不會呼叫 CMoney 更新）...
"%PYTHON_EXE%" "tools\refresh_data.py" "D:\工作區\主要工作檔\台灣篩選器_股期_ver.2.xlsm" "data"

if errorlevel 1 (
  echo.
  echo 匯出失敗，請確認 Excel 檔案已經存檔關閉，沒有被鎖住。
  pause
  exit /b 1
)

echo.
echo 匯出完成，接著推送上網...
call publish.bat
