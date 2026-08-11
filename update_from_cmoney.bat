@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

rem 這個 bat 要在「有裝 CMoney」的那台電腦上執行，這台工作機沒有裝 CMoney 沒辦法跑。
rem 流程：CMExcel.exe 更新 台灣篩選器_股期_ver.2.xlsm -> 存檔關閉 -> 重新匯出網站資料 (data/all_data.js)

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
echo 需要先安裝 pywin32 才能控制 Excel 存檔關閉（如果之前沒裝過，這裡先裝一次，已裝過會秒過）：
"%PYTHON_EXE%" -m pip install --quiet pywin32

echo.
echo 開始執行 CMoney 更新 + 網站資料重新匯出...
echo （CMExcel.exe 更新過程通常要數十秒，中間不要動 Excel）
echo.
"%PYTHON_EXE%" "tools\update_from_cmoney.py"

echo.
pause
