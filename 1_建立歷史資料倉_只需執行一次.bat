@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

rem ============================================================
rem  一次性作業：把「現在這份還有完整歷史的 Excel」灌進歷史資料倉 archive\
rem  這一步一定要在把 CMoney 的 TOP N 改小之前做完，做完請把 archive\ 另外備份。
rem ============================================================

set "XLSM=D:\工作區\主要工作檔\台灣篩選器_股期_ver.2.xlsm"

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
    "C:\ProgramData\anaconda3\python.exe"
    "C:\Anaconda3\python.exe"
) do (
    if exist %%P (
        set "PYTHON_EXE=%%~P"
        goto :found
    )
)
echo 找不到 python.exe，請打開這個檔案把 PYTHON_EXE 設成完整路徑。
pause
exit /b 1

:found
echo 使用 Python: %PYTHON_EXE%
echo 來源 Excel : %XLSM%
echo.
"%PYTHON_EXE%" "tools\bootstrap_archive.py" "%XLSM%"

echo.
pause
