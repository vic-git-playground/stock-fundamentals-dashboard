@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

rem Not normally needed - index.html works by just double-clicking it.
rem This is only a fallback local web server, kept in case that ever changes.
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

echo Could not find python.exe automatically.
echo Edit this file and set PYTHON_EXE to your python.exe path, for example:
echo     set "PYTHON_EXE=C:\Users\yourname\anaconda3\python.exe"
pause
exit /b 1

:found
echo Starting local web server at http://localhost:8765 ...
start http://localhost:8765/index.html
"%PYTHON_EXE%" -m http.server 8765
pause
