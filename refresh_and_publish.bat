@echo off
setlocal
cd /d "%~dp0"

rem 用於：你自己已經在 Excel 裡手動更新完 CMoney 資料，現在要「重新匯出網站資料 + 推送上網」。
rem 執行前請先把 Excel 檔案存檔關閉，避免檔案被鎖住讀不到。
rem
rem 寫法說明：這裡刻意全部用 goto 做流程控制，不用 if ( ... ) 的括號區塊。
rem 因為 cmd 會把 echo 文字裡沒跳脫的右括號當成區塊結束，很容易造成
rem 「明明成功卻提前 pause / exit」這種很難查的問題。

set "LOG=%~dp0publish_log.txt"
call :main > "%LOG%" 2>&1
set "RC=%ERRORLEVEL%"

type "%LOG%"
echo.
if "%RC%"=="0" goto :ok

echo ==========================================================
echo  這次沒有跑完，完整訊息已經存到：
echo    %LOG%
echo  可以把這個檔案的內容貼給 Claude 看
echo ==========================================================
pause
exit /b %RC%

:ok
echo ==========================================================
echo  完成！等 1 分鐘左右 GitHub Pages 部署好就會看到最新資料。
echo ==========================================================
pause
exit /b 0


rem ================= 以下是實際步驟，輸出都會進 log =================
:main
echo [步驟 1/3] 尋找 Python
set "PYTHON_EXE="
where python >nul 2>nul
if %errorlevel%==0 set "PYTHON_EXE=python"
if defined PYTHON_EXE goto :got_python

for %%P in (
    "%USERPROFILE%\anaconda3\python.exe"
    "%USERPROFILE%\miniconda3\python.exe"
    "%LOCALAPPDATA%\anaconda3\python.exe"
    "%LOCALAPPDATA%\miniconda3\python.exe"
    "C:\ProgramData\anaconda3\python.exe"
    "C:\Anaconda3\python.exe"
) do if exist %%P if not defined PYTHON_EXE set "PYTHON_EXE=%%~P"

if not defined PYTHON_EXE goto :no_python
goto :got_python

:no_python
echo [錯誤] 找不到 python.exe
echo        請打開這個 bat，把 PYTHON_EXE 設成你的 python.exe 完整路徑
exit /b 9

:got_python
echo        使用 Python: %PYTHON_EXE%
echo.
echo [步驟 2/3] 重新匯出網站資料（讀取 Excel 目前內容重新計算）
"%PYTHON_EXE%" "tools\refresh_data.py" "D:\工作區\主要工作檔\台灣篩選器_股期_ver.2.xlsm" "data"
if errorlevel 1 goto :export_failed
echo        匯出完成
echo.

echo [步驟 3/3] 推送到 GitHub
if not exist ".git" goto :no_git
echo        git add
git add -A
if errorlevel 1 goto :git_failed
echo        git commit
git commit -m "update data"
rem 沒有變更可以 commit 時 git 會回傳非 0，那是正常情況，所以這裡不檢查
echo        git push
git push
if errorlevel 1 goto :push_failed
echo        推送成功
exit /b 0

:export_failed
echo [錯誤] 匯出失敗，原因請看上面的訊息。常見狀況：
echo        - Excel 檔案還開著或被鎖住，存檔關閉後再跑一次
echo        - 有檔案總管開著 data\chunks 資料夾，關掉那個視窗再跑一次
exit /b 2

:no_git
echo [錯誤] 這個資料夾沒有 .git，不是 git 工作區，沒辦法推送
exit /b 3

:git_failed
echo [錯誤] git add 失敗
exit /b 4

:push_failed
echo [錯誤] git push 失敗。常見狀況：
echo        - 需要重新登入 GitHub
echo        - 網路連不到 GitHub
echo        - 遠端有別人推的新 commit，需要先 git pull
exit /b 5
