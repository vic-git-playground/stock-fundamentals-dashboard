@echo off
chcp 65001 >nul
cd /d "%~dp0"

if not exist ".git" (
  echo 這個資料夾還沒做過 git 初始化，請先照 DEPLOY.md 的「第一次設定」步驟做一次。
  pause
  exit /b 1
)

echo 推送資料更新到 GitHub（Cloudflare Pages 偵測到後會自動重新部署，通常 1~2 分鐘）...
git add -A
git commit -m "update data" >nul 2>nul
rem 上面這行如果「沒有變更可以 commit」會回傳非 0，那是正常情況，不代表出錯，所以不檢查它的結果
git push

if errorlevel 1 (
  echo.
  echo 推送失敗。常見原因：
  echo   1) 還沒照 DEPLOY.md 做過「第一次設定」（git remote add / 第一次登入）
  echo   2) 網路連不到 GitHub
  pause
  exit /b 1
)

echo.
echo 完成！等 1~2 分鐘讓 Cloudflare Pages 部署好就可以看到最新資料。
pause
