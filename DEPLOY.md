# 把網站放到網路上分享（一次性設定）

方案：GitHub 放程式碼（私有倉庫）＋ Cloudflare Pages 部署（免費、支援帳號密碼保護）。
GitHub Pages 免費方案沒辦法加密碼、放上去就是公開給任何人看，所以用 Cloudflare Pages 這層來擋。

## 第一次設定（大概 10 分鐘）

### 1. 在 GitHub 建一個新倉庫
- 到 https://github.com/new
- Repository name 填你喜歡的名字，例如 `stock-dashboard`
- 選 **Private**（私有，不要選 Public——裡面的資料是從付費的 CMoney 資料算出來的，不適合公開）
- 不要勾 "Add a README"，其他都不用選，直接建立

### 2. 把這個資料夾推上去
在這個資料夾（`D:\工作區\主要工作檔\Claude專案\基本面數據網站`）打開命令提示字元或 PowerShell，依序執行：

```
git init
git add -A
git commit -m "first commit"
git branch -M main
git remote add origin https://github.com/你的帳號/stock-dashboard.git
git push -u origin main
```

第一次 push 會跳出視窗要你登入 GitHub 帳號授權，登入一次之後電腦會記住，之後 `publish.bat` 就不用再登入。

（如果 `git` 指令說找不到，先安裝 Git for Windows：https://git-scm.com/download/win）

### 3. 到 Cloudflare Pages 建立專案
- 到 https://dash.cloudflare.com/ 註冊/登入（免費帳號即可）
- 左側選 **Workers & Pages** -> **Create** -> **Pages** -> **Connect to Git**
- 授權並選剛剛建立的 GitHub 倉庫（例如 `stock-dashboard`）
- Build settings 全部留空/ 預設值即可（這是純靜態網站，不需要 build command，
  Build output directory 留空或填 `/` 都可以）
- 按 **Save and Deploy**，等個 1 分鐘左右它就會部署完成，會給你一個
  `https://stock-dashboard-xxx.pages.dev` 這樣的網址

### 4. 設定帳號密碼保護
- 進剛剛建立的 Pages 專案 -> **Settings** -> **Environment variables**
- 新增兩個變數（Production 環境）：
  - `SITE_USER` = 你要設的帳號，例如 `vic`
  - `SITE_PASSWORD` = 你要設的密碼
- 存檔後，回到 **Deployments** 頁籤，找最新那筆部署，點右邊 `...` -> **Retry deployment**
  （環境變數要重新部署一次才會套用）

### 5. 分享給其他人
把 `SITE_USER`、`SITE_PASSWORD` 和網址（`https://stock-dashboard-xxx.pages.dev`）一起傳給你要分享的人即可。
之後想換掉密碼、或想撤銷某個人的存取權，就是回 Cloudflare 改這兩個環境變數再重新部署——沒辦法針對
「某一個人」個別開關，是整站共用同一組帳密（如果之後想要每個人有自己的帳號、或想踢掉特定的人，
跟我說一聲，可以改用 Cloudflare Access 做每人各自登入，只是設定步驟會多一些）。

## 之後資料更新怎麼推上網

雙擊資料夾裡的 **`publish.bat`**：它會把 `git add -A / commit / push` 做完，Cloudflare Pages 偵測到
GitHub 有新的 commit 會自動重新部署，通常 1~2 分鐘網站就會更新。

如果是在有裝 CMoney 的那台電腦跑 `update_from_cmoney.bat`，它跑完重新匯出資料後，只要那台電腦也做過
上面「第一次設定」的第 2 步（`git remote add origin ...` 且已經登入過），會自動幫你 `git push`，
不需要再手動點 `publish.bat`。

## 注意事項

- `台灣篩選器_股期_ver.2.xlsm` 本身**不會**被推上 GitHub（`.gitignore` 已經排除），只有算好的
  `data/*.json` 會上去。如果你不希望連算好的數字都被別人看到，就靠上面的帳號密碼保護。
- 股票資料存成 `data/chunks/chunk_XXXX.js`，每個檔案大約 2MB，是特意切開的（Cloudflare Pages
  單一檔案上限 25MB，全部塞一支檔案會超過上限沒辦法部署）。以後股票數量再增加，只要重新匯出
  （`refresh_data.py` / `update_from_cmoney.bat`）就會自動照這個大小重新切，不用手動處理。
