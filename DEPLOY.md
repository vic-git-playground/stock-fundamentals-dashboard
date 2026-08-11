# 把網站放到網路上分享（Google 帳號登入限制）

方案：**GitHub Pages**（公開的靜態網站託管）＋ **Google Identity Services 前端登入閘**。
做法跟你另一個「券商報告動態追蹤」網站一樣：登入畫面用 Google 帳號驗證，email 不在白名單裡就
擋在登入畫面外。

## 跟 Cloudflare Access 方案的差異（重要，務必了解）

這個做法只在「網頁前端」擋畫面，不是伺服器端真的擋存取：

- GitHub Pages 是公開託管，任何人只要知道網址，理論上都可以直接下載頁面或資料檔本身
  （例如 `data/chunks/chunk_0001.js`），不需要真的通過 Google 登入。
- Google 登入只是「打開頁面後，用 JS 判斷要不要顯示內容」，擋得住不知道網址、隨手點進來的人，
  擋不住刻意想抓資料的人。
- 因為要用 GitHub Pages 的免費方案，這個 repo 必須設成 **Public**（原本是 Private）。

如果之後想換成真正的伺服器端保護（白名單外的人連資料檔案內容都拿不到），可以改用先前規劃的
Cloudflare Pages + Zero Trust Access 方案，需要另外申請/登入 Cloudflare 帳號。

## 目前的設定

- Google OAuth Client 沿用你「券商報告動態追蹤」網站同一組（同一個 Google Cloud 專案），
  Client ID：`895601722476-kikjqnb5afmhbgktnb9q0hprlnhdutrc.apps.googleusercontent.com`，
  只是多把這個新網站的網址加進「Authorized JavaScript origins」清單。
- 白名單寫在 `index.html` 裡的 `ALLOWLIST` 陣列（搜尋 `AUTH_CONFIG_START`），目前是：
  - tangershen@gmail.com
  - vicshen.jva@gmail.com
- GitHub repo：`https://github.com/vic-git-playground/stock-fundamentals-dashboard`
  （改成 Public 後才能開 GitHub Pages）
- 部署網址：`https://vic-git-playground.github.io/stock-fundamentals-dashboard/`

## 之後怎麼加人/ 換人

編輯 `index.html` 裡的 `ALLOWLIST = [...]`，加入或刪除 email，存檔後跑 `publish.bat`
（或請 Claude 幫你改），GitHub Pages 偵測到新的 commit 會自動重新部署，通常 1 分鐘內生效。

## 之後資料更新怎麼推上網

雙擊資料夾裡的 **`publish.bat`**：把 `git add -A / commit / push` 做完，GitHub Pages 偵測到
GitHub 有新的 commit 會自動重新部署，通常 1 分鐘內網站就會更新（Google 登入白名單不受影響）。

如果是在有裝 CMoney 的那台電腦跑 `update_from_cmoney.bat`，只要那台電腦也接了同一個 git remote，
它跑完會自動幫你 `git push`。

## 注意事項

- `台灣篩選器_股期_ver.2.xlsm` 本身不會被推上 GitHub（`.gitignore` 已排除），只有算好的 `data/*.json`
  會上去，但因為 repo 是 Public，這些算好的數字理論上任何人都拿得到（見上方「差異」說明）。
- `robots.txt` 已設定 `Disallow: /`，會請搜尋引擎不要收錄，但不是存取限制。
- 股票資料存成 `data/chunks/chunk_XXXX.js`，每個約 2MB，GitHub 對單一檔案的限制是 100MB，
  不用像 Cloudflare Pages 那樣特別切檔，但既有的切檔架構保留著也沒有壞處。
