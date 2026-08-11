# 把網站放到網路上分享（Google 帳號登入限制）

方案：GitHub 放程式碼（私有倉庫）＋ Cloudflare Pages 部署 ＋ Cloudflare Zero Trust Access
（用 Google 帳號登入、白名單限制哪些 email 可以進來，跟券商報告追蹤系統常見做法一樣）。

這份文件是給我（Claude）用瀏覽器操作時的步驟依據，你不用自己動手，但如果哪一步需要你本人做
（例如登入、同意 OAuth 授權畫面），我會在當下提示你。

## 需要你先準備好的東西

1. 連上 Claude in Chrome 擴充套件，並且瀏覽器裡已經登入：
   - 你要拿來當「東家帳號」的 GitHub 帳號
   - Cloudflare 帳號（沒有的話我會帶你當場註冊，免費）
   - 你的 Google 帳號（用來設定 OAuth，之後這個網站登入畫面也會用 Google 帳號）
2. 允許登入這個網站的 Google email 名單（可以之後在 Cloudflare 後台隨時增減，不用重新部署）

## 我會依序做的事

### 1. GitHub：建立私有倉庫、把程式碼推上去
- 建一個新的 **Private** 倉庫（裡面的數字是從付費的 CMoney 資料算出來的，不適合設成 Public）
- 這個資料夾本機已經 `git init`、commit 好了，只差 `git remote add` + `git push`

### 2. Cloudflare Pages：部署靜態網站
- Workers & Pages -> Create -> Pages -> Connect to Git -> 選剛剛的倉庫
- 純靜態網站，不需要 build command
- 部署完會拿到一個 `https://xxx.pages.dev` 網址

### 3. Google Cloud Console：註冊 OAuth Client（讓 Cloudflare 可以用「登入 Google」）
- 建立一個新專案，設定 OAuth 同意畫面（External，任何 Google 帳號都能用來登入，不需要你有
  Google Workspace）
- 建立 OAuth Client（Web application），填入 Cloudflare 那邊會給的兩個網址：
  - Authorized JavaScript origins: `https://<你的team名稱>.cloudflareaccess.com`
  - Authorized redirect URIs: `https://<你的team名稱>.cloudflareaccess.com/cdn-cgi/access/callback`
- 拿到 Client ID 和 Client secret

### 4. Cloudflare Zero Trust：接上 Google 登入 + 設定白名單
- 第一次使用 Zero Trust 需要設定一個「team 名稱」（就是上面网址裡的 `<你的team名稱>`）
- Zero Trust -> Integrations -> Identity providers -> Add -> Google，貼上第 3 步拿到的
  Client ID / Client secret
- Zero Trust -> Access -> Applications -> Add an application -> Self-hosted，網域填
  `xxx.pages.dev`（第 2 步拿到的網址）
- 設定 Access policy：Allow，Include 條件選 **Emails**，貼上允許登入的 Google email 名單
- 儲存後，之後任何人開這個網址都要先用 Google 帳號登入，登入的帳號不在白名單裡會被擋下來

### 5. 之後怎麼加人/ 換人
不用重新部署網站，直接到 Cloudflare Zero Trust -> Access -> Applications -> 這個應用程式 -> 編輯
policy，把 email 加進去或刪掉即可，幾秒鐘生效。

## 之後資料更新怎麼推上網

雙擊資料夾裡的 **`publish.bat`**：把 `git add -A / commit / push` 做完，Cloudflare Pages 偵測到
GitHub 有新的 commit 會自動重新部署，通常 1~2 分鐘網站就會更新（Google 登入限制不受影響，還是一樣要
白名單裡的帳號才能看）。

如果是在有裝 CMoney 的那台電腦跑 `update_from_cmoney.bat`，只要那台電腦也接了同一個 git remote，
它跑完會自動幫你 `git push`。

## 注意事項

- `台灣篩選器_股期_ver.2.xlsm` 本身不會被推上 GitHub（`.gitignore` 已排除），只有算好的 `data/*.json`
  會上去。
- 股票資料存成 `data/chunks/chunk_XXXX.js`，每個約 2MB（Cloudflare Pages 單一檔案上限 25MB，
  所以特意切開），以後股票數量再增加，重新匯出時會自動照這個大小重新切。
