# 網站怎麼上線的

用 **GitHub Pages** 直接把這個資料夾發布成公開網站，沒有另外做登入限制（原本規劃過 Google 帳號
登入閘，後來決定不需要，網站直接公開瀏覽即可）。

- GitHub repo：`https://github.com/vic-git-playground/stock-fundamentals-dashboard`（Public）
- 線上網址：`https://vic-git-playground.github.io/stock-fundamentals-dashboard/`
- `robots.txt` 設定 `Disallow: /`，請搜尋引擎不要收錄，但網址本身仍是公開的，知道網址的人都看得到。

## 之後資料更新怎麼推上網

雙擊資料夾裡的 **`publish.bat`**：把 `git add -A / commit / push` 做完，GitHub Pages 偵測到
新的 commit 會自動重新部署，通常 1 分鐘內網站就會更新。

如果是在有裝 CMoney 的那台電腦，直接跑 **`update_from_cmoney.bat`**：會先更新 Excel 裡的
CMoney 資料，再重新匯出網站資料，最後自動幫你 `git push`（不用再另外跑 publish.bat）。

## 已知問題：這台電腦的 publish.bat 雙擊沒反應

目前還沒排查出原因，最可能是這台 Windows 電腦沒有安裝 Git，或是 Git 沒有加進系統 PATH。
排查方式：開一個命令提示字元（cmd），`cd` 到這個資料夾，手動打 `publish.bat` 執行，
看螢幕上印出的錯誤訊息是什麼，或直接打 `git --version` 確認有沒有裝 Git。把錯誤訊息告訴 Claude
就可以繼續排查。

## 注意事項

- `台灣篩選器_股期_ver.2.xlsm` 本身不會被推上 GitHub（`.gitignore` 已排除），只有算好的
  `data/*.json`／`data/chunks/` 會上去，這些數字理論上任何知道網址的人都拿得到。
- 股票資料存成 `data/chunks/chunk_XXXX.js`，每個約 2MB，GitHub 對單一檔案的限制是 100MB，
  遠遠夠用。
