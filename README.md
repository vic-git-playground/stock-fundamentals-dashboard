# 個股基本面圖表網站

由「台灣篩選器_股期_ver.2.xlsm」的「個股」sheet 匯出而成的靜態網頁，可輸入股票代號查詢 8 張圖：
PE Band、Fwd. PE 走勢（含可調整的目標PE）、PB、月營收、Margin、預估今年/明年EPS趨勢、法人持股比例、PE與Fwd.EPS Y/Y，
外加手動調整股價與評價試算器。

## 使用方式

直接雙擊 `index.html` 用瀏覽器開啟即可。資料放在 `data/chunks/`（切成一個個約 2MB 的檔案，
由 `data/chunks_index.js` 記錄清單），圖表套件放在 `lib/chart.umd.js`，都是本機檔案，不需要網路、
不需要架伺服器、不需要裝 Python。第一次開啟因為要載入全部約 33MB 的資料，搜尋框上方會先顯示
「載入股票資料中...」，通常 1~2 秒後就會變成「已匯出 265 檔股票」，之前搜尋框會先反白不能輸入。

`index.html`、`data/`、`lib/` 這三者要放在同一層資料夾，複製整包資料夾走可以，但不要單獨複製 `index.html`。
如果打開後圖表區塊整片空白，通常是 `lib/chart.umd.js` 沒有跟著移動；畫面上會直接顯示這個錯誤訊息。
`run_server.bat` 是備用方案（需要電腦裝 Python），一般不需要用到。

目前已匯出「統整」sheet 收錄的全部 265 檔股票（`data/chunks/`）。

## 線上分享版

網站也部署在線上（公開網址，沒有登入限制）：
`https://vic-git-playground.github.io/stock-fundamentals-dashboard/`。部署細節、資料更新方式見
`DEPLOY.md`。

## 多重篩選

點右上角的「多重篩選」按鈕展開，規則跟你在 Excel「統整」sheet 裡按「多重篩選」巨集
（`AdvancedFilter`，條件填在 AJ2:BF6）完全一樣：

- **同一組（同一列）裡的多個條件 → 全部都要成立（AND）**
- **不同組（不同列）之間 → 只要符合任一組就會被選出（OR）**，按「＋ 新增一組條件（OR）」可以再加，最多幾組都行
- 條件留空 = 不限制這個欄位

條件寫法：數值欄位可以填 `>10`、`<=25`、`=3`、`<>0`，另外多支援 Excel 沒有的 `5~20`（介於，含端點）；
文字欄位（股票代號／股票名稱／產業名稱）比照 Excel 用「開頭符合」，例如產業填 `電子` 會一併包含
「電子–半導體」「電子–光電」等等。

結果表格點欄位標題可以排序，點股票代號會直接跳去看那一檔的 8 張圖。

資料來源是「統整」sheet 的 A5:AH 那張表（收盤價、Fwd. PE/PB、PE/PB 歷史位階、各期 EPS YoY、
月營收排行、4Q GPM/OPM/NIM、預估殖利率、預估機構數、產業名稱、個股期貨），每次重新匯出時
會由 `tools/export_screener.py` 一併更新到 `data/screener.js`。

## 手動調整股價

查詢一檔股票後，畫面上會有「手動調整目前股價」的輸入框。留空就用 Excel 匯出資料裡最新的收盤價；
填了數字之後：

- PE Band、Fwd. PE 走勢、PB 這三張圖的**最新一點**（股價、Fwd.PE、PB、以及 PE/PB 上中下界）都會用你填的價格重新算一次
- 其餘歷史資料點完全不變
- 這個輸入值只存在你自己瀏覽器的 localStorage 裡（換電腦、清瀏覽器資料、或別人打開這個網站都看不到），不會寫回 Excel 或影響其他人

## 評價試算器

在手動調整股價下面，會顯示這檔股票最新的 Fwd. EPS 與 Book Value (BVPS)。你可以自己填「目標 PE 倍數」
或「目標 PB 倍數」，算出目標價跟預期報酬率。「目前股價」計算基準跟上面的手動調整股價連動：你填了手動價格
就用那個算，沒填就用原始收盤價。

## 資料怎麼更新（Excel 改了之後）

平常（沒有要跑 CMoney 更新，只是你自己在 Excel 裡改完資料）：跟 Claude 說「幫我重新匯出網站資料」，
Claude 會執行：

```
python tools/refresh_data.py "D:\工作區\主要工作檔\台灣篩選器_股期_ver.2.xlsm" data
```

不加 `--codes` 參數時，預設會依照「統整」sheet 目前收錄的股票清單全部重新匯出並覆蓋 `data/*.json`、
`data/chunks/`、`data/manifest.json`。只想更新特定幾檔的話可以用 `--codes 2330,2454,3037`。

如果你是自己手動在 Excel 裡更新完 CMoney 資料（沒有用下面的 `update_from_cmoney.bat`），存檔關閉
Excel 後，直接雙擊 **`refresh_and_publish.bat`**：會重新匯出網站資料，然後自動幫你推送上網，
不會再重複呼叫一次 CMoney 更新。

### 從 CMoney 抓新資料（要在裝了 CMoney 的那台電腦上跑）

這個工作資料夾所在的電腦沒有裝 CMoney，所以下面這個更新流程沒辦法在這裡測試，是照你原本
`process_etf.py` 呼叫 `CMExcel.exe` 的方式寫的：

雙擊 `update_from_cmoney.bat`（要在有裝 CMoney、且這個專案資料夾同步過去的那台電腦上執行），流程是：

1. 呼叫 `CMExcel.exe` 更新 `台灣篩選器_股期_ver.2.xlsm` 裡的 CMoney 報表
2. 存檔並關閉 Excel，釋放檔案鎖定
3. 自動跑 `tools/refresh_data.py` 重新產生網站資料

⚠️ `CMExcel.exe` 的更新參數（`M4||檔案路徑`）是照抄 `process_etf.py` 原本更新 `主動式ETF_all.xlsx`
用的寫法，這個代碼是不是通用於任何 CMoney 活頁簿、還是只對特定報表格式有效，我沒辦法在這裡確認。
第一次在那台電腦跑的時候麻煩你留意一下 xlsm 裡的數字（例如最上面日期）有沒有真的更新，如果沒有，
把正確的呼叫方式回報給我，我再調整 `tools/update_from_cmoney.py` 裡的 `CMONEY_UPDATE_ARG_TEMPLATE`。

## 資料是怎麼算出來的（給有興趣的你）

沒有透過 Excel 重新計算公式，而是直接讀「周收/PB/法人持股/投信持股/外資持股/自營持股/加權Fwd.EPS/
今年Est.EPS/明年Est.EPS/預估今年EPS成長/預估明年EPS成長/Bwd.EPS/Margin/月營收」這些原始資料 sheet
裡該股票的那一列，用跟「個股」sheet 公式一樣的邏輯（滾動三年最高/最低/中位數 PE、PB，月營收3個月均線，
毛利率/營業利益率用最近一季用值往後遞補到月頻率等等）重新算一次，所以不需要 Excel/LibreOffice 開檔重算，
匯出 265 檔股票大約 1~2 分鐘。

「Fwd. PE 走勢」圖的「目前PE」= 最新一週的 Fwd. PE（若你有填手動股價則用手動價格重算）；「目標PE」
可以直接在圖表上方的輸入框調整（預設值 = 目前的中位數PE），純網頁端顯示用，不會影響 Excel 檔案。
