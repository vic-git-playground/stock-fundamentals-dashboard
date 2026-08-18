"""
重新匯出流程（Claude 在使用者要求「重新匯出」時執行）：
  1. 解壓縮最新的 台灣篩選器_股期_ver.2.xlsm（xlsm 本質上是 zip）
  2. 從「統整」sheet 讀出目前的股票代號清單
  3. 對每檔股票，從各資料 sheet（周收/PB/法人持股.../加權Fwd.EPS/月營收/Margin/Bwd.EPS/Est.EPS...）
     直接抓該股票那一列的歷史數列，重算個股 SHEET 8 張圖表用得到的欄位，寫成 data/<code>.json
  4. 重建 data/manifest.json（供網頁做股票代號/名稱搜尋用）

用法：
  python refresh_data.py <xlsm路徑> <輸出data資料夾> [--codes 2330,2454,...]
  不帶 --codes 時，預設匯出「統整」sheet 目前收錄的全部股票。
"""
import sys, os, stat, time, zipfile, shutil, json, argparse, importlib, tempfile


def force_rmtree(path, retries=3):
    """Windows 上刪暫存資料夾常常失敗（解壓出來的檔案／資料夾帶唯讀屬性，
    或被防毒、搜尋索引程式短暫鎖住）。這裡先把整棵樹的唯讀屬性都拔掉再刪，
    失敗就等一下重試。真的刪不掉也只回傳 False，不會讓整個匯出中斷。"""
    def make_writable(p):
        # 資料夾一定要保留執行權限，否則之後就進不去、反而更刪不掉
        try:
            os.chmod(p, 0o700 if os.path.isdir(p) else 0o600)
        except Exception:
            pass

    def clear_readonly(root):
        make_writable(root)
        # 由上往下走，先讓每層資料夾可進入，才有辦法處理裡面的東西
        for dirpath, dirnames, filenames in os.walk(root, topdown=True):
            for name in dirnames:
                make_writable(os.path.join(dirpath, name))
            for name in filenames:
                make_writable(os.path.join(dirpath, name))

    def on_error(func, p, exc_info):
        try:
            make_writable(os.path.dirname(p))
            make_writable(p)
            func(p)
        except Exception:
            pass

    for i in range(retries):
        if not os.path.exists(path):
            return True
        try:
            clear_readonly(path)
            shutil.rmtree(path, onerror=on_error)
        except Exception:
            pass
        if not os.path.exists(path):
            return True
        time.sleep(0.4 * (i + 1))
    return not os.path.exists(path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('xlsm_path')
    ap.add_argument('out_dir')
    ap.add_argument('--extract-dir', default=None)
    ap.add_argument('--codes', default=None, help='逗號分隔的股票代號清單；不指定則匯出統整sheet全部')
    args = ap.parse_args()

    work = os.path.dirname(os.path.abspath(__file__))

    # 每次都解壓到一個全新的暫存資料夾。
    # 以前是固定用 tools/_xlsm_extract，但在 Windows 上重複刪同一個資料夾很容易踩到
    # 「PermissionError: 存取被拒」（唯讀屬性、或防毒/搜尋索引正好開著裡面的檔案）。
    # 換成每次開新的、跑完再清掉，就完全避開這個問題。
    if args.extract_dir:
        extract_dir = args.extract_dir
        force_rmtree(extract_dir)
        os.makedirs(extract_dir, exist_ok=True)
        temp_extract = False
    else:
        extract_dir = tempfile.mkdtemp(prefix='xlsm_extract_')
        temp_extract = True

    # 順手把舊版留下來的固定資料夾清掉（刪不掉也無所謂，不影響這次執行）
    legacy = os.path.join(work, '_xlsm_extract')
    if os.path.isdir(legacy):
        force_rmtree(legacy)

    try:
        with zipfile.ZipFile(args.xlsm_path) as z:
            z.extractall(extract_dir)
    except Exception as e:
        if temp_extract:
            force_rmtree(extract_dir)
        print(f'解壓縮 Excel 失敗：{e}')
        print('請確認 Excel 檔案已經存檔關閉、沒有被其他程式鎖住，路徑也正確。')
        sys.exit(1)

    try:
        # point the helper modules at the fresh extract
        sys.path.insert(0, work)
        import xhelp2
        xhelp2.BASE = extract_dir + '/'
        xhelp2.shared_strings.cache_clear()

        import export_stock
        export_stock.export.__globals__  # noop, just to ensure module loaded after BASE patch

        from xml.etree import ElementTree as ET
        NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

        # resolve sheet name -> worksheetN.xml (robust to sheet reordering)
        import re
        wb = open(os.path.join(extract_dir, 'xl/workbook.xml'), encoding='utf-8').read()
        rels = open(os.path.join(extract_dir, 'xl/_rels/workbook.xml.rels'), encoding='utf-8').read()
        sheets = re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wb)
        relmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="worksheets/(sheet\d+\.xml)"', rels))
        name2file = {name: relmap[rid] for name, rid in sheets if rid in relmap}
        wanted_sheets = ['周收', 'PB', '法人持股', '自營持股', '投信持股', '外資持股',
                         '加權Fwd.EPS', '明年Est.EPS', '預估明年EPS成長', '今年Est.EPS',
                         '預估今年EPS成長', 'Bwd.EPS', 'Margin', '月營收']
        export_stock.SHEETS = {k: name2file[k] for k in wanted_sheets if k in name2file}

        if args.codes:
            codes = [(c.strip(), None) for c in args.codes.split(',') if c.strip()]
        else:
            # read 統整 sheet code list
            stock_sheet = name2file['統整']
            path = os.path.join(extract_dir, 'xl/worksheets', stock_sheet)
            ss = xhelp2.shared_strings()
            def cellval(c):
                t = c.get('t'); v = c.find(NS + 'v')
                if v is None or v.text is None: return None
                raw = v.text
                if t == 's':
                    try: return ss[int(raw)]
                    except Exception: return raw
                try:
                    f = float(raw); return int(f) if f.is_integer() else f
                except ValueError: return raw
            codes = []
            for event, el in ET.iterparse(path, events=('end',)):
                if el.tag == NS + 'row':
                    rn = int(el.get('r'))
                    if rn > 5:
                        acell = el.find(f"{NS}c[@r='A{rn}']")
                        bcell = el.find(f"{NS}c[@r='B{rn}']")
                        if acell is not None:
                            av = cellval(acell)
                            bv = cellval(bcell) if bcell is not None else None
                            if av is not None:
                                codes.append((str(av), bv))
                    el.clear()

        os.makedirs(args.out_dir, exist_ok=True)
        ok, fail = [], []
        for code, name_hint in codes:
            try:
                data, err = export_stock.export(code)
                if err or data is None:
                    fail.append((code, name_hint, err)); continue
                with open(os.path.join(args.out_dir, f'{code}.json'), 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, separators=(',', ':'), default=str)
                ok.append((code, data.get('name') or name_hint))
            except Exception as e:
                fail.append((code, name_hint, str(e)))

        # rebuild manifest + chunked data files the webpage loads (見 build_web_data.py 說明：
        # 用多個 <script src> chunk 檔取代單一 all_data.js，因為 Cloudflare Pages 單檔上限 25MB)
        import build_web_data
        result = build_web_data.build(args.out_dir)

        # 篩選器用的「統整」sheet 資料表（跟圖表資料是分開的一份）
        try:
            import export_screener
            export_screener.export(args.xlsm_path, args.out_dir)
        except Exception as e:
            print(f'（警告）篩選器資料匯出失敗，網頁的篩選器會沒有資料：{e}')

        print(f"完成：成功 {len(ok)} 檔，失敗 {len(fail)} 檔，manifest 共 {result['stocks']} 檔，"
              f"切成 {result['chunks']} 個 chunk 檔案。")
        if fail:
            print('失敗清單:', fail[:20])
    finally:
        # 用完就把暫存解壓資料夾清掉；萬一系統暫時鎖住刪不掉也不要讓程式報錯
        if temp_extract:
            force_rmtree(extract_dir)

if __name__ == '__main__':
    main()
