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
import sys, os, zipfile, shutil, json, argparse, importlib

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('xlsm_path')
    ap.add_argument('out_dir')
    ap.add_argument('--extract-dir', default=None)
    ap.add_argument('--codes', default=None, help='逗號分隔的股票代號清單；不指定則匯出統整sheet全部')
    args = ap.parse_args()

    work = os.path.dirname(os.path.abspath(__file__))
    extract_dir = args.extract_dir or os.path.join(work, '_xlsm_extract')
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)
    os.makedirs(extract_dir, exist_ok=True)
    with zipfile.ZipFile(args.xlsm_path) as z:
        z.extractall(extract_dir)

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
    export_stock.SHEETS = {
        '周收': name2file['周收'], 'PB': name2file['PB'], '法人持股': name2file['法人持股'],
        '自營持股': name2file['自營持股'], '投信持股': name2file['投信持股'], '外資持股': name2file['外資持股'],
        '加權Fwd.EPS': name2file['加權Fwd.EPS'], '明年Est.EPS': name2file['明年Est.EPS'],
        '預估明年EPS成長': name2file['預估明年EPS成長'], '今年Est.EPS': name2file['今年Est.EPS'],
        '預估今年EPS成長': name2file['預估今年EPS成長'], 'Bwd.EPS': name2file['Bwd.EPS'],
        'Margin': name2file['Margin'], '月營收': name2file['月營收'],
    }

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

    print(f"完成：成功 {len(ok)} 檔，失敗 {len(fail)} 檔，manifest 共 {result['stocks']} 檔，"
          f"切成 {result['chunks']} 個 chunk 檔案。")
    if fail:
        print('失敗清單:', fail[:20])

if __name__ == '__main__':
    main()
