"""
把「統整」sheet 的篩選用資料表（A5:AH，row5 是標題、row6 起是資料）匯出成網站用的
data/screener.js。

對應 Excel 裡「多重篩選」巨集的做法：
    Range("A5:AH2500").AdvancedFilter CriteriaRange:=Range("AJ2:BF6"), CopyToRange:=Range("AJ9:BF9")
也就是 AJ2:BF2 是條件欄位標題、AJ3:BF6 是最多 4 組條件（同一列的條件之間是 AND，
不同列之間是 OR）。網頁端的篩選器會用同樣的規則重現。

用法：
    python tools/export_screener.py <xlsm路徑> <輸出資料夾>
一般不用單獨執行，refresh_data.py 會自動呼叫。
"""
import os
import re
import sys
import json
import zipfile
import datetime
import shutil
import tempfile
import xml.etree.ElementTree as ET

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

HEADER_ROW = 5      # 統整 sheet 的資料表標題列
FIRST_DATA_ROW = 6  # 資料從第 6 列開始

# 要匯出的欄位：(Excel 欄位字母, 輸出欄位 key, 顯示名稱, 型別)
# 型別 num = 數值可用 > < >= <= = <> 篩選；text = 文字，比照 Excel AdvancedFilter 用「開頭符合」
COLUMNS = [
    ('A',  'code',        '股票代號',        'text'),
    ('B',  'name',        '股票名稱',        'text'),
    ('C',  'price',       '收盤價',          'num'),
    ('N',  'fwd_pe',      'Fwd. PE',        'num'),
    ('O',  'fwd_pb',      'Fwd. PB',        'num'),
    ('P',  'pe_rank',     'PE 歷史位階',     'num'),
    ('Q',  'pb_rank',     'PB 歷史位階',     'num'),
    ('R',  'eps_yoy_4q',  '累積近4Q EPS YoY', 'num'),
    ('S',  'eps_yoy_cur', '今年Est. EPS YoY', 'num'),
    ('T',  'eps_yoy_nxt', '明年Est. EPS YoY', 'num'),
    ('U',  'rev_rank',    '10年月營收排行',   'num'),
    ('V',  'gpm_4q',      '4Q GPM',         'num'),
    ('W',  'opm_4q',      '4Q OPM',         'num'),
    ('X',  'nim_4q',      '4Q NIM',         'num'),
    ('AD', 'div_yield',   '預估殖利率',      'num'),
    ('AE', 'analysts',    '預估機構數',      'num'),
    ('AF', 'industry',    '產業名稱',        'text'),
    ('AH', 'has_future',  '個股期貨',        'num'),
]


def col_num(ref):
    s = ''.join(ch for ch in ref if ch.isalpha())
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


def find_sheet_file(extract_dir, sheet_name):
    wb = ET.parse(os.path.join(extract_dir, 'xl', 'workbook.xml')).getroot()
    rels = ET.parse(os.path.join(extract_dir, 'xl', '_rels', 'workbook.xml.rels')).getroot()
    rmap = {c.get('Id'): c.get('Target') for c in rels}
    for sh in wb.find(NS + 'sheets'):
        if sh.get('name') == sheet_name:
            target = rmap.get(sh.get(RNS + 'id'))
            return os.path.join(extract_dir, 'xl', target.replace('/', os.sep))
    raise RuntimeError(f'找不到 sheet: {sheet_name}')


def load_shared_strings(extract_dir):
    path = os.path.join(extract_dir, 'xl', 'sharedStrings.xml')
    if not os.path.exists(path):
        return []
    out = []
    for si in ET.parse(path).getroot():
        out.append(''.join(t.text or '' for t in si.iter(NS + 't')))
    return out


def export(xlsm_path, out_dir):
    tmp = tempfile.mkdtemp(prefix='screener_')
    try:
        with zipfile.ZipFile(xlsm_path) as z:
            z.extractall(tmp)

        ss = load_shared_strings(tmp)
        sheet_path = find_sheet_file(tmp, '統整')

        wanted = {col_num(letter): (key, typ) for letter, key, _, typ in COLUMNS}
        rows = []

        for _, el in ET.iterparse(sheet_path, events=('end',)):
            if el.tag != NS + 'row':
                continue
            rn = int(el.get('r'))
            if rn < FIRST_DATA_ROW:
                el.clear()
                continue

            rec = {}
            for c in el:
                ref = re.sub(r'[0-9]', '', c.get('r'))
                cn = col_num(ref)
                if cn not in wanted:
                    continue
                key, typ = wanted[cn]
                v = c.find(NS + 'v')
                if v is None or v.text in (None, ''):
                    continue
                val = v.text
                if c.get('t') == 's':
                    val = ss[int(val)]
                elif c.get('t') == 'str':
                    pass
                if typ == 'num':
                    try:
                        val = float(val)
                    except ValueError:
                        continue
                    if val != val:  # NaN
                        continue
                else:
                    val = str(val).strip()
                    if val == '':
                        continue
                rec[key] = val
            el.clear()

            # 沒有股票代號的列直接跳過（表格底下的空白列）
            if not rec.get('code'):
                continue
            rec['code'] = str(rec['code']).strip()
            rows.append(rec)


        result = {
            'generated_at': datetime.datetime.now().isoformat(),
            'columns': [
                {'key': key, 'label': label, 'type': typ}
                for _, key, label, typ in COLUMNS
            ],
            'rows': rows,
        }

        os.makedirs(out_dir, exist_ok=True)
        js_path = os.path.join(out_dir, 'screener.js')
        with open(js_path, 'w', encoding='utf-8') as f:
            f.write('window.SCREENER = ')
            json.dump(result, f, ensure_ascii=False, separators=(',', ':'))
            f.write(';\n')

        print(f'篩選器資料已匯出：{len(rows)} 檔股票 -> {js_path}')
        return result
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main():
    if len(sys.argv) < 3:
        print('用法: python tools/export_screener.py <xlsm路徑> <輸出資料夾>')
        sys.exit(1)
    export(sys.argv[1], sys.argv[2])


if __name__ == '__main__':
    main()
