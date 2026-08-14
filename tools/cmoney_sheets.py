"""
讀 Excel 裡「CMoney 原始資料表」的共用工具。

這些 sheet 的排法都一樣：
    第 3 列  參考標題   例：20260811收盤價 / 202607單月營收(千) / 2026Q2毛利率(%)
    第 4 列  期別       例：20260811 / 202607 / 202604      （有些表這一列是空的）
    第 5 列  欄位標題
    第 6 列起 資料，A=股票代號、B=股票名稱、C 欄以後每一欄是一期
期別欄一律由新到舊，由左往右。

哪些表是「原始」的、哪些是公式算出來的
--------------------------------------
原始（CMoney 直接給的，要進資料倉）：
    周收 / PB / 法人持股 / 自營持股 / 投信持股 / 外資持股 / 月收 /
    今年Est.EPS / 明年Est.EPS / 預估今年EPS成長 / Bwd.EPS / BVPS / Margin / 月營收 / payout
公式算出來的（不進資料倉，改由 derived.py 用 Python 重算）：
    FwdPE_周      = 周收 ÷ 加權Fwd.EPS
    加權Fwd.EPS   = 今年Est.EPS 與 明年Est.EPS 依月份加權
    預估明年EPS成長 = 明年Est.EPS ÷ 今年Est.EPS - 1
"""
import os
import re
import zipfile
import tempfile
import shutil
import xml.etree.ElementTree as ET

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

HEADER_ROW = 5
FIRST_DATA_ROW = 6

# series 名稱 -> (Excel sheet 名稱, 起始欄, 結束欄或 None 代表到最後)
# Margin 一張表裡塞了 3 個欄位區塊，每塊 40 欄，所以拆成三個 series
SERIES = {
    '周收':            ('周收', 3, None),
    'PB':              ('PB', 3, None),
    '法人持股':         ('法人持股', 3, None),
    '自營持股':         ('自營持股', 3, None),
    '投信持股':         ('投信持股', 3, None),
    '外資持股':         ('外資持股', 3, None),
    '月收':            ('月收', 3, None),
    '今年Est.EPS':     ('今年Est.EPS', 3, None),
    '明年Est.EPS':     ('明年Est.EPS', 3, None),
    '預估今年EPS成長':   ('預估今年EPS成長', 3, None),
    'Bwd.EPS':         ('Bwd.EPS', 3, None),
    'BVPS':            ('BVPS', 3, None),
    '月營收':           ('月營收', 3, None),
    'payout':          ('payout', 3, None),
    'Margin_GPM':      ('Margin', 3, 42),
    'Margin_OPM':      ('Margin', 43, 82),
    'Margin_NIM':      ('Margin', 83, 122),
}

# 期別長這樣：20260811 / 202607 / 2026Q2 / 2026
PERIOD_RE = re.compile(r'^(\d{4}Q[1-4]|\d{8}|\d{6}|\d{4})')


def col_num(ref):
    s = ''.join(ch for ch in ref if ch.isalpha())
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


def unzip_xlsm(xlsm_path, dest=None):
    """把 xlsm 解開（xlsm 本質是 zip）。沒指定 dest 就開暫存資料夾，記得自己刪。"""
    d = dest or tempfile.mkdtemp(prefix='xlsm_')
    with zipfile.ZipFile(xlsm_path) as z:
        z.extractall(d)
    return d


def sheet_file_map(extract_dir):
    """{sheet名稱: sheetN.xml 的完整路徑}"""
    wb = ET.parse(os.path.join(extract_dir, 'xl', 'workbook.xml')).getroot()
    rels = ET.parse(os.path.join(extract_dir, 'xl', '_rels', 'workbook.xml.rels')).getroot()
    rmap = {c.get('Id'): c.get('Target') for c in rels}
    out = {}
    for sh in wb.find(NS + 'sheets'):
        target = rmap.get(sh.get(RNS + 'id'))
        if target:
            out[sh.get('name')] = os.path.join(extract_dir, 'xl', target.replace('/', os.sep))
    return out


def load_shared_strings(extract_dir):
    path = os.path.join(extract_dir, 'xl', 'sharedStrings.xml')
    if not os.path.exists(path):
        return []
    return [''.join(t.text or '' for t in si.iter(NS + 't'))
            for si in ET.parse(path).getroot()]


def _cell_text(c, ss):
    v = c.find(NS + 'v')
    if v is None or v.text is None:
        return None
    if c.get('t') == 's':
        try:
            return ss[int(v.text)]
        except (ValueError, IndexError):
            return None
    return v.text


def read_series(extract_dir, ss, sheets, series):
    """讀一個 series，回傳 (names, data, periods)，格式跟 archive_store 一致。

      names   : {code: 股票名稱}
      data    : {code: {period: float}}
      periods : [period, ...] 由新到舊（就是 Excel 欄位的左到右順序）
    """
    sheet_name, col_from, col_to = SERIES[series]
    path = sheets.get(sheet_name)
    if not path or not os.path.exists(path):
        raise RuntimeError(f'找不到 sheet「{sheet_name}」（series={series}）')

    row3, row4, row5 = {}, {}, {}
    names, data = {}, {}
    col_period = {}

    for _, el in ET.iterparse(path, events=('end',)):
        if el.tag != NS + 'row':
            continue
        rn = int(el.get('r'))

        if rn in (3, 4, 5):
            target = {3: row3, 4: row4, 5: row5}[rn]
            for c in el:
                cn = col_num(re.sub(r'[0-9]', '', c.get('r')))
                if cn < col_from or (col_to and cn > col_to):
                    continue
                t = _cell_text(c, ss)
                if t is not None:
                    target[cn] = t
            el.clear()
            if rn == 5:
                # 期別優先用第 4 列；那一列沒有（或不是期別格式）就從標題文字前面解析出來
                for cn in sorted(set(row3) | set(row4) | set(row5)):
                    p = None
                    raw4 = str(row4.get(cn, '')).strip()
                    m = PERIOD_RE.match(raw4)
                    if m:
                        p = m.group(1)
                    else:
                        for src in (row5, row3):
                            m = PERIOD_RE.match(str(src.get(cn, '')).strip())
                            if m:
                                p = m.group(1)
                                break
                    if p:
                        col_period[cn] = p
            continue

        if rn < FIRST_DATA_ROW:
            el.clear()
            continue

        code = name = None
        vals = {}
        for c in el:
            cn = col_num(re.sub(r'[0-9]', '', c.get('r')))
            if cn == 1:
                code = _cell_text(c, ss)
            elif cn == 2:
                name = _cell_text(c, ss)
            elif cn in col_period:
                t = _cell_text(c, ss)
                if t is None or t == '':
                    continue
                try:
                    vals[col_period[cn]] = float(t)
                except ValueError:
                    pass
        el.clear()

        if code is None:
            continue
        code = str(code).strip()
        # 台股代號是 4~6 位數字；表格最後常有 0 或空白的殘留列，濾掉
        if not code or not code.isdigit() or not (4 <= len(code) <= 6):
            continue
        names[code] = (name or '').strip()
        data[code] = vals

    # 依欄位順序（左到右＝新到舊）保留期別順序，重複的只留第一次出現的
    periods, seen = [], set()
    for cn in sorted(col_period):
        p = col_period[cn]
        if p not in seen:
            seen.add(p)
            periods.append(p)
    return names, data, periods


def read_all(xlsm_path, series_list=None, extract_dir=None, keep_extract=False):
    """一次讀多個 series，回傳 {series: (names, data, periods)}。"""
    series_list = series_list or list(SERIES.keys())
    tmp = extract_dir or tempfile.mkdtemp(prefix='xlsm_')
    created = extract_dir is None
    try:
        if created or not os.path.exists(os.path.join(tmp, 'xl', 'workbook.xml')):
            unzip_xlsm(xlsm_path, tmp)
        ss = load_shared_strings(tmp)
        sheets = sheet_file_map(tmp)
        out = {}
        for s in series_list:
            out[s] = read_series(tmp, ss, sheets, s)
        return out
    finally:
        if created and not keep_extract:
            shutil.rmtree(tmp, ignore_errors=True)
