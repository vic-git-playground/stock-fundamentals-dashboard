"""Streaming helpers to pull one stock's row + header row out of a big worksheet xml
using iterparse so we don't load the whole sheet into memory / regex over it."""
import re, os
from xml.etree import ElementTree as ET
import functools

# 預設指向 tools/ 旁邊的 _xlsm_extract（refresh_data.py 執行時會用實際路徑覆蓋這個值）
BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '_xlsm_extract') + os.sep
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

@functools.lru_cache(maxsize=1)
def shared_strings():
    data = open(BASE + 'xl/sharedStrings.xml', encoding='utf-8').read()
    items = re.findall(r'<si>(.*?)</si>', data, re.S)
    out = []
    for it in items:
        texts = re.findall(r'<t[^>]*>(.*?)</t>', it, re.S)
        out.append(''.join(texts))
    return out

def _cell_value(c):
    t = c.get('t')
    v = c.find(NS + 'v')
    if v is None or v.text is None:
        f = c.find(NS + 'f')
        return None
    raw = v.text
    if t == 's':
        try:
            return shared_strings()[int(raw)]
        except Exception:
            return raw
    if t == 'str' or t == 'e':
        return raw
    try:
        fv = float(raw)
        return int(fv) if fv.is_integer() else fv
    except ValueError:
        return raw

def row_to_dict(row_el):
    d = {}
    for c in row_el.findall(NS + 'c'):
        coord = c.get('r')
        d[coord] = _cell_value(c)
    return d

def col_num(coord):
    m = re.match(r'[A-Z]+', coord)
    col = m.group()
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n

def find_rows(sheetfile, header_rows=(4, 5), code_col='A', code_value=None, max_scan_rows=None):
    """Iterate through a worksheet xml, capture given header rows fully, and
    the first data row (row index > max(header_rows)) whose code_col cell equals code_value (as string, numeric-normalized).
    Returns (headers: {rownum: {coord:val}}, matched_row: {coord:val} or None, matched_rownum)
    """
    path = BASE + 'xl/worksheets/' + sheetfile
    headers = {}
    matched = None
    matched_rn = None
    target = str(code_value).strip() if code_value is not None else None
    scanned = 0
    context = ET.iterparse(path, events=('end',))
    for event, el in context:
        if el.tag == NS + 'row':
            rn = int(el.get('r'))
            if rn in header_rows:
                headers[rn] = row_to_dict(el)
            elif target is not None and matched is None:
                acell = el.find(f"{NS}c[@r='{code_col}{rn}']")
                if acell is not None:
                    val = _cell_value(acell)
                    if val is not None and str(val).strip().split('.')[0] == target:
                        matched = row_to_dict(el)
                        matched_rn = rn
            scanned += 1
            el.clear()
            if matched is not None and len(headers) == len(header_rows):
                break
            if max_scan_rows and scanned > max_scan_rows:
                break
    return headers, matched, matched_rn

def col_letters(coord):
    return re.match(r'[A-Z]+', coord).group()

def row_series(headers_row, data_row, start_col_num=3):
    """Pair up header (date key) and data value for columns >= start_col_num, in column order.
    headers_row: {coord(e.g. 'D4'): date_key}; data_row: {coord(e.g. 'D134'): value} -- match by column letters."""
    data_by_col = {col_letters(coord): v for coord, v in (data_row or {}).items()}
    pairs = []
    for coord, key in headers_row.items():
        col = col_letters(coord)
        cn = col_num(coord)
        if cn < start_col_num:
            continue
        val = data_by_col.get(col)
        pairs.append((cn, key, val))
    pairs.sort(key=lambda x: x[0])
    return [(k, v) for _, k, v in pairs]
