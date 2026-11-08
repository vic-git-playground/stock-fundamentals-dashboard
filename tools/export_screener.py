"""
匯出網站「多重篩選」用的資料表 data/screener.js。

【新版做法：不再讀「統整」L~AH 欄的公式結果】
舊版是直接讀「統整」sheet A5:AH 的快取值。但 L~AH 是手動往下拉的 Excel 公式，
股票清單變長之後公式沒跟著拉，後面的股票就整排空白（2026-11 就是 525 檔裡只有前 265 檔有值）。

新版只從「統整」讀 CMoney 外掛直接抓下來的原始值（A~H 欄：代號、名稱、收盤價、今年/明年
預估 EPS、每股淨值、預測機構數、產業名稱），其餘所有指標都由這支程式自己從各原始資料 sheet
用「股票代號」對應後計算，不依賴列號對齊，也不依賴 Excel 有沒有重算、公式有沒有拉到底。

各欄位算法（對應原本「統整」的公式，括號內是原本的欄位）：
  預估配息 (L)   = 今年EPS × 平均(payout 近三年現金股利發放率) / 100
  加權EPS  (M)   = (m-1)/12 × 明年EPS + (1-(m-1)/12) × 今年EPS；明年空白就用今年
                   m = 「今年Est.EPS」最新一欄的月份（原本是 MONTH(TODAY())，改成跟資料月份一致）
  Fwd. PE  (N)   = 收盤價 / 加權EPS
  Fwd. PB  (O)   = 收盤價 / (每股淨值 + 今年EPS − 預估配息)
  PE 歷史位階 (P) = Fwd. PE 在該股「每週 Fwd. PE」歷史序列中的 PERCENTRANK × 100
                   每週 Fwd. PE = 周收 / 該週所屬月份的加權 EPS（同「FwdPE_周」sheet 的算法）
  PE 近3年/近5年位階 = 同上，但歷史序列只取最新一週往回 3 / 5 年內的週資料（新增欄位）
  PB 歷史位階 (Q) = Fwd. PB 在「PB」sheet 該股每週 PB 歷史序列中的 PERCENTRANK × 100
  累積近4Q EPS YoY (R) = Bwd.EPS 最新一個季底欄 vs 去年同季 −1（×100）
  今年Est. EPS YoY (S) = 預估今年EPS成長 最新一欄
  明年Est. EPS YoY (T) = 明年EPS / 今年EPS − 1（×100）；今年EPS<0 或明年空白則不給值
  10年月營收排行 (U)    = 最新單月營收在月營收 sheet 全部月份中的名次（1 = 最高）；
                         最新月還沒公布就用上個月
  4Q GPM / OPM / NIM (V/W/X) = Margin sheet 三個區塊（毛利率、營業利益率、稅前純益率）
                         各自「最近 4 個有資料的季度」平均
  預估殖利率 (AD) = 預估配息 / 收盤價 × 100
  個股期貨 (AH)   = 股票代號有沒有在「個股期貨」sheet C 欄

用法：
    python tools/export_screener.py <xlsm路徑> <輸出資料夾>
一般不用單獨執行，refresh_data.py 會自動呼叫。
"""
import os
import re
import bisect
import sys
import json
import zipfile
import datetime
import shutil
import tempfile
import xml.etree.ElementTree as ET

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

FIRST_DATA_ROW = 6  # 各 sheet 資料都從第 6 列開始

# 網頁端吃的欄位定義（key 不能改，index.html 靠這些 key 顯示與篩選）
COLUMNS = [
    ('code',        '股票代號',        'text'),
    ('name',        '股票名稱',        'text'),
    ('price',       '收盤價',          'num'),
    ('fwd_pe',      'Fwd. PE',        'num'),
    ('fwd_pb',      'Fwd. PB',        'num'),
    ('pe_rank',     'PE 歷史位階',     'num'),
    ('pe_rank_3y',  'PE 近3年位階',    'num'),
    ('pe_rank_5y',  'PE 近5年位階',    'num'),
    ('pb_rank',     'PB 歷史位階',     'num'),
    ('eps_yoy_4q',  '累積近4Q EPS YoY', 'num'),
    ('eps_yoy_cur', '今年Est. EPS YoY', 'num'),
    ('eps_yoy_nxt', '明年Est. EPS YoY', 'num'),
    ('rev_rank',    '10年月營收排行',   'num'),
    ('gpm_4q',      '4Q GPM',         'num'),
    ('opm_4q',      '4Q OPM',         'num'),
    ('nim_4q',      '4Q NIM',         'num'),
    ('div_yield',   '預估殖利率',      'num'),
    ('analysts',    '預估機構數',      'num'),
    ('industry',    '產業名稱',        'text'),
    ('has_future',  '個股期貨',        'num'),
]


# ---------------------------------------------------------------- xlsx 讀取工具

def col_num(ref):
    n = 0
    for ch in ref:
        if ch.isalpha():
            n = n * 26 + ord(ch) - 64
    return n


def sheet_files(extract_dir):
    wb = ET.parse(os.path.join(extract_dir, 'xl', 'workbook.xml')).getroot()
    rels = ET.parse(os.path.join(extract_dir, 'xl', '_rels', 'workbook.xml.rels')).getroot()
    rmap = {c.get('Id'): c.get('Target') for c in rels}
    out = {}
    for sh in wb.find(NS + 'sheets'):
        target = rmap.get(sh.get(RNS + 'id'))
        if target:
            target = target.lstrip('/')
            if target.startswith('xl/'):
                target = target[3:]
            out[sh.get('name')] = os.path.join(extract_dir, 'xl', target.replace('/', os.sep))
    return out


def load_shared_strings(extract_dir):
    path = os.path.join(extract_dir, 'xl', 'sharedStrings.xml')
    if not os.path.exists(path):
        return []
    out = []
    for si in ET.parse(path).getroot():
        out.append(''.join(t.text or '' for t in si.iter(NS + 't')))
    return out


def norm_code(v):
    """股票代號統一成字串：1101 / 1101.0 / '1101' 都變成 '1101'。"""
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    if s.endswith('.0') and s[:-2].isdigit():
        s = s[:-2]
    return s or None


def read_sheet(path, ss, header_rows=(), codes=None, max_col=None):
    """串流讀一張 sheet。

    回傳 (headers, rows)：
      headers = {列號: {欄號: 值}}           （只收 header_rows 指定的列）
      rows    = {股票代號: {欄號: 值}}       （第 6 列起、A 欄是代號的列；只收數值與文字）
    codes 有給的話只保留這些代號，省記憶體。
    """
    headers, rows = {}, {}
    for _, el in ET.iterparse(path, events=('end',)):
        if el.tag != NS + 'row':
            continue
        rn = int(el.get('r'))
        want_header = rn in header_rows
        if rn < FIRST_DATA_ROW and not want_header:
            el.clear()
            continue
        vals = {}
        for c in el:
            v = c.find(NS + 'v')
            if v is None or v.text is None or v.text == '':
                continue
            cn = col_num(c.get('r'))
            if max_col and cn > max_col:
                continue
            t = c.get('t')
            raw = v.text
            if t == 's':
                val = ss[int(raw)]
            elif t == 'e':
                continue                      # #N/A 之類的錯誤值當作沒資料
            elif t in ('str', 'inlineStr'):
                val = raw
            else:
                try:
                    val = float(raw)
                except ValueError:
                    val = raw
            vals[cn] = val
        el.clear()
        if want_header:
            headers[rn] = vals
            continue
        code = norm_code(vals.get(1))
        if not code or code == '0000':
            continue
        if codes is not None and code not in codes:
            continue
        if code not in rows:                  # 同代號重複出現時以第一列為準（同 VLOOKUP）
            rows[code] = vals
    return headers, rows


def num(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v) if v == v else None
    if isinstance(v, str):
        try:
            f = float(v.replace(',', ''))
            return f if f == f else None
        except ValueError:
            return None
    return None


def key_int(v):
    """header 上的期別（20261106 / 202611 / '202611'）轉成 int。"""
    if v is None:
        return None
    try:
        return int(float(str(v).strip()))
    except ValueError:
        m = re.match(r'(\d{6,8})', str(v))
        return int(m.group(1)) if m else None


# ---------------------------------------------------------------- Excel 函數

def excel_percentrank(values, x, significance=3):
    """Excel PERCENTRANK(array, x)（= PERCENTRANK.INC），結果無條件捨去到 3 位小數。
    x 不在陣列裡時於相鄰兩值之間線性內插。"""
    arr = sorted(values)
    n = len(arr)
    if n < 2 or x < arr[0] or x > arr[-1]:
        return None

    i = bisect.bisect_left(arr, x)              # 嚴格小於 x 的個數
    if i < n and arr[i] == x:
        r = i / (n - 1)
    else:
        # 內插：下方相鄰值若有重複，Excel 取它「最後一個」位置（已用 265 檔實際結果比對驗證）
        x0, x1 = arr[i - 1], arr[i]
        r0 = (bisect.bisect_right(arr, x0) - 1) / (n - 1)
        r1 = i / (n - 1)
        r = r0 + (x - x0) / (x1 - x0) * (r1 - r0)
    f = 10 ** significance
    return int(r * f + 1e-9) / f


def position_pct(series, x):
    """統整 P/Q 欄：>= 最大值 → 100，<= 最小值 → 0，其餘 PERCENTRANK × 100。"""
    vals = [v for v in series if v is not None]
    if x is None or not vals:
        return None
    if x >= max(vals):
        return 100.0
    if x <= min(vals):
        return 0.0
    pr = excel_percentrank(vals, x)
    return None if pr is None else pr * 100


def weighted_eps(ty, ny, month):
    """加權 EPS：明年空白就用今年；否則 (m-1)/12*明年 + (1-(m-1)/12)*今年。"""
    if ty is None:
        return None
    if ny is None:
        return ty
    w = (month - 1) / 12
    return w * ny + (1 - w) * ty


def avg(vals):
    vals = [v for v in vals if v is not None]
    return sum(vals) / len(vals) if vals else None


def safe_div(a, b):
    if a is None or b is None or b == 0:
        return None
    return a / b


# ---------------------------------------------------------------- 主計算

def compute(extract_dir, log=print):
    ss = load_shared_strings(extract_dir)
    files = sheet_files(extract_dir)

    def need(name):
        if name not in files:
            raise RuntimeError(f'找不到 sheet: {name}')
        return files[name]

    # 1) 統整：只讀 A~I（CMoney 外掛直接抓的原始值，不是公式）
    _, base = read_sheet(need('統整'), ss, max_col=9)
    codes_order = []
    for code, v in base.items():
        codes_order.append(code)
    code_set = set(codes_order)
    log(f'  統整股票清單 {len(codes_order)} 檔')

    def load(name, header_rows=(4,)):
        h, r = read_sheet(need(name), ss, header_rows=header_rows, codes=code_set)
        log(f'  讀取 {name}：{len(r)} 檔')
        return h, r

    h_ty, r_ty = load('今年Est.EPS')
    h_ny, r_ny = load('明年Est.EPS')
    h_wk, r_wk = load('周收')
    h_pb, r_pb = load('PB')
    h_bwd, r_bwd = load('Bwd.EPS')
    h_gty, r_gty = load('預估今年EPS成長')
    h_rev, r_rev = load('月營收')
    h_mg, r_mg = load('Margin')
    _, r_pay = load('payout', header_rows=())
    # 個股期貨：C 欄是標的代號（序號在 B 欄），整欄掃一遍
    fut_codes = set()
    if '個股期貨' in files:
        for _, el in ET.iterparse(files['個股期貨'], events=('end',)):
            if el.tag == NS + 'c' and re.match(r'C\d+$', el.get('r') or ''):
                v = el.find(NS + 'v')
                if v is not None and v.text:
                    raw = ss[int(v.text)] if el.get('t') == 's' else v.text
                    f = num(raw)
                    if f is not None and f.is_integer():
                        fut_codes.add(str(int(f)))
            elif el.tag == NS + 'row':
                el.clear()

    # ---- 各 header：欄號 -> 期別
    def hdr(h):
        return {cn: key_int(v) for cn, v in h.get(4, {}).items() if cn >= 3 and key_int(v)}

    ty_cols, ny_cols = hdr(h_ty), hdr(h_ny)
    wk_cols, pb_cols = hdr(h_wk), hdr(h_pb)
    bwd_cols, gty_cols, rev_cols = hdr(h_bwd), hdr(h_gty), hdr(h_rev)
    mg_cols = hdr(h_mg)

    latest_ty_col = min(ty_cols)                 # C 欄 = 最新月
    data_month = ty_cols[latest_ty_col] % 100    # 例：202611 -> 11
    log(f'  加權 EPS 使用月份 m = {data_month}（{ty_cols[latest_ty_col]}）')

    # 今年/明年 Est. EPS：期別(yyyymm) -> 欄號
    ty_ym2col = {ym: cn for cn, ym in ty_cols.items()}
    ny_ym2col = {ym: cn for cn, ym in ny_cols.items()}

    # Bwd.EPS：最新的季底欄（月份是 3/6/9/12 的第一欄）與去年同季
    bwd_ym2col = {}
    for cn in sorted(bwd_cols):
        bwd_ym2col.setdefault(bwd_cols[cn], cn)
    q_cols = [cn for cn in sorted(bwd_cols) if bwd_cols[cn] % 100 in (3, 6, 9, 12)]
    bwd_latest_col = q_cols[0] if q_cols else None
    bwd_prev_col = bwd_ym2col.get(bwd_cols[bwd_latest_col] - 100) if bwd_latest_col else None

    gty_latest_col = min(gty_cols) if gty_cols else None
    rev_sorted_cols = sorted(rev_cols)

    # Margin：同一季別由左到右出現三次 = 毛利率 / 營業利益率 / 稅前純益率
    mg_blocks = [[], [], []]
    seen = {}
    for cn in sorted(mg_cols):
        ym = mg_cols[cn]
        k = seen.get(ym, 0)
        seen[ym] = k + 1
        if k < 3:
            mg_blocks[k].append(cn)          # 每個區塊內由新到舊

    # 近 3 年 / 近 5 年位階的起始日：以周收最新一週往回推 3 / 5 年
    latest_wk = max(wk_cols.values())
    cutoff_3y = latest_wk - 30000            # 20261106 -> 20231106
    cutoff_5y = latest_wk - 50000

    rows = []
    for code in codes_order:
        b = base[code]
        name = b.get(2)
        price = num(b.get(3))
        eps_ty = num(b.get(4))
        eps_ny = num(b.get(5))
        bvps = num(b.get(6))
        analysts = num(b.get(7))
        industry = b.get(8)

        # 預估配息 / 加權 EPS / Fwd. PE / Fwd. PB
        pay = r_pay.get(code, {})
        payout_avg = avg([num(pay.get(c)) for c in (3, 4, 5)])
        div = eps_ty * payout_avg / 100 if (eps_ty is not None and payout_avg is not None) else None
        w_eps = weighted_eps(eps_ty, eps_ny, data_month)
        fwd_pe = safe_div(price, w_eps)
        fwd_pb = None
        if price is not None and bvps is not None and eps_ty is not None and div is not None:
            fwd_pb = safe_div(price, bvps + eps_ty - div)

        # PE 歷史位階：每週 Fwd. PE 序列
        wk = r_wk.get(code, {})
        ty_row, ny_row = r_ty.get(code, {}), r_ny.get(code, {})
        pe_hist = []                         # [(yyyymmdd, PE)]
        for cn, ymd in wk_cols.items():
            px = num(wk.get(cn))
            ym = ymd // 100
            if px is None:
                continue
            t = num(ty_row.get(ty_ym2col.get(ym)))
            n = num(ny_row.get(ny_ym2col.get(ym)))
            e = weighted_eps(t, n, ym % 100)
            pe = safe_div(px, e)
            if pe is not None:
                pe_hist.append((ymd, pe))
        pe_hist.sort(reverse=True)           # 由新到舊
        pe_rank = position_pct([v for _, v in pe_hist], fwd_pe)
        pe_rank_3y = position_pct([v for d, v in pe_hist if d >= cutoff_3y], fwd_pe)
        pe_rank_5y = position_pct([v for d, v in pe_hist if d >= cutoff_5y], fwd_pe)

        # PB 歷史位階
        pbr = r_pb.get(code, {})
        pb_hist = [num(pbr.get(cn)) for cn in pb_cols]
        pb_rank = position_pct(pb_hist, fwd_pb)

        # EPS YoY
        bw = r_bwd.get(code, {})
        eps_yoy_4q = None
        if bwd_latest_col and bwd_prev_col:
            cur, prev = num(bw.get(bwd_latest_col)), num(bw.get(bwd_prev_col))
            r = safe_div(cur, prev)
            eps_yoy_4q = 100 * (r - 1) if r is not None else None
        eps_yoy_cur = num(r_gty.get(code, {}).get(gty_latest_col)) if gty_latest_col else None
        eps_yoy_nxt = None
        if eps_ny is not None and eps_ty is not None and eps_ty >= 0:
            r = safe_div(eps_ny, eps_ty)
            eps_yoy_nxt = 100 * (r - 1) if r is not None else None

        # 10 年月營收排行
        rv = r_rev.get(code, {})
        rev_vals = [num(rv.get(cn)) for cn in rev_sorted_cols]
        rev_rank = None
        target = next((v for v in rev_vals[:2] if v is not None), None)   # 最新月，沒有就上個月
        if target is not None:
            rev_rank = 1 + sum(1 for v in rev_vals if v is not None and v > target)

        # 4Q 利潤率：各區塊最近 4 個有資料的季度
        mg = r_mg.get(code, {})
        def last4(block):
            vals = [num(mg.get(cn)) for cn in block]
            vals = [v for v in vals if v is not None][:4]
            return avg(vals)
        gpm, opm, npm = (last4(blk) for blk in mg_blocks)

        rec = {
            'code': code,
            'name': name,
            'price': price,
            'fwd_pe': fwd_pe,
            'fwd_pb': fwd_pb,
            'pe_rank': pe_rank,
            'pe_rank_3y': pe_rank_3y,
            'pe_rank_5y': pe_rank_5y,
            'pb_rank': pb_rank,
            'eps_yoy_4q': eps_yoy_4q,
            'eps_yoy_cur': eps_yoy_cur,
            'eps_yoy_nxt': eps_yoy_nxt,
            'rev_rank': rev_rank,
            'gpm_4q': gpm,
            'opm_4q': opm,
            'nim_4q': npm,
            'div_yield': safe_div(div, price) * 100 if safe_div(div, price) is not None else None,
            'analysts': analysts,
            'industry': industry.strip() if isinstance(industry, str) else None,
            'has_future': 1.0 if code in fut_codes else 0.0,
        }
        # 跟舊版一樣：沒有值的欄位不放進去（網頁顯示成「-」）
        rows.append({k: v for k, v in rec.items() if v is not None and v != ''})
    return rows


def export(xlsm_path, out_dir):
    tmp = tempfile.mkdtemp(prefix='screener_')
    try:
        with zipfile.ZipFile(xlsm_path) as z:
            z.extractall(tmp)
        print('篩選器資料計算中（直接由原始資料 sheet 計算，不讀統整的公式欄）...', flush=True)
        rows = compute(tmp, log=lambda s: print(s, flush=True))

        result = {
            'generated_at': datetime.datetime.now().isoformat(),
            'columns': [{'key': k, 'label': lbl, 'type': t} for k, lbl, t in COLUMNS],
            'rows': rows,
        }
        os.makedirs(out_dir, exist_ok=True)
        js_path = os.path.join(out_dir, 'screener.js')
        with open(js_path, 'w', encoding='utf-8') as f:
            f.write('window.SCREENER = ')
            json.dump(result, f, ensure_ascii=False, separators=(',', ':'))
            f.write(';\n')

        n_pe = sum(1 for r in rows if 'fwd_pe' in r)
        print(f'篩選器資料已匯出：{len(rows)} 檔股票（其中 {n_pe} 檔有 Fwd. PE）-> {js_path}')
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
