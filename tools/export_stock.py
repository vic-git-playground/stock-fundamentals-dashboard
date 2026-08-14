"""
Export one stock's 8-chart dataset from 台灣篩選器_股期_ver.2.xlsm (unzipped to xlsm_extract/)
into a JSON file the webpage can read, without needing Excel/LibreOffice to recalc.

Logic replicated from the "個股" sheet formulas (reverse engineered):
  - Weekly backbone dates come from 周收 (weekly close) sheet, row4 = yyyymmdd keys.
  - Weekly weighted Fwd EPS from 加權Fwd.EPS sheet, row5 = yyyymm keys (matched by week's month).
  - Weekly PE = weekly price / weekly weighted Fwd EPS.
  - Weekly PB from PB sheet, row4 = yyyymmdd keys, forward-filled if a week is missing.
  - Rolling 3yr (156-week) Min/Max/Median PE and PB bands, trailing window ending at each week.
  - Weekly EPS growth (this-year / next-year) from 預估今年EPS成長 / 預估明年EPS成長, row4 = yyyymm (month-level, repeated across weeks in the month).
  - Weekly institutional holdings % from 法人持股 / 投信持股 / 外資持股 / 自營持股, row4 = yyyymmdd keys, forward-filled.
  - Monthly backbone from 月營收 sheet's own yyyymm columns (row4).
  - Monthly revenue + 3-month trailing moving average.
  - Monthly GPM/OPM from Margin sheet (quarterly data in offset blocks), forward-filled onto monthly axis.
  - Monthly Bwd EPS (trailing 4Q) from Bwd.EPS sheet (quarterly), forward-filled onto monthly axis.
  - Monthly Fwd EPS (weighted) from 加權Fwd.EPS (already monthly-native).
  - Monthly EPS estimates (this-year / next-year) from 今年Est.EPS / 明年Est.EPS (monthly-native).
"""
import sys, os, json, re, statistics, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xhelp2 import find_rows, row_series, shared_strings

SHEETS = {
    '周收': 'sheet5.xml',
    'PB': 'sheet4.xml',
    '法人持股': 'sheet6.xml',
    '自營持股': 'sheet7.xml',
    '投信持股': 'sheet8.xml',
    '外資持股': 'sheet9.xml',
    '加權Fwd.EPS': 'sheet10.xml',
    '明年Est.EPS': 'sheet12.xml',
    '預估明年EPS成長': 'sheet13.xml',
    '今年Est.EPS': 'sheet14.xml',
    '預估今年EPS成長': 'sheet15.xml',
    'Bwd.EPS': 'sheet16.xml',
    'Margin': 'sheet18.xml',
    '月營收': 'sheet19.xml',
}


# 資料來源：None = 照舊直接讀 Excel；設成 ArchiveSource 物件 = 改讀歷史資料倉。
# refresh_data.py 會在資料倉存在時自動指定，這樣 Excel 只要留最新一期就好。
SOURCE = None


def set_source(src):
    global SOURCE
    SOURCE = src


def get_series(sheet_key, code, header_row):
    if SOURCE is not None:
        return SOURCE.get_series(sheet_key, code, header_row)
    headers, row, rn = find_rows(SHEETS[sheet_key], header_rows=(header_row,), code_col='A', code_value=code)
    if row is None:
        return None, None
    name = row.get('B' + str(rn))
    series = row_series(headers[header_row], row, start_col_num=3)
    # series: list of (key, value) as read (key may be str or number); normalize key to str
    out = [(str(k), v) for k, v in series if k is not None]
    return name, out


def to_ymd(key):
    s = str(key)
    return datetime.date(int(s[0:4]), int(s[4:6]), int(s[6:8]))


def to_ym(key):
    s = str(key)
    return int(s[0:6])


def ym_to_iso(ym):
    y = ym // 100
    m = ym % 100
    return f"{y:04d}-{m:02d}-01"


def rolling(values, window):
    """Trailing rolling min/max/median ending at each index (values assumed chronological asc)."""
    mins, maxs, meds = [], [], []
    for i in range(len(values)):
        lo = max(0, i - window + 1)
        chunk = [v for v in values[lo:i + 1] if v is not None]
        if not chunk:
            mins.append(None); maxs.append(None); meds.append(None)
            continue
        mins.append(min(chunk))
        maxs.append(max(chunk))
        meds.append(statistics.median(chunk))
    return mins, maxs, meds


def forward_fill_onto(monthly_axis_ym, src_pairs_ym):
    """src_pairs_ym: list of (ym:int, value) sorted ascending. For each ym in monthly_axis_ym,
    take the value at that ym if present, else the most recent prior ym's value."""
    src = {ym: v for ym, v in src_pairs_ym if v is not None}
    src_yms = sorted(src.keys())
    out = []
    j = -1
    last = None
    si = 0
    for ym in monthly_axis_ym:
        while si < len(src_yms) and src_yms[si] <= ym:
            last = src[src_yms[si]]
            si += 1
        out.append(last)
    return out


def margin_series(code):
    """Margin sheet: 3 blocks (GPM, OPM, pretax) of same quarter columns, OPM block offset +40 cols from GPM.
    Returns dict ym(quarter-end month int e.g. 202606) -> (gpm, opm)."""
    if SOURCE is not None:
        return SOURCE.margin_series(code)
    headers, row, rn = find_rows(SHEETS['Margin'], header_rows=(4,), code_col='A', code_value=code)
    if row is None:
        return {}
    from xhelp2 import col_num, col_letters
    hdr = headers[4]
    # sort header cells by column number
    items = sorted(((col_num(c), c, v) for c, v in hdr.items() if v is not None), key=lambda x: x[0])
    data_by_col = {col_letters(c): val for c, val in row.items()}
    out = {}
    n = len(items)
    for idx, (cn, coord, key) in enumerate(items):
        try:
            ymq = int(str(key))
        except Exception:
            continue
        gpm_col = col_letters(coord)
        gpm = data_by_col.get(gpm_col)
        # OPM lives 40 columns to the right (same relative row); find header entry ~40 later in this sorted list
        opm = None
        if idx + 40 < n:
            opm_coord = items[idx + 40][1]
            opm = data_by_col.get(col_letters(opm_coord))
        if gpm is not None:
            out.setdefault(ymq, [None, None])[0] = gpm
        if opm is not None:
            out.setdefault(ymq, [None, None])[1] = opm
    return out


def export(code, code_name_hint=None):
    code = str(code)
    result = {'code': code, 'generated_at': datetime.datetime.now().isoformat()}

    # ---- Weekly backbone: price ----
    name, price_series = get_series('周收', code, 4)
    if name is None:
        return None, f'股票代號 {code} 找不到資料 (周收 sheet)'
    result['name'] = name
    price_series.sort(key=lambda kv: to_ymd(kv[0]))  # ascending oldest->newest
    weeks = [k for k, v in price_series]
    prices = [v for k, v in price_series]

    # weighted fwd EPS, monthly-native -> map onto weekly by week's yyyymm
    _, fwdeps_m = get_series('加權Fwd.EPS', code, 5)
    fwdeps_m = fwdeps_m or []
    fwdeps_by_ym = {int(k): v for k, v in fwdeps_m if v is not None}

    fwd_pe_weekly = []
    for wk, px in zip(weeks, prices):
        ym = to_ym(wk)
        eps = fwdeps_by_ym.get(ym)
        if px is not None and eps not in (None, 0):
            fwd_pe_weekly.append(px / eps)
        else:
            fwd_pe_weekly.append(None)

    # PB weekly, forward-filled
    _, pb_series = get_series('PB', code, 4)
    pb_by_date = {k: v for k, v in (pb_series or [])}
    pb_weekly = []
    last_pb = None
    for wk in weeks:
        v = pb_by_date.get(wk)
        if v not in (None, 0):
            last_pb = v
        pb_weekly.append(last_pb)

    # EPS growth weekly (month-level lookups)
    _, growth_ty_m = get_series('預估今年EPS成長', code, 4)
    _, growth_ny_m = get_series('預估明年EPS成長', code, 4)
    gty_by_ym = {int(k): v for k, v in (growth_ty_m or []) if v is not None}
    gny_by_ym = {int(k): v for k, v in (growth_ny_m or []) if v is not None}
    growth_ty_weekly = [gty_by_ym.get(to_ym(wk)) for wk in weeks]
    growth_ny_weekly = [gny_by_ym.get(to_ym(wk)) for wk in weeks]

    # Holdings weekly, forward-filled
    def holdings_ff(sheet_key):
        _, s = get_series(sheet_key, code, 4)
        by_date = {k: v for k, v in (s or [])}
        out, last = [], None
        for wk in weeks:
            v = by_date.get(wk)
            if v is not None:
                last = v
            out.append(last)
        return out

    hold_total = holdings_ff('法人持股')
    hold_trust = holdings_ff('投信持股')
    hold_foreign = holdings_ff('外資持股')
    hold_dealer = holdings_ff('自營持股')

    # Rolling 3yr (156wk) PE / PB bands
    pe_min, pe_max, pe_med = rolling(fwd_pe_weekly, 156)
    pb_min, pb_max, pb_med = rolling(pb_weekly, 156)

    def mul(a, b):
        return a * b if (a is not None and b is not None) else None

    price_band_min = [mul(pe_min[i], fwdeps_by_ym.get(to_ym(weeks[i]))) for i in range(len(weeks))]
    price_band_max = [mul(pe_max[i], fwdeps_by_ym.get(to_ym(weeks[i]))) for i in range(len(weeks))]
    price_band_med = [mul(pe_med[i], fwdeps_by_ym.get(to_ym(weeks[i]))) for i in range(len(weeks))]

    iso_weeks = [str(to_ymd(w).isoformat()) for w in weeks]

    result['weekly'] = {
        'dates': iso_weeks,
        'price': prices,
        'fwd_pe': fwd_pe_weekly,
        'pb': pb_weekly,
        'pe_band_min': price_band_min,
        'pe_band_max': price_band_max,
        'pe_band_med': price_band_med,
        'pe_roll_min_x': pe_min,
        'pe_roll_max_x': pe_max,
        'pe_roll_med_x': pe_med,
        'pb_roll_min_x': pb_min,
        'pb_roll_max_x': pb_max,
        'pb_roll_med_x': pb_med,
        'eps_growth_this_year': growth_ty_weekly,
        'eps_growth_next_year': growth_ny_weekly,
        'holdings_total': hold_total,
        'holdings_trust': hold_trust,
        'holdings_foreign': hold_foreign,
        'holdings_dealer': hold_dealer,
    }

    # ---- Monthly backbone: revenue ----
    _, rev_series = get_series('月營收', code, 4)
    rev_series = [(int(k), v) for k, v in (rev_series or []) if v is not None]
    rev_series.sort(key=lambda kv: kv[0])
    months = [ym for ym, v in rev_series]
    rev = [v for ym, v in rev_series]
    rev_3ma = []
    for i in range(len(rev)):
        lo = max(0, i - 2)
        chunk = rev[lo:i + 1]
        rev_3ma.append(sum(chunk) / len(chunk) if chunk else None)

    _, eps_ty_m = get_series('今年Est.EPS', code, 4)
    _, eps_ny_m = get_series('明年Est.EPS', code, 4)
    eps_ty_by_ym = {int(k): v for k, v in (eps_ty_m or []) if v is not None}
    eps_ny_by_ym = {int(k): v for k, v in (eps_ny_m or []) if v is not None}
    eps_ty_monthly = [eps_ty_by_ym.get(m) for m in months]
    eps_ny_monthly = [eps_ny_by_ym.get(m) for m in months]

    fwdeps_monthly = [fwdeps_by_ym.get(m) for m in months]

    _, bwdeps_q = get_series('Bwd.EPS', code, 4)
    bwdeps_pairs = [(int(k), v) for k, v in (bwdeps_q or []) if v is not None]
    bwdeps_pairs.sort(key=lambda kv: kv[0])
    bwdeps_monthly = forward_fill_onto(months, bwdeps_pairs)

    margin_q = margin_series(code)  # ym(quarter) -> [gpm, opm]
    gpm_pairs = sorted([(ym, vv[0]) for ym, vv in margin_q.items() if vv[0] is not None])
    opm_pairs = sorted([(ym, vv[1]) for ym, vv in margin_q.items() if vv[1] is not None])
    gpm_monthly = forward_fill_onto(months, gpm_pairs)
    opm_monthly = forward_fill_onto(months, opm_pairs)

    result['monthly'] = {
        'dates': [ym_to_iso(m) for m in months],
        'revenue': rev,
        'revenue_3ma': rev_3ma,
        'eps_this_year': eps_ty_monthly,
        'eps_next_year': eps_ny_monthly,
        'fwd_eps': fwdeps_monthly,
        'bwd_eps': bwdeps_monthly,
        'gpm': gpm_monthly,
        'opm': opm_monthly,
    }

    # Latest snapshot values for reference lines / headline numbers
    def last_non_none(lst):
        for v in reversed(lst):
            if v is not None:
                return v
        return None

    latest_price = last_non_none(prices)
    latest_pb = last_non_none(pb_weekly)
    eps_used_per_week = [fwdeps_by_ym.get(to_ym(wk)) for wk in weeks]
    latest_weekly_fwd_eps = last_non_none(eps_used_per_week)
    latest_bvps = (latest_price / latest_pb) if (latest_price is not None and latest_pb not in (None, 0)) else None

    result['latest'] = {
        'price': latest_price,
        'fwd_pe': last_non_none(fwd_pe_weekly),
        'pb': latest_pb,
        'pe_band_min_x': last_non_none(pe_min),
        'pe_band_max_x': last_non_none(pe_max),
        'pe_band_med_x': last_non_none(pe_med),
        'pb_band_min_x': last_non_none(pb_min),
        'pb_band_max_x': last_non_none(pb_max),
        'pb_band_med_x': last_non_none(pb_med),
        'gpm': last_non_none(gpm_monthly),
        'opm': last_non_none(opm_monthly),
        # 給網頁「手動調整股價」與「評價試算器」用：最新一週 PE 分母用的加權 Fwd EPS，以及推回來的每股淨值
        'weekly_fwd_eps': latest_weekly_fwd_eps,
        'bvps': latest_bvps,
    }

    return result, None


if __name__ == '__main__':
    code = sys.argv[1] if len(sys.argv) > 1 else '3037'
    data, err = export(code)
    if err:
        print('ERROR:', err)
        sys.exit(1)
    print(json.dumps({
        'code': data['code'], 'name': data['name'],
        'n_weekly': len(data['weekly']['dates']),
        'n_monthly': len(data['monthly']['dates']),
        'latest': data['latest'],
        'weekly_tail': {k: v[-3:] for k, v in data['weekly'].items() if isinstance(v, list)},
        'monthly_tail': {k: v[-3:] for k, v in data['monthly'].items() if isinstance(v, list)},
    }, ensure_ascii=False, indent=1, default=str))
