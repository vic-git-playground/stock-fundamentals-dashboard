"""
用 Python 重算 Excel 裡那三張「純公式表」，這樣那三張表就可以從 Excel 裡刪掉
（它們佔了 227MB 裡的 110MB，而且每次開檔都要重算，是拖慢速度的主因之一）。

對應關係（公式是從 Excel 裡挖出來逐字對照的）：

1. 加權Fwd.EPS
   Excel: IF(ISBLANK(明年Est.EPS), 今年Est.EPS,
              明年Est.EPS*((月-1)/12) + 今年Est.EPS*(1-(月-1)/12))
   「月」= 該期別的月份（202608 -> 8）。意思是年初幾乎全看今年預估，
   越接近年底越往明年預估靠過去，12 月時就完全用明年的。

2. FwdPE_周
   Excel: 周收[某週] / 加權Fwd.EPS[該週所屬的年月]

3. 預估明年EPS成長
   Excel: IF(AND(ISNUMBER(明年Est.EPS),ISNUMBER(今年Est.EPS)), (明年/今年-1)*100, "")

（預估今年EPS成長 是 CMoney 直接給的原始資料，不用算，直接進資料倉。）
"""


def weighted_fwd_eps(ty_by_ym, ny_by_ym):
    """加權 Fwd. EPS：{ym: value}

    ty_by_ym / ny_by_ym 是 {期別(字串或int): 值}，期別格式 yyyymm。
    """
    out = {}
    for ym in set(ty_by_ym) | set(ny_by_ym):
        try:
            month = int(str(ym)[4:6])
        except (ValueError, IndexError):
            continue
        ty = ty_by_ym.get(ym)
        ny = ny_by_ym.get(ym)
        if ty is None and ny is None:
            continue
        if ny is None:
            out[str(ym)] = ty
            continue
        if ty is None:
            # Excel 的 IF 只判斷「明年」是否空白；明年有值而今年空白時，
            # 公式裡的 VLOOKUP 會把空白格當成 0，這裡照樣處理以維持數字一致
            ty = 0.0
        w = (month - 1) / 12.0
        out[str(ym)] = ny * w + ty * (1 - w)
    return out


def fwd_pe_weekly(price_by_date, weighted_by_ym, all_dates=None, excel_compat=False):
    """FwdPE_周：{yyyymmdd: value}，用該週所屬年月的加權 Fwd. EPS 當分母。

    excel_compat=True 時，會完整重現 Excel 那張表的行為：某一週股價是空白（停牌、
    尚未上市、或當週沒有收盤資料）時，Excel 的公式 `周收!C6/...` 會把空白格當成 0，
    所以那一格算出來是 0 而不是空白。這些 0 會被 PERCENTRANK 算進去、影響歷史位階，
    所以計算「PE 歷史位階」時要用這個模式才會跟 Excel 對得起來。
    畫圖用的序列則不要開，否則圖上會出現假的 0 值。
    """
    out = {}
    dates = all_dates if all_dates is not None else price_by_date.keys()
    for d in dates:
        px = price_by_date.get(d)
        ym = str(d)[:6]
        eps = weighted_by_ym.get(ym)
        if eps in (None, 0):
            continue          # 分母查不到 -> Excel 是 IFERROR 出來的空字串
        if px is None:
            if excel_compat:
                out[str(d)] = 0.0
            continue
        out[str(d)] = px / eps
    return out


def eps_growth_next_year(ty_by_ym, ny_by_ym):
    """預估明年EPS成長(%)：{ym: value}"""
    out = {}
    for ym in set(ty_by_ym) & set(ny_by_ym):
        ty = ty_by_ym.get(ym)
        ny = ny_by_ym.get(ym)
        if ty in (None, 0) or ny is None:
            continue
        out[str(ym)] = (ny / ty - 1) * 100
    return out
