"""
用 Python 重算「統整」sheet 裡那些需要長歷史才能算的欄位。

為什麼需要這支程式
------------------
把 CMoney 的 TOP N 改小之後，Excel 裡的這幾欄就會失真，因為它們是對整列歷史做統計：
    P 欄 PE 歷史位階   = PERCENTRANK(FwdPE_周 整列, 目前Fwd.PE)
    Q 欄 PB 歷史位階   = PERCENTRANK(PB 整列, 目前Fwd.PB)
    U 欄 10年月營收排行 = RANK(最新月營收, 月營收 整列)
    V/W/X 4Q GPM/OPM/NIM = 最近 4 季平均
    R 欄 累積近4Q EPS YoY = Bwd.EPS 最新 ÷ 四季前 - 1
改由這裡從歷史資料倉計算，Excel 就只要負責抓最新一期。
"""
import statistics


def percentrank(values, x, significance=3):
    """重現 Excel 的 PERCENTRANK()。

    Excel: (小於 x 的個數 + 內插比例) / (n-1)，結果截斷到指定位數。
    統整 sheet 外面還包了一層：x >= 最大值 -> 100%，x <= 最小值 -> 0%。
    """
    arr = sorted(v for v in values if v is not None)
    n = len(arr)
    if n == 0 or x is None:
        return None
    if n == 1:
        return 1.0 if x >= arr[0] else 0.0
    if x >= arr[-1]:
        return 1.0
    if x <= arr[0]:
        return 0.0

    # 找出 x 落在哪兩個值之間
    lo = 0
    for i in range(n - 1):
        if arr[i] <= x <= arr[i + 1]:
            lo = i
            break
    span = arr[lo + 1] - arr[lo]
    frac = 0.0 if span == 0 else (x - arr[lo]) / span
    pr = (lo + frac) / (n - 1)

    # Excel 會截斷（不是四捨五入）到 significance 位小數
    factor = 10 ** significance
    return int(pr * factor) / factor


def rank_desc(values, x):
    """重現 Excel 的 RANK()（由大到小，最大的是第 1 名）。"""
    if x is None:
        return None
    vals = [v for v in values if v is not None]
    return sum(1 for v in vals if v > x) + 1


def avg_latest(values_by_period, periods_desc, count):
    """最近 count 期的平均，跳過沒有資料的期別（跟 Excel 的 AVERAGE 一樣忽略空白）。"""
    picked = []
    for p in periods_desc[:count]:
        v = values_by_period.get(p)
        if v is not None:
            picked.append(v)
    return sum(picked) / len(picked) if picked else None


def sorted_periods_desc(period_list):
    return sorted(period_list, key=lambda p: (len(str(p)), str(p)), reverse=True)


class ScreenerCalc:
    """從資料倉算出篩選器需要的歷史統計欄位。"""

    def __init__(self, source):
        """source 是 archive_source.ArchiveSource"""
        self.src = source
        raw = source.raw
        self.pb = raw['PB']
        self.rev = raw['月營收']
        self.bwd = raw['Bwd.EPS']
        self.gpm = raw['Margin_GPM']
        self.opm = raw['Margin_OPM']
        self.nim = raw.get('Margin_NIM')
        self.ty = raw['今年Est.EPS']
        self.ny = raw['明年Est.EPS']
        self.growth_ty = raw['預估今年EPS成長']

        self.pb_periods = sorted_periods_desc(self.pb['periods'])
        self.rev_periods = sorted_periods_desc(self.rev['periods'])
        self.bwd_periods = sorted_periods_desc(self.bwd['periods'])
        self.gpm_periods = sorted_periods_desc(self.gpm['periods'])
        self.opm_periods = sorted_periods_desc(self.opm['periods'])
        self.nim_periods = sorted_periods_desc(self.nim['periods']) if self.nim else []
        self.ty_periods = sorted_periods_desc(self.ty['periods'])

    def compute(self, code, fwd_pe=None, fwd_pb=None):
        """回傳這一檔的歷史統計欄位。

        fwd_pe / fwd_pb 是「目前」的值（由 統整 sheet 的最新一期算出來的）；
        沒有給的話就用資料倉裡最新一期自己算。
        """
        code = str(code)
        out = {}

        # ---- PE 歷史位階：用完整的 FwdPE_周 歷史（Excel 相容版，含空白股價的 0）----
        fwdpe_hist = self.src.fwdpe_excel.get(code, {})
        if fwd_pe is None:
            live = self.src.fwdpe.get(code, {})
            ps = sorted_periods_desc(live.keys())
            fwd_pe = live.get(ps[0]) if ps else None
        pr = percentrank(fwdpe_hist.values(), fwd_pe)
        out['pe_rank'] = None if pr is None else pr * 100

        # ---- PB 歷史位階：用完整的 PB 歷史 ----
        pb_hist = self.pb['data'].get(code, {})
        if fwd_pb is None:
            fwd_pb = pb_hist.get(self.pb_periods[0]) if self.pb_periods else None
        pr = percentrank(pb_hist.values(), fwd_pb)
        out['pb_rank'] = None if pr is None else pr * 100

        # ---- 10 年月營收排行：最新月營收在整段歷史裡排第幾 ----
        rev_hist = self.rev['data'].get(code, {})
        latest_rev = None
        for p in self.rev_periods:
            if rev_hist.get(p) is not None:
                latest_rev = rev_hist[p]
                break
        out['rev_rank'] = rank_desc(rev_hist.values(), latest_rev)

        # ---- 累積近 4Q EPS 的 YoY：最新一期 ÷ 四季前 - 1 ----
        bwd = self.bwd['data'].get(code, {})
        cur = bwd.get(self.bwd_periods[0]) if len(self.bwd_periods) > 0 else None
        yr_ago = bwd.get(self.bwd_periods[4]) if len(self.bwd_periods) > 4 else None
        out['eps_yoy_4q'] = (100 * (cur / yr_ago - 1)) if (cur is not None and yr_ago not in (None, 0)) else None

        # ---- 最近 4 季的 GPM / OPM / NIM 平均 ----
        out['gpm_4q'] = avg_latest(self.gpm['data'].get(code, {}), self.gpm_periods, 4)
        out['opm_4q'] = avg_latest(self.opm['data'].get(code, {}), self.opm_periods, 4)
        if self.nim:
            out['nim_4q'] = avg_latest(self.nim['data'].get(code, {}), self.nim_periods, 4)

        # ---- 今年 / 明年 EPS 預估成長 ----
        ty = self.ty['data'].get(code, {})
        ny = self.ny['data'].get(code, {})
        gty = self.growth_ty['data'].get(code, {})
        latest_ym = self.ty_periods[0] if self.ty_periods else None
        # 統整 S 欄是 VLOOKUP 到 預估今年EPS成長 的最新一欄；那格空白時 VLOOKUP 會回 0，
        # 這裡照樣處理（只要這檔股票有在那張表裡）
        gty_periods = sorted_periods_desc(self.growth_ty['periods'])
        gty_latest = gty_periods[0] if gty_periods else None
        if code in self.growth_ty['data']:
            v = gty.get(gty_latest)
            out['eps_yoy_cur'] = 0.0 if v is None else v
        else:
            out['eps_yoy_cur'] = None
        cur_ty, cur_ny = ty.get(latest_ym), ny.get(latest_ym)
        # 跟統整 T 欄一樣：今年預估是負的就不算成長率（除出來沒有意義）
        if cur_ty is not None and cur_ny is not None and cur_ty > 0:
            out['eps_yoy_nxt'] = 100 * (cur_ny / cur_ty - 1)
        else:
            out['eps_yoy_nxt'] = None

        return out
