"""
讓 export_stock.py 改從「歷史資料倉」拿資料，而不是從 Excel 拿。

介面刻意做成跟 export_stock.py 原本的 get_series() / margin_series() 一樣，
所以那支程式的計算邏輯完全不用改，只是換了資料來源。

三張原本是 Excel 公式表的資料（加權Fwd.EPS / FwdPE_周 / 預估明年EPS成長）
在這裡由 derived.py 即時算出來，所以 Excel 裡可以把那三張表刪掉。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import archive_store
import derived

# export_stock.py 會用到的 series
RAW_SERIES = [
    '周收', 'PB', '法人持股', '自營持股', '投信持股', '外資持股',
    '今年Est.EPS', '明年Est.EPS', '預估今年EPS成長',
    'Bwd.EPS', '月營收', 'Margin_GPM', 'Margin_OPM', 'Margin_NIM',
]


class ArchiveSource:
    """把整個資料倉載進記憶體一次，之後每檔股票查詢都是 dict lookup。

    資料倉全部才 4MB 左右，一次載進來最單純也最快。
    """

    def __init__(self, archive_dir):
        self.archive_dir = archive_dir
        self.raw = {}
        for s in RAW_SERIES:
            names, data, periods = archive_store.load(archive_dir, s)
            self.raw[s] = {'names': names, 'data': data, 'periods': periods}

        missing = [s for s in RAW_SERIES if not self.raw[s]['periods']]
        if missing:
            raise RuntimeError(
                '資料倉缺少這些資料表：' + '、'.join(missing) +
                '\n請先執行 tools/bootstrap_archive.py 從完整的 Excel 灌一次歷史。')

        # 衍生資料先全部算好放著
        ty = self.raw['今年Est.EPS']['data']
        ny = self.raw['明年Est.EPS']['data']
        px = self.raw['周收']['data']
        self.weighted = {}
        self.fwdpe = {}          # 畫圖用：股價空白的那週就是沒有值
        self.fwdpe_excel = {}    # 算 PE 歷史位階用：完整重現 Excel（空白股價 -> 0）
        self.growth_ny = {}
        week_dates = self.raw['周收']['periods']
        for code in set(ty) | set(ny):
            w = derived.weighted_fwd_eps(ty.get(code, {}), ny.get(code, {}))
            self.weighted[code] = w
            self.growth_ny[code] = derived.eps_growth_next_year(ty.get(code, {}), ny.get(code, {}))
            if code in px:
                self.fwdpe[code] = derived.fwd_pe_weekly(px[code], w)
                self.fwdpe_excel[code] = derived.fwd_pe_weekly(
                    px[code], w, all_dates=week_dates, excel_compat=True)

    # ---- export_stock.py 用的介面 ----

    def stock_name(self, code):
        for s in ('周收', '月營收', '今年Est.EPS'):
            n = self.raw[s]['names'].get(code)
            if n:
                return n
        return None

    def all_codes(self):
        """資料倉裡有周收資料的股票代號（＝可以畫圖的）"""
        return sorted(self.raw['周收']['data'].keys())

    def get_series(self, sheet_key, code, header_row=None):
        """回傳 (name, [(期別字串, 值), ...])，跟原本讀 Excel 的版本一樣。"""
        code = str(code)

        if sheet_key == '加權Fwd.EPS':
            vals = self.weighted.get(code)
            return (self.stock_name(code), sorted(vals.items())) if vals else (self.stock_name(code), [])

        if sheet_key == '預估明年EPS成長':
            vals = self.growth_ny.get(code)
            return (self.stock_name(code), sorted(vals.items())) if vals else (self.stock_name(code), [])

        if sheet_key == 'FwdPE_周':
            vals = self.fwdpe.get(code)
            return (self.stock_name(code), sorted(vals.items())) if vals else (self.stock_name(code), [])

        if sheet_key not in self.raw:
            raise KeyError(f'資料倉沒有這個資料表：{sheet_key}')

        blk = self.raw[sheet_key]
        if code not in blk['data']:
            return None, None
        vals = blk['data'][code]
        # 一定要回傳「完整的期別軸」，沒有值的期別給 None。
        # 因為 PE/PB 的三年滾動區間是用「筆數」數的，如果把空白期直接跳過，
        # 視窗會偷偷位移，算出來的高低中位數就會跟 Excel 對不起來。
        return blk['names'].get(code), [(p, vals.get(p)) for p in sorted(blk['periods'])]

    def margin_series(self, code):
        """{季別(int, 例 202606): [gpm, opm]}，對齊原本 export_stock.margin_series() 的輸出。"""
        code = str(code)
        gpm = self.raw['Margin_GPM']['data'].get(code, {})
        opm = self.raw['Margin_OPM']['data'].get(code, {})
        out = {}
        for per, v in gpm.items():
            try:
                out.setdefault(int(per), [None, None])[0] = v
            except ValueError:
                continue
        for per, v in opm.items():
            try:
                out.setdefault(int(per), [None, None])[1] = v
            except ValueError:
                continue
        return out
