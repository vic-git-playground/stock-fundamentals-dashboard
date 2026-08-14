"""
歷史資料倉：把每一張 CMoney 原始資料表的完整歷史保存在 archive/<series>.csv。

為什麼要有這個東西
------------------
原本的做法是每次都叫 CMoney 把「全部期數」重抓一遍（例如 周收 TOP 600、月營收 TOP 120），
資料量大、很花時間。改成資料倉之後：
  1. CMoney 只抓最新 1~2 期（在 CMoney 設定裡把 TOP N 改小），更新很快
  2. Python 把抓回來的那幾期「覆蓋」進資料倉（以期別為鍵，重複匯入不會累積重複資料）
  3. 網站要用的完整歷史，一律從資料倉讀，而不是從 Excel 讀

檔案格式
--------
archive/<series>.csv，UTF-8-SIG（Excel 打得開），排法跟 Excel 原本的 sheet 一樣：
    股票代號,股票名稱,20260811,20260807,20260731,...
    1101,台泥,24.65,24.35,24.3,...
期別欄由新到舊排序（跟 CMoney 抓下來的順序一致）。

期別（period）就是 Excel 那些 sheet 第 4 列的值：
  日/週資料 = 20260811，月資料 = 202607，季資料 = 202604（CMoney 用季末月表示）
"""
import os
import csv


def archive_path(archive_dir, series):
    return os.path.join(archive_dir, f'{series}.csv')


def _period_sort_key(p):
    """期別由新到舊排序。期別都是數字字串（20260811 / 202607），長度不同時分開比。"""
    s = str(p)
    return (len(s), s)


def load(archive_dir, series):
    """讀回一個 series 的完整歷史。

    回傳 (names, data, periods)：
      names   : {code: 股票名稱}
      data    : {code: {period: value(float or str)}}
      periods : [period, ...] 由新到舊
    資料倉還沒有這個 series 時回傳空的結構，不會報錯。
    """
    path = archive_path(archive_dir, series)
    names, data, periods = {}, {}, []
    if not os.path.exists(path):
        return names, data, periods

    with open(path, 'r', encoding='utf-8-sig', newline='') as f:
        reader = csv.reader(f)
        header = next(reader, None)
        if not header or len(header) < 2:
            return names, data, periods
        periods = [h for h in header[2:]]
        for row in reader:
            if not row or not row[0].strip():
                continue
            code = row[0].strip()
            names[code] = row[1].strip() if len(row) > 1 else ''
            vals = {}
            for i, p in enumerate(periods, start=2):
                if i >= len(row):
                    break
                raw = row[i].strip()
                if raw == '':
                    continue
                try:
                    vals[p] = float(raw)
                except ValueError:
                    vals[p] = raw
            data[code] = vals
    return names, data, periods


def save(archive_dir, series, names, data, periods):
    """把完整歷史寫回 archive/<series>.csv。"""
    os.makedirs(archive_dir, exist_ok=True)
    path = archive_path(archive_dir, series)
    periods = sorted(set(periods), key=_period_sort_key, reverse=True)

    # 先寫暫存檔再換掉，避免寫到一半當掉就把原本的歷史弄壞
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['股票代號', '股票名稱'] + periods)
        for code in sorted(data.keys()):
            vals = data[code]
            row = [code, names.get(code, '')]
            for p in periods:
                v = vals.get(p)
                row.append('' if v is None else v)
            w.writerow(row)
    os.replace(tmp, path)
    return path


def merge(archive_dir, series, new_names, new_data, new_periods):
    """把新抓到的幾期覆蓋進資料倉，回傳 (新增期數, 覆蓋期數, 目前總期數)。

    規則：
      - 同一個期別已經存在 -> 用新的值整欄覆蓋（CMoney 有時會事後修正數字，以新的為準）
      - 新的期別 -> 直接加進去
      - 資料倉裡有、這次沒抓到的期別 -> 原封不動保留
      - 這次沒抓到的股票 -> 舊資料保留，不會被清掉
    """
    names, data, periods = load(archive_dir, series)

    existing = set(periods)
    added = [p for p in new_periods if p not in existing]
    overwritten = [p for p in new_periods if p in existing]

    names.update({c: n for c, n in new_names.items() if n})

    for code, vals in new_data.items():
        cur = data.setdefault(code, {})
        for p in new_periods:
            if p in vals:
                cur[p] = vals[p]
            else:
                # 這一期 CMoney 給空值 -> 該格清成空白（代表這檔這期沒有資料）
                cur.pop(p, None)

    all_periods = sorted(existing | set(new_periods), key=_period_sort_key, reverse=True)
    save(archive_dir, series, names, data, all_periods)
    return len(added), len(overwritten), len(all_periods)


def summary(archive_dir, series_list):
    """列出資料倉現況，給更新流程印訊息用。"""
    out = []
    for s in series_list:
        names, data, periods = load(archive_dir, s)
        out.append({
            'series': s,
            'stocks': len(data),
            'periods': len(periods),
            'newest': periods[0] if periods else None,
            'oldest': periods[-1] if periods else None,
        })
    return out
