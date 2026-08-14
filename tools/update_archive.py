"""
把 Excel 裡「這次 CMoney 抓回來的最新幾期」併進歷史資料倉。

這是每次更新都會跑的步驟（refresh_data.py 會自動呼叫），跟只跑一次的
bootstrap_archive.py 不一樣：
  - bootstrap_archive.py：一次性，從還有完整歷史的 Excel 把所有期數灌進資料倉
  - update_archive.py   ：每次更新，把 Excel 現在有的那幾期覆蓋進資料倉

因為是以「期別」為鍵覆蓋，所以：
  - Excel 只留最新 1~2 期也沒關係，舊歷史都在資料倉裡
  - 同一期重複跑很多次也不會重複累積
  - CMoney 事後修正數字時，重跑就會用新的值蓋掉舊的
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import cmoney_sheets as cs
import archive_store


def merge_from_excel(xlsm_path, archive_dir, extract_dir=None, series_list=None, verbose=True):
    """回傳 [(series, 新增期數, 覆蓋期數, 資料倉總期數), ...]"""
    series_list = series_list or list(cs.SERIES.keys())

    # 只處理 Excel 裡真的還存在的 sheet（有些表可能已經被刪掉或改名）
    tmp = extract_dir
    created = tmp is None
    if created:
        tmp = cs.unzip_xlsm(xlsm_path)
    elif not os.path.exists(os.path.join(tmp, 'xl', 'workbook.xml')):
        cs.unzip_xlsm(xlsm_path, tmp)

    try:
        ss = cs.load_shared_strings(tmp)
        sheets = cs.sheet_file_map(tmp)
        results = []
        skipped = []
        for s in series_list:
            sheet_name = cs.SERIES[s][0]
            if sheet_name not in sheets:
                skipped.append(s)
                continue
            try:
                names, data, periods = cs.read_series(tmp, ss, sheets, s)
            except Exception as e:
                skipped.append(f'{s}({e})')
                continue
            if not periods:
                skipped.append(s)
                continue
            added, overwritten, total = archive_store.merge(archive_dir, s, names, data, periods)
            results.append((s, added, overwritten, total))
            if verbose and (added or overwritten):
                newest = max(periods, key=lambda p: (len(str(p)), str(p)))
                print(f'  資料倉 {s:<16} 新增 {added} 期、更新 {overwritten} 期'
                      f'（最新 {newest}），累計 {total} 期')
        if verbose and skipped:
            print(f'  （Excel 裡沒有這些表，沿用資料倉既有歷史：{"、".join(map(str, skipped))}）')
        return results
    finally:
        if created:
            import shutil
            shutil.rmtree(tmp, ignore_errors=True)


def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument('xlsm_path')
    ap.add_argument('--archive-dir', default=None)
    args = ap.parse_args()
    project_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    archive_dir = args.archive_dir or os.path.join(project_dir, 'archive')
    if not os.path.isdir(archive_dir):
        print(f'找不到資料倉 {archive_dir}，請先執行一次 tools/bootstrap_archive.py')
        sys.exit(1)
    print(f'把 {args.xlsm_path} 的最新期數併進 {archive_dir}')
    merge_from_excel(args.xlsm_path, archive_dir)
    print('完成。')


if __name__ == '__main__':
    main()
