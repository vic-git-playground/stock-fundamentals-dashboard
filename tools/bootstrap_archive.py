"""
一次性作業：把「現在這份完整的 Excel」裡所有 CMoney 原始資料的完整歷史，
灌進歷史資料倉 archive/。

⚠️ 這一步一定要在你把 CMoney 的 TOP N 改小之前做完，不然歷史就抓不回來了。
   做完之後請把 archive/ 整個資料夾另外備份一份（例如複製到雲端硬碟），
   它就是你所有歷史資料的唯一底本了。

用法：
    python tools/bootstrap_archive.py "D:\\工作區\\主要工作檔\\台灣篩選器_股期_ver.2.xlsm"
不指定 archive 資料夾的話，預設放在專案資料夾底下的 archive/。
"""
import os
import sys
import time
import argparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import cmoney_sheets as cs
import archive_store


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('xlsm_path')
    ap.add_argument('--archive-dir', default=None)
    ap.add_argument('--series', default=None,
                    help='逗號分隔，只灌指定的幾個 series；不給就全部')
    args = ap.parse_args()

    project_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    archive_dir = args.archive_dir or os.path.join(project_dir, 'archive')

    if not os.path.exists(args.xlsm_path):
        print(f'找不到 Excel 檔案：{args.xlsm_path}')
        sys.exit(1)

    series_list = args.series.split(',') if args.series else list(cs.SERIES.keys())

    print(f'來源 Excel : {args.xlsm_path}')
    print(f'資料倉位置 : {archive_dir}')
    print(f'要灌的表   : {len(series_list)} 張')
    print('-' * 66)

    t0 = time.time()
    print('解壓縮並讀取 Excel...')
    res = cs.read_all(args.xlsm_path, series_list)
    print(f'讀取完成，耗時 {time.time()-t0:.1f} 秒')
    print('-' * 66)

    total_cells = 0
    for s in series_list:
        names, data, periods = res[s]
        if not periods:
            print(f'{s:<18} 沒有期別欄，略過')
            continue
        added, overwritten, total = archive_store.merge(archive_dir, s, names, data, periods)
        cells = sum(len(v) for v in data.values())
        total_cells += cells
        print(f'{s:<18} 股票 {len(data):>4} 檔，新增 {added:>4} 期，覆蓋 {overwritten:>4} 期，'
              f'資料倉現有 {total:>4} 期（{periods[-1]} ~ {periods[0]}）')

    print('-' * 66)
    size = sum(os.path.getsize(os.path.join(archive_dir, f))
               for f in os.listdir(archive_dir) if f.endswith('.csv'))
    print(f'完成，共寫入約 {total_cells:,} 格資料，archive/ 佔用 {size/1e6:.1f} MB，'
          f'總耗時 {time.time()-t0:.1f} 秒')
    print()
    print('⚠️ 請務必把 archive/ 這個資料夾另外備份一份再繼續下一步。')


if __name__ == '__main__':
    main()
