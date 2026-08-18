"""
Windows 上刪檔案／資料夾的防呆工具。

為什麼需要這個
--------------
在 Windows 上刪資料夾很容易踩到「PermissionError: [WinError 5] 存取被拒」，常見原因有：
  - 解壓縮出來的檔案帶唯讀屬性
  - 檔案總管正好開著那個資料夾（資料夾本身刪不掉，但裡面的檔案刪得掉）
  - 防毒軟體、雲端同步、搜尋索引正在掃描裡面的檔案
這些都跟「Excel 有沒有關掉」無關，所以不該讓整個匯出流程因此中斷。
"""
import os
import stat
import time
import shutil
import glob


def make_writable(path):
    """把唯讀屬性拿掉。資料夾一定要保留執行權限，否則之後反而進不去、更刪不掉。"""
    try:
        os.chmod(path, 0o700 if os.path.isdir(path) else 0o600)
    except Exception:
        pass


def force_rmtree(path, retries=3):
    """盡力刪掉整個資料夾樹，刪不掉也只回傳 False，不會丟出例外。"""
    def clear_readonly(root):
        make_writable(root)
        for dirpath, dirnames, filenames in os.walk(root, topdown=True):
            for name in dirnames:
                make_writable(os.path.join(dirpath, name))
            for name in filenames:
                make_writable(os.path.join(dirpath, name))

    def on_error(func, p, exc_info):
        try:
            make_writable(os.path.dirname(p))
            make_writable(p)
            func(p)
        except Exception:
            pass

    for i in range(retries):
        if not os.path.exists(path):
            return True
        try:
            clear_readonly(path)
            shutil.rmtree(path, onerror=on_error)
        except Exception:
            pass
        if not os.path.exists(path):
            return True
        time.sleep(0.4 * (i + 1))
    return not os.path.exists(path)


def clean_dir(dir_path, pattern='*', retries=3):
    """清空資料夾裡符合 pattern 的檔案，但「保留資料夾本身」。

    比 rmtree 整個砍掉安全得多：資料夾被檔案總管開著時 rmdir 會失敗，
    但刪裡面的檔案通常沒問題。回傳 (刪掉幾個, 刪不掉的清單)。
    """
    os.makedirs(dir_path, exist_ok=True)
    # 資料夾本身如果是唯讀的，裡面的檔案也刪不掉，先把它解開
    make_writable(dir_path)
    failed = []
    removed = 0
    for p in glob.glob(os.path.join(dir_path, pattern)):
        if os.path.isdir(p):
            if not force_rmtree(p):
                failed.append(p)
            continue
        ok = False
        for i in range(retries):
            try:
                os.remove(p)
                ok = True
                break
            except Exception:
                make_writable(p)
                time.sleep(0.2 * (i + 1))
        if ok:
            removed += 1
        else:
            failed.append(p)
    return removed, failed
