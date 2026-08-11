"""
在「有裝 CMoney」的那台電腦上執行：
  1. 呼叫 CMExcel.exe 更新 台灣篩選器_股期_ver.2.xlsm 裡的 CMoney 報表（做法照抄你原本 process_etf.py 的
     update_cmoney_excel()：subprocess 呼叫 CMExcel.exe，參數是 "M4||活頁簿路徑"）
  2. 確保 Excel 存檔關閉，釋放檔案鎖定
  3. 呼叫 refresh_data.py，把更新後的資料重新匯出成網站要用的 data/all_data.js

⚠️ 沒辦法在這裡（Claude 的工作環境）實際測試 CMExcel.exe，因為這台機器沒有裝 CMoney。
   "M4||路徑" 這個呼叫方式是照你給的 process_etf.py 抄的，那支程式是用在另一個檔案
   (主動式ETF_all.xlsx)，「M4」這個代碼在 CMoney 裡有沒有通用、還是綁特定報表格式，我沒辦法確認，
   要請你在那台機器上實際跑一次看看。如果跑起來 xlsm 裡的數字沒變，最直接的排除法是：
   打開 CMoney Excel 增益集，手動點一次「更新」，同時開工作管理員或用 Process Monitor 看它
   實際怎麼呼叫 CMExcel.exe（通常都在標題列或說明文件找得到用法），把正確參數改到下面
   CMONEY_UPDATE_ARG 這一行即可。
"""
import os
import sys
import time
import subprocess
import datetime

XLSM_FILENAME = "台灣篩選器_股期_ver.2.xlsm"
CMONEY_UPDATE_ARG_TEMPLATE = "M4||{file_path}"  # 跟 process_etf.py 用的格式一樣，不確定的話請調整


def log(msg):
    print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")


def find_cmexcel():
    candidates = [
        r"C:\Program Files\CMoney\CMExcel.exe",
        r"C:\Program Files (x86)\CMoney\CMExcel.exe",
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None


def update_cmoney_excel(file_path):
    cm_exe = find_cmexcel()
    if not cm_exe:
        log("找不到 CMExcel.exe（找過 C:\\Program Files\\CMoney 和 (x86) 兩個路徑），略過 CMoney 更新，"
            "直接用現有 Excel 資料重新匯出網站。")
        return False

    log(f"啟動 CMoney 更新（{cm_exe}），過程通常要數十秒，請耐心等待...")
    try:
        arg = CMONEY_UPDATE_ARG_TEMPLATE.format(file_path=file_path)
        subprocess.run([cm_exe, arg])
        log("CMExcel.exe 執行完成。")
    except Exception as e:
        log(f"呼叫 CMExcel.exe 發生錯誤: {e}")
        return False

    # 確保 Excel 存檔並關閉，釋放檔案鎖定，這樣後面 refresh_data.py 才讀得到最新檔案
    log("嘗試存檔並關閉 Excel...")
    try:
        import win32com.client
        excel = win32com.client.GetActiveObject("Excel.Application")
        excel.DisplayAlerts = False
        target_name = os.path.basename(file_path)
        closed = False
        for wb in excel.Workbooks:
            if wb.Name == target_name:
                wb.Save()
                wb.Close()
                closed = True
                break
        if not closed:
            log(f"沒找到開著的活頁簿 {target_name}（也許 CMExcel.exe 自己存檔關閉了），繼續往下跑。")
        if excel.Workbooks.Count == 0:
            excel.Quit()
    except Exception as e:
        log(f"存檔/關閉 Excel 時發生問題（可能本來就沒開，或已經自動關閉了）: {e}")

    # 給檔案系統一點時間釋放檔案鎖
    time.sleep(2)
    return True


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))     # tools/
    project_dir = os.path.dirname(base_dir)                    # 基本面數據網站/
    xlsm_path = os.path.join(project_dir, XLSM_FILENAME)

    if not os.path.exists(xlsm_path):
        log(f"找不到 {xlsm_path}，請確認這個 bat/py 是放在專案資料夾底下的 tools 資料夾裡執行。")
        sys.exit(1)

    log(f"目標檔案: {xlsm_path}")
    update_cmoney_excel(xlsm_path)

    log("開始重新匯出網站資料 (refresh_data.py)...")
    import refresh_data
    sys.argv = ["refresh_data.py", xlsm_path, os.path.join(project_dir, "data")]
    refresh_data.main()
    log("網站資料已經是最新的了（data/chunks/、data/manifest.js）。")

    publish_to_git(project_dir)


def publish_to_git(project_dir):
    """如果這個資料夾已經照 DEPLOY.md 設定過 git remote，就自動推上去，
    Cloudflare Pages 偵測到會自動重新部署。還沒設定的話就跳過，不當成錯誤。"""
    git_dir = os.path.join(project_dir, ".git")
    if not os.path.isdir(git_dir):
        log("這個資料夾還沒做過 git 初始化（見 DEPLOY.md），略過自動發布，網站資料只有更新在本機。")
        return
    try:
        remotes = subprocess.run(["git", "remote"], cwd=project_dir, capture_output=True, text=True)
        if "origin" not in (remotes.stdout or ""):
            log("git 還沒設定 remote origin（見 DEPLOY.md 第一次設定），略過自動發布。")
            return
        log("推送更新到 GitHub（Cloudflare Pages 會自動重新部署）...")
        subprocess.run(["git", "add", "-A"], cwd=project_dir)
        subprocess.run(["git", "commit", "-m", "update data (CMoney auto refresh)"], cwd=project_dir)
        push = subprocess.run(["git", "push"], cwd=project_dir)
        if push.returncode == 0:
            log("推送成功，等 1~2 分鐘 Cloudflare Pages 部署好就能看到最新資料。")
        else:
            log("git push 失敗，請手動執行一次 publish.bat 看詳細錯誤訊息。")
    except Exception as e:
        log(f"自動發布時發生錯誤（不影響本機資料已經更新完成）: {e}")


if __name__ == "__main__":
    main()
