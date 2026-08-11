"""
把 data/*.json（每檔股票一個檔案）打包成網頁要用的載入檔。

不用單一一支 all_data.js 的原因：Cloudflare Pages 免費方案單一檔案上限 25MB，265 檔股票全塞一支
all_data.js 會到 34MB、超過上限，部署會失敗。改成切成多個 chunk（每個約 2MB），網頁用
data/chunks_index.js 記錄有哪些 chunk 檔，動態把它們都插入 <script> 標籤載入——一樣是本機 <script src>
載入方式，file:// 雙擊打開跟部署到網路上兩種情境都適用，不會被 CORS 擋。

輸出：
  data/manifest.js       window.STOCK_MANIFEST = {...}   （股票代號/名稱清單，給搜尋用）
  data/manifest.json     跟上面內容一樣，純 JSON 版本（給人或其他程式看）
  data/chunks/chunk_XXXX.js   每個檔案 Object.assign 到 window.STOCK_DATA
  data/chunks_index.js   window.STOCK_CHUNKS = ['chunks/chunk_0001.js', ...]
"""
import os
import json
import datetime
import shutil

TARGET_CHUNK_BYTES = 2_000_000  # 每個 chunk 大約抓 2MB 上下（遠低於 Cloudflare 25MB 上限）


def build(data_dir):
    files = sorted(fn for fn in os.listdir(data_dir) if fn.endswith('.json') and fn != 'manifest.json')

    items = []
    raws = []  # (code, raw_json_text)
    for fn in files:
        code = fn[:-5]
        path = os.path.join(data_dir, fn)
        with open(path, encoding='utf-8') as f:
            raw = f.read()
        try:
            dd = json.loads(raw)
        except Exception:
            continue
        items.append({'code': dd['code'], 'name': dd['name'],
                      'latest_pe': dd['latest'].get('fwd_pe'), 'latest_pb': dd['latest'].get('pb')})
        raws.append((code, raw))

    # manifest
    manifest_obj = {'generated_at': datetime.datetime.now().isoformat(), 'stocks': items}
    with open(os.path.join(data_dir, 'manifest.json'), 'w', encoding='utf-8') as f:
        json.dump(manifest_obj, f, ensure_ascii=False, indent=1)
    with open(os.path.join(data_dir, 'manifest.js'), 'w', encoding='utf-8') as f:
        f.write('window.STOCK_MANIFEST = ' + json.dumps(manifest_obj, ensure_ascii=False) + ';\n')

    # chunks
    chunks_dir = os.path.join(data_dir, 'chunks')
    if os.path.exists(chunks_dir):
        shutil.rmtree(chunks_dir)
    os.makedirs(chunks_dir, exist_ok=True)

    chunk_files = []
    buf = []
    buf_size = 0
    chunk_idx = 1

    def flush():
        nonlocal buf, buf_size, chunk_idx
        if not buf:
            return
        fname = f'chunk_{chunk_idx:04d}.js'
        with open(os.path.join(chunks_dir, fname), 'w', encoding='utf-8') as f:
            f.write('Object.assign((window.STOCK_DATA = window.STOCK_DATA || {}), {\n')
            f.write(',\n'.join(buf))
            f.write('\n});\n')
        chunk_files.append('chunks/' + fname)
        chunk_idx += 1
        buf = []
        buf_size = 0

    for code, raw in raws:
        entry = json.dumps(code, ensure_ascii=False) + ':' + raw
        buf.append(entry)
        buf_size += len(entry)
        if buf_size >= TARGET_CHUNK_BYTES:
            flush()
    flush()

    with open(os.path.join(data_dir, 'chunks_index.js'), 'w', encoding='utf-8') as f:
        f.write('window.STOCK_CHUNKS = ' + json.dumps(chunk_files, ensure_ascii=False) + ';\n')

    # 舊版殘留的單一大檔，若還在就清掉，避免跟新架構搞混
    old_all_data = os.path.join(data_dir, 'all_data.js')
    if os.path.exists(old_all_data):
        os.remove(old_all_data)

    return {'stocks': len(items), 'chunks': len(chunk_files)}


if __name__ == '__main__':
    import sys
    data_dir = sys.argv[1] if len(sys.argv) > 1 else 'data'
    result = build(data_dir)
    print(f"完成：{result['stocks']} 檔股票，切成 {result['chunks']} 個 chunk 檔案。")
