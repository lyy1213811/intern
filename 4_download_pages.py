import os
import sys
import re
import time
import argparse
import requests
import pandas as pd
from urllib.parse import urlparse, parse_qs
from concurrent.futures import ThreadPoolExecutor, as_completed

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7',
}

def sanitize_filename(name):
    # 替换 Windows 文件名非法字符
    return re.sub(r'[\\/*?:"<>|]', '_', name).strip()

def get_url_slug(url):
    try:
        parsed = urlparse(url)
        params = parse_qs(parsed.query)
        if 'programme' in params:
            return params['programme'][0]
        slug = os.path.basename(parsed.path.rstrip('/'))
        return slug if slug else "page"
    except Exception:
        return "page"

def download_single_page(item, html_dir, session):
    url = item['url']
    filename = item['filename']
    filepath = os.path.join(html_dir, filename)

    # 断点续传检查：如果文件已存在且大小大于 5KB，则直接跳过
    if os.path.exists(filepath) and os.path.getsize(filepath) > 5000:
        return {
            "status": "skipped",
            "url": url,
            "filename": filename,
            "filepath": filepath,
            "size": os.path.getsize(filepath)
        }

    retries = 3
    for attempt in range(retries):
        try:
            resp = session.get(url, headers=HEADERS, timeout=25)
            if resp.status_code == 200 and len(resp.text) > 1000:
                with open(filepath, "w", encoding="utf-8", errors="replace") as f:
                    f.write(resp.text)
                return {
                    "status": "success",
                    "url": url,
                    "filename": filename,
                    "filepath": filepath,
                    "size": len(resp.text)
                }
            elif resp.status_code == 404:
                return {
                    "status": "not_found",
                    "url": url,
                    "filename": filename,
                    "filepath": filepath,
                    "error": f"HTTP {resp.status_code}"
                }
        except Exception as e:
            if attempt == retries - 1:
                return {
                    "status": "failed",
                    "url": url,
                    "filename": filename,
                    "filepath": filepath,
                    "error": str(e)
                }
            time.sleep(1.5)

    return {
        "status": "failed",
        "url": url,
        "filename": filename,
        "filepath": filepath,
        "error": "Max retries reached"
    }

def download_pages(school="hku", max_workers=5):
    data_dir = os.path.join("data", school.lower())
    html_dir = os.path.join(data_dir, "htmls")
    os.makedirs(html_dir, exist_ok=True)

    # 优先从 diff_programs.csv 中读取包含 db_id 的对齐清单，若无则回退读取 official_list.csv
    diff_file = os.path.join(data_dir, "diff_programs.csv")
    official_file = os.path.join(data_dir, "official_list.csv")

    targets = []
    seen_urls = set()

    if os.path.exists(diff_file):
        print(f"[*] 读取大盘比对表: {diff_file}")
        df = pd.read_csv(diff_file)
        for _, row in df.iterrows():
            url = str(row.get('official_url', '')).strip()
            if not url or url.lower() == 'nan' or url in seen_urls:
                continue
            seen_urls.add(url)
            
            db_id = row.get('db_id', '')
            slug = get_url_slug(url)
            
            fname = f"{sanitize_filename(slug)}.html"
            if pd.notna(db_id) and str(db_id).strip():
                try:
                    db_id_int = int(float(str(db_id).strip()))
                    fname = f"{db_id_int}_{sanitize_filename(slug)}.html"
                except Exception:
                    pass
                
            targets.append({
                "url": url,
                "filename": fname,
                "name": row.get('official_name', slug)
            })
    elif os.path.exists(official_file):
        print(f"[*] 读取官网列表: {official_file}")
        df = pd.read_csv(official_file)
        for _, row in df.iterrows():
            url = str(row.get('official_url', '')).strip()
            if not url or url.lower() == 'nan' or url in seen_urls:
                continue
            seen_urls.add(url)
            slug = get_url_slug(url)
            fname = f"{sanitize_filename(slug)}.html"
            targets.append({
                "url": url,
                "filename": fname,
                "name": row.get('official_program_name', slug)
            })
    else:
        raise FileNotFoundError(f"未找到 {diff_file} 或 {official_file}，请先执行前序步骤！")

    total_count = len(targets)
    print(f"[*] 准备抓取详情页总数: {total_count} 个")
    print(f"[*] 存储目录: {html_dir}")
    print(f"[*] 并发工作线程数: {max_workers}（带断点续传）")

    session = requests.Session()
    success_cnt = 0
    skipped_cnt = 0
    failed_cnt = 0

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_map = {
            executor.submit(download_single_page, item, html_dir, session): item
            for item in targets
        }
        
        idx = 0
        for future in as_completed(future_map):
            idx += 1
            res = future.result()
            status = res["status"]
            if status == "skipped":
                skipped_cnt += 1
                if idx % 20 == 0 or idx == total_count:
                    print(f"[{idx}/{total_count}] [SKIP] 已存在快照: {res['filename']} ({res['size']//1024} KB)")
            elif status == "success":
                success_cnt += 1
                print(f"[{idx}/{total_count}] [OK] 下载成功: {res['filename']} ({res['size']//1024} KB)")
            else:
                failed_cnt += 1
                print(f"[{idx}/{total_count}] [FAIL] 下载失败: {res['filename']} 原因: {res.get('error')}")

    print(f"\n[OK] 批量下载完成！")
    print(f"    - 新下载成功: {success_cnt} 个")
    print(f"    - 本地已存在(续传跳过): {skipped_cnt} 个")
    print(f"    - 失败/异常: {failed_cnt} 个")
    print(f"    - 快照目录: {html_dir}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="步骤4：批量下载专业详情页 HTML 快照")
    parser.add_argument("--school", default="hku", help="院校代码，如 hku")
    parser.add_argument("--workers", type=int, default=5, help="并发线程数，默认 5")
    args = parser.parse_args()

    download_pages(args.school, args.workers)
