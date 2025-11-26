import os
import sys
import time
from datetime import datetime
from typing import List, Dict

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from urllib.parse import urljoin, quote


HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/129.0 Safari/537.36"
    ),
    "Accept": "text/html,application/json;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Referer": "https://weibo.com/",
}

# 可选：通过环境变量注入 Cookie 与代理，增强在被反爬场景下的可用性
COOKIE = os.environ.get("WEIBO_COOKIE")
if COOKIE:
    HEADERS["Cookie"] = COOKIE

PROXIES = {
    "http": os.environ.get("HTTP_PROXY") or os.environ.get("http_proxy"),
    "https": os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy"),
}
if not PROXIES["http"] and not PROXIES["https"]:
    PROXIES = None


def fetch_hot_via_api() -> List[Dict]:
    """
    优先使用微博公开的侧边热搜接口（返回 JSON），解析前30条。
    如果接口不可用将抛出异常，由上层处理。
    """
    url = "https://weibo.com/ajax/side/hotSearch"
    r = requests.get(url, headers=HEADERS, timeout=10, proxies=PROXIES)
    r.raise_for_status()
    data = r.json()
    realtime = data.get("realtime", [])

    results = []
    for item in realtime:
        # 过滤广告、置顶等非普通热搜项
        if item.get("is_ad"):
            continue
        word = item.get("word") or ""
        if not word:
            continue
        rank = item.get("rank")
        # 构建可点击搜索链接
        link = f"https://s.weibo.com/weibo?q={quote(word)}"
        results.append({
            "rank": rank if rank else len(results) + 1,
            "title": word,
            "link": link,
        })
        if len(results) >= 30:
            break

    return results


def fetch_hot_via_html() -> List[Dict]:
    """
    解析 s.weibo.com 热搜汇总页面的 HTML，提取前30条。
    作为 JSON 接口不可用时的回退方案。
    """
    url = "https://s.weibo.com/top/summary?cate=realtimehot"
    r = requests.get(url, headers=HEADERS, timeout=10, proxies=PROXIES)
    r.raise_for_status()
    r.encoding = "utf-8"
    soup = BeautifulSoup(r.text, "html.parser")

    rows = soup.select("#pl_top_realtimehot table tbody tr")
    results = []
    for tr in rows:
        tds = tr.find_all("td")
        if len(tds) < 3:
            continue
        rank_text = tds[0].get_text(strip=True)
        # 跳过置顶/标题等非数字排名的行
        if not rank_text.isdigit():
            continue
        a = tds[1].find("a")
        if not a:
            continue
        title = a.get_text(strip=True)
        href = a.get("href", "")
        if href.startswith("//"):
            link = "https:" + href
        else:
            link = urljoin("https://s.weibo.com", href)

        results.append({
            "rank": int(rank_text),
            "title": title,
            "link": link,
        })
        if len(results) >= 30:
            break

    return results


def fetch_hot_via_mirror() -> List[Dict]:
    """
    通过公开的内容镜像服务抓取 JSON 接口，绕过部分地区/登录限制。
    若返回非 JSON 或解析失败则回退空列表。
    """
    url = "https://r.jina.ai/http://weibo.com/ajax/side/hotSearch"
    r = requests.get(url, headers=HEADERS, timeout=10, proxies=PROXIES)
    r.raise_for_status()
    txt = r.text.strip()
    try:
        data = requests.utils.json.loads(txt)
    except Exception:
        # 有时该服务会返回页面文本而非 JSON
        return []

    realtime = data.get("realtime", [])
    results = []
    for item in realtime:
        if item.get("is_ad"):
            continue
        word = item.get("word") or ""
        if not word:
            continue
        rank = item.get("rank")
        link = f"https://s.weibo.com/weibo?q={quote(word)}"
        results.append({
            "rank": rank if rank else len(results) + 1,
            "title": word,
            "link": link,
        })
        if len(results) >= 30:
            break

    return results

def get_hot_top30() -> List[Dict]:
    """获取微博热搜前30，优先 JSON 接口，失败则回退 HTML。"""
    try:
        items = fetch_hot_via_api()
        if items:
            return items
    except Exception:
        pass

    try:
        items = fetch_hot_via_html()
        return items
    except Exception:
        pass

    # 最后尝试镜像服务
    try:
        items = fetch_hot_via_mirror()
        return items
    except Exception:
        return []


def save_to_excel(items: List[Dict], path: str = "weibo_hot_top30.xlsx") -> str:
    """保存为 Excel（xlsx），并设置表头与基本列宽。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Top30"
    ws.append(["排名", "标题", "链接", "抓取时间"])

    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for it in items:
        ws.append([
            it.get("rank"),
            it.get("title"),
            it.get("link"),
            ts,
        ])

    # 冻结首行
    ws.freeze_panes = "A2"

    # 简单按内容长度设置列宽
    dims = {}
    for row in ws.rows:
        for cell in row:
            if cell.value:
                dims[cell.column_letter] = max(
                    dims.get(cell.column_letter, 0), len(str(cell.value))
                )
    for col, width in dims.items():
        ws.column_dimensions[col].width = min(width * 1.2, 80)

    wb.save(path)
    return path


def main() -> int:
    items = get_hot_top30()
    if not items:
        print("未获取到热搜数据，可能被反爬或网络异常。")
        return 1

    out = save_to_excel(items)
    print(f"已保存至: {os.path.abspath(out)}")
    return 0


if __name__ == "__main__":
    # 简单的重试机制，避免瞬时网络抖动
    for i in range(2):
        code = main()
        if code == 0:
            sys.exit(0)
        time.sleep(1.0)
    sys.exit(1)