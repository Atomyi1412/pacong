import os
import requests
from typing import List, Dict

from playwright.sync_api import BrowserType


def fetch_hn_via_api(browser_type: BrowserType, headless: bool = True, limit: int = 30) -> List[Dict]:
    """通过官方 Firebase API 获取 HN Top Stories。"""
    try:
        ids = requests.get("https://hacker-news.firebaseio.com/v0/topstories.json", timeout=10).json()
    except Exception:
        return []
    items: List[Dict] = []
    for idx, story_id in enumerate(ids[:limit], start=1):
        try:
            data = requests.get(f"https://hacker-news.firebaseio.com/v0/item/{story_id}.json", timeout=10).json()
        except Exception:
            continue
        title = data.get("title")
        url = data.get("url") or f"https://news.ycombinator.com/item?id={story_id}"
        if title and url:
            items.append({
                "rank": idx,
                "title": title,
                "link": url,
                "channel": "Hacker News",
            })
    return items


def fetch_hn_via_dom(browser_type: BrowserType, headless: bool = True, limit: int = 30) -> List[Dict]:
    """解析 HN 首页列表，提取标题与得分。"""
    proxy_url = os.environ.get("HTTPS_PROXY") or os.environ.get("HTTP_PROXY")
    proxy = {"server": proxy_url} if proxy_url else None
    with browser_type.launch(headless=headless, proxy=proxy) as browser:
        context = browser.new_context(locale="en-US")
        page = context.new_page()
        page.goto("https://news.ycombinator.com/", wait_until="domcontentloaded")
        items: List[Dict] = []
        try:
            page.wait_for_selector("tr.athing", timeout=12000)
            rows = page.locator("tr.athing")
            count = rows.count()
            n = min(count, limit)
            # 仅获取标题与链接，忽略得分
            for i in range(n):
                r = rows.nth(i)
                a = r.locator("span.titleline a")
                title = a.inner_text().strip() if a.count() else ""
                href = a.get_attribute("href") if a.count() else None
                if title and href:
                    items.append({
                        "rank": i + 1,
                        "title": title,
                        "link": href,
                        "channel": "Hacker News",
                    })
            context.close()
            return items
        except Exception:
            context.close()
            return []