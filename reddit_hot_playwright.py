import os
import json
from typing import List, Dict

from playwright.sync_api import BrowserType


def _normalize_reddit_items(children: List[Dict], limit: int = 30) -> List[Dict]:
    items: List[Dict] = []
    for idx, child in enumerate(children[:limit], start=1):
        data = child.get("data", {})
        title = data.get("title")
        permalink = data.get("permalink") or ""
        link = ("https://www.reddit.com" + permalink) if permalink.startswith("/") else data.get("url")
        if title and link:
            items.append({
                "rank": idx,
                "title": title,
                "link": link,
                "channel": "Reddit",
            })
    return items


def fetch_reddit_via_api(browser_type: BrowserType, headless: bool = True, limit: int = 30) -> List[Dict]:
    """通过 Reddit JSON 接口抓取 r/all 热帖。使用浏览器环境请求以绕过部分防护。"""
    proxy_url = os.environ.get("HTTPS_PROXY") or os.environ.get("HTTP_PROXY")
    proxy = {"server": proxy_url} if proxy_url else None
    with browser_type.launch(headless=headless, proxy=proxy) as browser:
        context = browser.new_context(
            locale="en-US",
            viewport={"width": 1280, "height": 800},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/129.0 Safari/537.36"
            ),
        )
        page = context.new_page()
        try:
            url = f"https://www.reddit.com/r/all/hot.json?limit={limit}"
            page.goto(url, wait_until="domcontentloaded")
            txt = page.evaluate("() => document.body.innerText")
            data = json.loads(txt)
            children = data.get("data", {}).get("children", [])
            return _normalize_reddit_items(children, limit)
        except Exception:
            return []
        finally:
            context.close()


def fetch_reddit_via_dom(browser_type: BrowserType, headless: bool = True, limit: int = 30) -> List[Dict]:
    """备用方式：使用 popular.json 接口。"""
    proxy_url = os.environ.get("HTTPS_PROXY") or os.environ.get("HTTP_PROXY")
    proxy = {"server": proxy_url} if proxy_url else None
    with browser_type.launch(headless=headless, proxy=proxy) as browser:
        context = browser.new_context(locale="en-US")
        page = context.new_page()
        try:
            url = f"https://www.reddit.com/r/popular.json?limit={limit}"
            page.goto(url, wait_until="domcontentloaded")
            txt = page.evaluate("() => document.body.innerText")
            data = json.loads(txt)
            children = data.get("data", {}).get("children", [])
            return _normalize_reddit_items(children, limit)
        except Exception:
            return []
        finally:
            context.close()