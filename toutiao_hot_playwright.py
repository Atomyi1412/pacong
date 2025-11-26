import os
from datetime import datetime
from typing import List, Dict

from openpyxl import Workbook
from urllib.parse import urljoin
from playwright.sync_api import BrowserType


def _normalize_toutiao_items(data: Dict, limit: int) -> List[Dict]:
    items: List[Dict] = []
    if not data:
        return items
    # 兼容不同结构：优先 data 数组，其次嵌套 hotEvent/hotBoard
    arr = None
    if isinstance(data, dict):
        arr = data.get("data")
        if not arr:
            hot_event = data.get("hotEvent") or {}
            hot_board = hot_event.get("hotBoard") or {}
            arr = hot_board.get("data")

    if not isinstance(arr, list):
        return items

    for idx, rec in enumerate(arr):
        if not isinstance(rec, dict):
            continue
        title = rec.get("Title") or rec.get("title") or rec.get("Query") or rec.get("query")
        url = rec.get("Url") or rec.get("url") or rec.get("Link") or rec.get("link")
        if not title:
            continue
        if url and not url.startswith("http"):
            url = urljoin("https://www.toutiao.com/", url)
        items.append({
            "rank": rec.get("Rank") or rec.get("rank") or idx + 1,
            "title": title,
            "link": url or "",
            "channel": "头条",
        })
        if len(items) >= limit:
            break
    return items


def fetch_toutiao_via_api(browser_type: BrowserType, headless: bool = True, limit: int = 30) -> List[Dict]:
    """
    在浏览器环境内直接请求头条热榜 API（签名由前端生成，浏览器请求更稳妥）。
    兼容结构：返回对象含 data 数组，或 window.__INITIAL_STATE__ 的 hotEvent.hotBoard.data。
    """
    proxy_url = os.environ.get("HTTPS_PROXY") or os.environ.get("HTTP_PROXY")
    proxy = {"server": proxy_url} if proxy_url else None

    with browser_type.launch(headless=headless, proxy=proxy) as browser:
        context = browser.new_context()
        page = context.new_page()
        # 直接调用 JSON 接口（一般会包含 _signature，浏览器环境下可返回数据）
        js = """
        (async () => {
          try {
            const url = 'https://www.toutiao.com/hot-event/hot-board/?origin=toutiao_pc';
            const res = await fetch(url, { headers: { 'Accept': 'application/json' } });
            if (!res.ok) return null;
            return await res.json();
          } catch (e) { return null; }
        })();
        """
        page.goto("https://www.toutiao.com/", wait_until="domcontentloaded")
        data = page.evaluate(js)
        if not data:
            # 尝试访问可视化页面并读取 window.__INITIAL_STATE__
            page.goto("https://www.toutiao.com/hot-event/hotboard/?origin=toutiao_pc", wait_until="domcontentloaded")
            data = page.evaluate("(() => { try { return window.__INITIAL_STATE__ || null } catch(e){ return null } })()")
        context.close()
        return _normalize_toutiao_items(data, limit)


def fetch_toutiao_via_dom(browser_type: BrowserType, headless: bool = True, limit: int = 30) -> List[Dict]:
    """
    进入热榜展示页，优先解析 window.__INITIAL_STATE__；若不可用可在后续迭代补充 DOM 解析。
    """
    proxy_url = os.environ.get("HTTPS_PROXY") or os.environ.get("HTTP_PROXY")
    proxy = {"server": proxy_url} if proxy_url else None

    with browser_type.launch(headless=headless, proxy=proxy) as browser:
        page = browser.new_page()
        page.goto("https://www.toutiao.com/hot-event/hotboard/?origin=toutiao_pc", wait_until="domcontentloaded")
        state = page.evaluate("(() => { try { return window.__INITIAL_STATE__ || null } catch(e){ return null } })()")
        return _normalize_toutiao_items(state or {}, limit)