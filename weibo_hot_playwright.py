import os
from datetime import datetime
from typing import List, Dict

from openpyxl import Workbook
from urllib.parse import quote
from playwright.sync_api import sync_playwright, BrowserType


def parse_cookie_string(cookie_str: str, domain: str) -> List[Dict]:
    cookies = []
    if not cookie_str:
        return cookies
    parts = cookie_str.split(";")
    for p in parts:
        p = p.strip()
        if not p or "=" not in p:
            continue
        name, value = p.split("=", 1)
        cookies.append({
            "name": name.strip(),
            "value": value.strip(),
            "domain": domain,
            "path": "/",
            "httpOnly": False,
            "secure": True,
        })
    return cookies


def save_to_excel(items: List[Dict], path: str = "weibo_hot_top30.xlsx") -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "Top30"
    ws.append(["排名", "标题", "链接", "渠道", "抓取时间"])
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for it in items:
        ws.append([
            it.get("rank"),
            it.get("title"),
            it.get("link"),
            it.get("channel") or "微博",
            ts,
        ])
    ws.freeze_panes = "A2"
    dims = {}
    for row in ws.rows:
        for cell in row:
            if cell.value:
                dims[cell.column_letter] = max(dims.get(cell.column_letter, 0), len(str(cell.value)))
    for col, width in dims.items():
        ws.column_dimensions[col].width = min(width * 1.2, 80)
    wb.save(path)
    return path


def fetch_top_via_dom(browser_type: BrowserType, headless: bool = True, limit: int = 30) -> List[Dict]:
    proxy_url = os.environ.get("HTTPS_PROXY") or os.environ.get("HTTP_PROXY")
    proxy = {"server": proxy_url} if proxy_url else None

    with browser_type.launch(headless=headless, proxy=proxy) as browser:
        ua = (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/129.0 Safari/537.36"
        )
        context = browser.new_context(
            locale="zh-CN",
            viewport={"width": 1280, "height": 800},
            user_agent=ua,
        )

        cookie_str = os.environ.get("WEIBO_COOKIE")
        if cookie_str:
            # 注入到两域名，提升访问成功率
            cookies = []
            cookies.extend(parse_cookie_string(cookie_str, ".weibo.com"))
            cookies.extend(parse_cookie_string(cookie_str, ".s.weibo.com"))
            if cookies:
                context.add_cookies(cookies)

        page = context.new_page()
        page.goto("https://s.weibo.com/top/summary?cate=realtimehot", wait_until="domcontentloaded")
        page.wait_for_selector("#pl_top_realtimehot table tbody tr", timeout=8000)

        rows = page.locator("#pl_top_realtimehot table tbody tr")
        count = rows.count()
        items: List[Dict] = []
        for i in range(count):
            tr = rows.nth(i)
            tds = tr.locator("td")
            if tds.count() < 3:
                continue
            rank_text = tds.nth(0).inner_text().strip()
            # 跳过非数字排名（如置顶/标题行）
            if not rank_text.isdigit():
                continue
            title_link = tds.nth(1).locator("a")
            if title_link.count() == 0:
                continue
            title = title_link.inner_text().strip()
            href = title_link.get_attribute("href") or ""
            if href.startswith("//"):
                link = "https:" + href
            else:
                link = href if href.startswith("http") else f"https://s.weibo.com{href}"
            items.append({
                "rank": int(rank_text),
                "title": title,
                "link": link,
            })
            if len(items) >= limit:
                break

        context.close()
        return items


def fetch_top_via_api(browser_type: BrowserType, headless: bool = True, limit: int = 30) -> List[Dict]:
    proxy_url = os.environ.get("HTTPS_PROXY") or os.environ.get("HTTP_PROXY")
    proxy = {"server": proxy_url} if proxy_url else None

    with browser_type.launch(headless=headless, proxy=proxy) as browser:
        context = browser.new_context()
        cookie_str = os.environ.get("WEIBO_COOKIE")
        if cookie_str:
            cookies = parse_cookie_string(cookie_str, ".weibo.com")
            if cookies:
                context.add_cookies(cookies)
        # 用页面的 fetch 在浏览器环境内请求，继承上下文 Cookie
        page = context.new_page()
        page.goto("https://weibo.com", wait_until="domcontentloaded")
        js = """
        (async () => {
          try {
            const res = await fetch('https://weibo.com/ajax/side/hotSearch', {
              headers: { 'Accept': 'application/json' }
            });
            if(!res.ok) return null;
            return await res.json();
          } catch(e){ return null }
        })();
        """
        data = page.evaluate(js)
        context.close()
        results: List[Dict] = []
        if not data or "realtime" not in data:
            return results
        for item in data.get("realtime", []):
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
            if len(results) >= limit:
                break
        return results


def main() -> int:
    headless_env = os.environ.get("WEIBO_HEADLESS", "1").strip()
    headless = headless_env not in ("0", "false", "False")

    with sync_playwright() as p:
        # 先尝试 API，失败则回退到 DOM 解析
        items = fetch_top_via_api(p.chromium, headless=headless)
        if not items:
            items = fetch_top_via_dom(p.chromium, headless=headless)
        if not items:
            print("未获取到热搜数据，可能需要登录或网络受限。")
            return 1
        out = save_to_excel(items)
        print(out)
        return 0


if __name__ == "__main__":
    raise SystemExit(main())