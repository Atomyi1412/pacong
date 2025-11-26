import sys, os

# 绿色版支持：设置 Playwright 浏览器路径
# 必须在导入 playwright 之前设置
if getattr(sys, 'frozen', False):
    __BASE_DIR = os.path.dirname(sys.executable)
else:
    __BASE_DIR = os.path.dirname(os.path.abspath(__file__))

__local_browsers = os.path.join(__BASE_DIR, "browsers")
if os.path.exists(__local_browsers):
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = __local_browsers

# 仅在非打包环境下启用虚拟环境自动切换
if not getattr(sys, 'frozen', False):
    # Windows
    if sys.platform == 'win32':
        __VENV_PYW = os.path.join(__BASE_DIR, ".venv", "Scripts", "pythonw.exe")
        __VENV_PY = os.path.join(__BASE_DIR, ".venv", "Scripts", "python.exe")
    # macOS / Linux
    else:
        __VENV_PYW = os.path.join(__BASE_DIR, ".venv", "bin", "python")
        __VENV_PY = os.path.join(__BASE_DIR, ".venv", "bin", "python3")

    __target = None
    if os.path.exists(__VENV_PYW):
        __target = __VENV_PYW
    elif os.path.exists(__VENV_PY):
        __target = __VENV_PY

    if __target:
        try:
            is_same = os.path.samefile(sys.executable, __target)
        except (OSError, AttributeError):
            # Windows 下可能没有 samefile (Python 3.2+ 有，但为了兼容性)
            is_same = os.path.normcase(os.path.abspath(sys.executable)) == os.path.normcase(os.path.abspath(__target))

        if not is_same:
            # 额外的检查：如果当前路径看起来像是在 venv 里（例如通过 IDE 启动），也不要重启
            if ".venv" not in sys.executable:
                import subprocess
                subprocess.Popen([__target, os.path.abspath(__file__)] + sys.argv[1:], cwd=__BASE_DIR)
                raise SystemExit(0)
import json
import threading
import uuid
from datetime import datetime, timedelta
from typing import List, Dict

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from playwright.sync_api import sync_playwright
from weibo_hot_playwright import fetch_top_via_api, fetch_top_via_dom, save_to_excel
from toutiao_hot_playwright import fetch_toutiao_via_api, fetch_toutiao_via_dom
from reddit_hot_playwright import fetch_reddit_via_api, fetch_reddit_via_dom
from hn_hot_playwright import fetch_hn_via_api, fetch_hn_via_dom

# 可选：写入飞书所需配置（支持环境变量）
try:
    from baseopensdk import BaseClient
    from baseopensdk.api.base.v1.model.app_table_record import AppTableRecord
    from baseopensdk.api.base.v1.model.create_app_table_record_request import (
        CreateAppTableRecordRequest,
    )
    BASEOPENSDK_AVAILABLE = True
except Exception:
    BASEOPENSDK_AVAILABLE = False

# 系统托盘支持（可选）
try:
    import pystray  # type: ignore
    from PIL import Image, ImageDraw  # type: ignore
    TRAY_AVAILABLE = True
except Exception:
    TRAY_AVAILABLE = False

APP_TOKEN = os.environ.get("FEISHU_APP_TOKEN", "")
TABLE_ID = os.environ.get("FEISHU_TABLE_ID", "")
PBT = os.environ.get("FEISHU_PBT", "")

# 使用脚本所在目录作为持久化文件路径，避免工作目录不一致导致读写不同文件
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCHEDULE_FILE = os.path.join(BASE_DIR, "schedule.json")
SCHEDULES_FILE = os.path.join(BASE_DIR, "schedules.json")  # 多任务版持久化文件
FEISHU_CONFIG_PATH = os.path.join(BASE_DIR, "feishu_config.json")

def load_feishu_config() -> Dict:
    try:
        with open(FEISHU_CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_feishu_config(cfg: Dict) -> None:
    try:
        with open(FEISHU_CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def scrape_items(limit: int, headless: bool = True, channel: str = "微博", google_geo: str | None = None) -> List[Dict]:
    """根据渠道抓取数据。微博优先API，失败回退DOM；头条优先API，失败回退DOM。"""
    with sync_playwright() as p:
        ch = (channel or "微博").strip()
        if ch == "头条":
            items = fetch_toutiao_via_api(p.chromium, headless=headless, limit=limit)
            if not items:
                items = fetch_toutiao_via_dom(p.chromium, headless=headless, limit=limit)
            # 标注渠道
            for it in items or []:
                if "channel" not in it:
                    it["channel"] = "头条"
        elif ch == "Reddit":
            items = fetch_reddit_via_api(p.chromium, headless=headless, limit=limit)
            if not items:
                items = fetch_reddit_via_dom(p.chromium, headless=headless, limit=limit)
            for it in items or []:
                if "channel" not in it:
                    it["channel"] = "Reddit"
        elif ch == "Google Trends":
            items = fetch_google_trends_via_api(p.chromium, headless=headless, limit=limit, geo=(google_geo or "US"))
            for it in items or []:
                if "channel" not in it:
                    it["channel"] = "Google Trends"
        elif ch == "Hacker News":
            items = fetch_hn_via_api(p.chromium, headless=headless, limit=limit)
            if not items:
                items = fetch_hn_via_dom(p.chromium, headless=headless, limit=limit)
            for it in items or []:
                if "channel" not in it:
                    it["channel"] = "Hacker News"
        else:
            items = fetch_top_via_api(p.chromium, headless=headless, limit=limit)
            if not items:
                items = fetch_top_via_dom(p.chromium, headless=headless, limit=limit)
            # 标注渠道
            for it in items:
                if "channel" not in it:
                    it["channel"] = "微博"
    return items


def write_to_feishu(items: List[Dict], app_token: str, table_id: str, pbt: str) -> int:
    if not BASEOPENSDK_AVAILABLE:
        raise RuntimeError("BaseOpenSDK 未安装，无法写入飞书。")
    client = BaseClient.builder().app_token(app_token).personal_base_token(pbt).build()
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ok = 0
    for it in items:
        fields = {
            "排名": it.get("rank"),
            "标题": it.get("title"),
            "链接": it.get("link"),
            "渠道": it.get("channel") or "微博",
            "抓取时间": ts,
        }
        body = AppTableRecord.builder().fields(fields).build()
        req = (
            CreateAppTableRecordRequest
            .builder()
            .app_token(app_token)
            .table_id(table_id)
            .request_body(body)
            .build()
        )
        resp = client.base.v1.app_table_record.create(req)
        data = getattr(resp, "data", None)
        rec = getattr(data, "record", None) if data else None
        if rec and getattr(rec, "record_id", None):
            ok += 1
    return ok


class HotGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("热榜抓取（微博/头条）")
        self.root.geometry("620x360")

        self.headless_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="就绪")
        # 渠道独立配置：是否抓取 + 数量
        self.weibo_enabled_var = tk.BooleanVar(value=True)
        self.weibo_limit_var = tk.IntVar(value=30)
        self.toutiao_enabled_var = tk.BooleanVar(value=False)
        self.toutiao_limit_var = tk.IntVar(value=30)
        self.reddit_enabled_var = tk.BooleanVar(value=False)
        self.reddit_limit_var = tk.IntVar(value=30)
        self.youtube_enabled_var = tk.BooleanVar(value=False)
        self.youtube_limit_var = tk.IntVar(value=30)
        self.youtube_geo_var = tk.StringVar(value="US")
        self.hn_enabled_var = tk.BooleanVar(value=False)
        self.hn_limit_var = tk.IntVar(value=30)
        # 载入飞书配置（优先使用本地文件，其次环境变量）
        cfg = load_feishu_config()
        self.app_token_var = tk.StringVar(value=cfg.get("app_token", APP_TOKEN))
        self.table_id_var = tk.StringVar(value=cfg.get("table_id", TABLE_ID))
        self.pbt_var = tk.StringVar(value=cfg.get("pbt", PBT))

        # 定时相关变量
        self.time_var = tk.StringVar(value="08:00")
        self.freq_var = tk.StringVar(value="每天")
        self.schedule_timer: threading.Timer | None = None
        self.schedule_running = False
        self.schedule_conf: Dict | None = None

        frm = ttk.Frame(root, padding=16)
        frm.pack(fill=tk.BOTH, expand=True)

        # 微博行：渠道名—勾选框—抓取数量
        ttk.Label(frm, text="微博").grid(row=0, column=0, sticky=tk.W)
        self.weibo_chk = ttk.Checkbutton(frm, text="抓取", variable=self.weibo_enabled_var)
        self.weibo_chk.grid(row=0, column=1, sticky=tk.W)
        ttk.Label(frm, text="数量").grid(row=0, column=2, sticky=tk.W)
        self.weibo_spin = ttk.Spinbox(frm, from_=1, to=50, textvariable=self.weibo_limit_var, width=8)
        self.weibo_spin.grid(row=0, column=3, sticky=tk.W, padx=6)
        # 头条行
        ttk.Label(frm, text="头条").grid(row=1, column=0, sticky=tk.W)
        self.toutiao_chk = ttk.Checkbutton(frm, text="抓取", variable=self.toutiao_enabled_var)
        self.toutiao_chk.grid(row=1, column=1, sticky=tk.W)
        ttk.Label(frm, text="数量").grid(row=1, column=2, sticky=tk.W)
        self.toutiao_spin = ttk.Spinbox(frm, from_=1, to=50, textvariable=self.toutiao_limit_var, width=8)
        self.toutiao_spin.grid(row=1, column=3, sticky=tk.W, padx=6)
        # Reddit 行
        ttk.Label(frm, text="Reddit").grid(row=2, column=0, sticky=tk.W)
        self.reddit_chk = ttk.Checkbutton(frm, text="抓取", variable=self.reddit_enabled_var)
        self.reddit_chk.grid(row=2, column=1, sticky=tk.W)
        ttk.Label(frm, text="数量").grid(row=2, column=2, sticky=tk.W)
        self.reddit_spin = ttk.Spinbox(frm, from_=1, to=50, textvariable=self.reddit_limit_var, width=8)
        self.reddit_spin.grid(row=2, column=3, sticky=tk.W, padx=6)
        # （已移除）Google Trends 选项
        # Hacker News 行
        ttk.Label(frm, text="Hacker News").grid(row=4, column=0, sticky=tk.W)
        self.hn_chk = ttk.Checkbutton(frm, text="抓取", variable=self.hn_enabled_var)
        self.hn_chk.grid(row=4, column=1, sticky=tk.W)
        ttk.Label(frm, text="数量").grid(row=4, column=2, sticky=tk.W)
        self.hn_spin = ttk.Spinbox(frm, from_=1, to=50, textvariable=self.hn_limit_var, width=8)
        self.hn_spin.grid(row=4, column=3, sticky=tk.W, padx=6)

        # 无界面模式固定启用（不在界面显示）

        self.btn_excel = ttk.Button(frm, text="抓取并保存到Excel", command=self.on_excel)
        self.btn_excel.grid(row=5, column=0, columnspan=3, sticky=tk.EW, pady=12)

        self.btn_feishu = ttk.Button(frm, text="抓取并写入飞书", command=self.on_feishu)
        self.btn_feishu.grid(row=5, column=3, columnspan=2, sticky=tk.EW, pady=12)

        # 多维表参数设置（弹窗）
        self.btn_settings = ttk.Button(frm, text="设置多维表参数…", command=self.open_feishu_settings_dialog)
        self.btn_settings.grid(row=6, column=0, columnspan=5, sticky=tk.EW)

        ttk.Separator(frm).grid(row=7, column=0, columnspan=5, sticky=tk.EW, pady=8)
        ttk.Label(frm, textvariable=self.status_var, foreground="#555").grid(row=7, column=0, columnspan=5, sticky=tk.W)
        # 一键收至右下角（系统托盘）
        self.btn_tray = ttk.Button(frm, text="收至右下角", command=self.minimize_to_tray)
        self.btn_tray.grid(row=8, column=4, sticky=tk.E, pady=4)

        frm.columnconfigure(0, weight=1)
        frm.columnconfigure(1, weight=1)
        frm.columnconfigure(2, weight=1)
        frm.columnconfigure(3, weight=1)
        frm.columnconfigure(4, weight=1)
        frm.columnconfigure(5, weight=1)

        # === 定时抓取区域 ===
        ttk.Separator(frm).grid(row=9, column=0, columnspan=5, sticky=tk.EW, pady=10)
        ttk.Label(frm, text="定时抓取").grid(row=10, column=0, sticky=tk.W)

        # 仅保留两个入口按钮
        self.btn_add_sched = ttk.Button(frm, text="新增定时任务", command=self.add_schedule_from_ui)
        self.btn_add_sched.grid(row=11, column=0, columnspan=3, sticky=tk.EW, pady=10)
        self.btn_manage_sched = ttk.Button(frm, text="查看定时任务…", command=self.open_schedules_manager_dialog)
        self.btn_manage_sched.grid(row=11, column=3, columnspan=2, sticky=tk.EW, pady=10)

        # 内存中的任务与线程管理
        self.tasks: List[Dict] = []
        self.task_threads: Dict[str, threading.Thread] = {}
        self.task_events: Dict[str, threading.Event] = {}

        # 尝试恢复并启动既有定时任务（多任务版）
        self.restore_schedules()
        # 托盘图标引用
        self.tray_icon = None

    def set_busy(self, busy: bool):
        state = (tk.DISABLED if busy else tk.NORMAL)
        self.btn_excel.configure(state=state)
        self.btn_feishu.configure(state=state)
        self.btn_settings.configure(state=state)
        # 渠道控件联动
        self.weibo_chk.configure(state=state)
        self.weibo_spin.configure(state=state)
        self.toutiao_chk.configure(state=state)
        self.toutiao_spin.configure(state=state)
        self.reddit_chk.configure(state=state)
        self.reddit_spin.configure(state=state)
        # Google Trends 控件已移除
        self.hn_chk.configure(state=state)
        self.hn_spin.configure(state=state)
        # 定时按钮在手动抓取时仍可用，无需联动

    # ===== 系统托盘相关 =====
    def _create_tray_image(self):
        if not TRAY_AVAILABLE:
            return None
        img = Image.new("RGBA", (64, 64), (255, 255, 255, 0))
        d = ImageDraw.Draw(img)
        # 简洁图形：蓝色圆 + 白色羽毛形状（抽象）
        d.ellipse((8, 8, 56, 56), fill=(0, 122, 255, 255))
        d.rectangle((30, 20, 34, 44), fill=(255, 255, 255, 255))
        d.polygon([(26, 44), (32, 52), (38, 44)], fill=(255, 255, 255, 255))
        return img

    def minimize_to_tray(self):
        if not TRAY_AVAILABLE:
            messagebox.showwarning("提示", "托盘功能需要安装 pystray 与 Pillow。\n可在环境中运行: pip install pystray pillow")
            # 回退为任务栏最小化，方便恢复
            self.root.iconify()
            return

        # 隐藏主窗口并创建托盘图标
        self.root.withdraw()
        img = self._create_tray_image()

        def on_show(icon, item):
            try:
                icon.stop()
            except Exception:
                pass
            self.root.after(0, self.restore_from_tray)

        def on_exit(icon, item):
            try:
                icon.stop()
            except Exception:
                pass
            self.root.after(0, self.root.destroy)

        menu = pystray.Menu(
            pystray.MenuItem("显示窗口", on_show),
            pystray.MenuItem("退出", on_exit)
        )
        self.tray_icon = pystray.Icon("HotGUI", img, "热榜抓取（微博/头条）", menu)
        try:
            self.tray_icon.run_detached()
        except Exception:
            # 失败则恢复窗口，避免不可见
            self.root.deiconify()

    def restore_from_tray(self):
        try:
            if self.tray_icon:
                self.tray_icon.stop()
        except Exception:
            pass
        self.tray_icon = None
        self.root.deiconify()
        try:
            self.root.focus_force()
        except Exception:
            pass

    def on_excel(self):
        headless = bool(self.headless_var.get())
        weibo_enabled = bool(self.weibo_enabled_var.get())
        toutiao_enabled = bool(self.toutiao_enabled_var.get())
        weibo_limit = max(1, int(self.weibo_limit_var.get()))
        toutiao_limit = max(1, int(self.toutiao_limit_var.get()))
        reddit_enabled = bool(self.reddit_enabled_var.get())
        reddit_limit = max(1, int(self.reddit_limit_var.get()))
        # 已移除 Google Trends 选项
        hn_enabled = bool(self.hn_enabled_var.get())
        hn_limit = max(1, int(self.hn_limit_var.get()))

        if not (weibo_enabled or toutiao_enabled or reddit_enabled or hn_enabled):
            messagebox.showwarning("提示", "请至少勾选一个渠道进行抓取。")
            return

        def run():
            try:
                self.status_var.set("抓取中...")
                all_items: List[Dict] = []
                if weibo_enabled:
                    self.status_var.set("抓取微博中...")
                    wb_items = scrape_items(limit=weibo_limit, headless=headless, channel="微博")
                    all_items.extend(wb_items or [])
                if toutiao_enabled:
                    self.status_var.set("抓取头条中...")
                    tt_items = scrape_items(limit=toutiao_limit, headless=headless, channel="头条")
                    all_items.extend(tt_items or [])
                if reddit_enabled:
                    self.status_var.set("抓取Reddit中...")
                    rd_items = scrape_items(limit=reddit_limit, headless=headless, channel="Reddit")
                    all_items.extend(rd_items or [])
                # Google Trends 抓取已移除
                if hn_enabled:
                    self.status_var.set("抓取Hacker News中...")
                    hn_items = scrape_items(limit=hn_limit, headless=headless, channel="Hacker News")
                    all_items.extend(hn_items or [])
                if not all_items:
                    self.status_var.set("未获取到数据")
                    messagebox.showwarning("提示", "未获取到热榜数据，可能需要登录或网络受限。")
                    return
                # 保存文件名：同时勾选则合并保存为 hot_all_topX；仅单渠道保持原命名
                enabled_count = sum([weibo_enabled, toutiao_enabled, reddit_enabled, hn_enabled])
                if enabled_count > 1:
                    total = len(all_items)
                    out = f"hot_all_top{total}.xlsx"
                elif weibo_enabled:
                    out = f"weibo_hot_top{weibo_limit}.xlsx"
                elif toutiao_enabled:
                    out = f"toutiao_hot_top{toutiao_limit}.xlsx"
                elif reddit_enabled:
                    out = f"reddit_hot_top{reddit_limit}.xlsx"
                # Google Trends 文件命名已移除
                else:
                    out = f"hn_hot_top{hn_limit}.xlsx"
                save_to_excel(all_items, path=out)
                self.status_var.set(f"已保存: {os.path.abspath(out)}")
                messagebox.showinfo("完成", f"已保存至\n{os.path.abspath(out)}")
            except Exception as e:
                self.status_var.set("错误")
                messagebox.showerror("错误", str(e))
            finally:
                self.set_busy(False)

        self.set_busy(True)
        threading.Thread(target=run, daemon=True).start()

    def on_feishu(self):
        headless = bool(self.headless_var.get())
        weibo_enabled = bool(self.weibo_enabled_var.get())
        toutiao_enabled = bool(self.toutiao_enabled_var.get())
        weibo_limit = max(1, int(self.weibo_limit_var.get()))
        toutiao_limit = max(1, int(self.toutiao_limit_var.get()))
        reddit_enabled = bool(self.reddit_enabled_var.get())
        reddit_limit = max(1, int(self.reddit_limit_var.get()))
        # 已移除 Google Trends 选项
        hn_enabled = bool(self.hn_enabled_var.get())
        hn_limit = max(1, int(self.hn_limit_var.get()))

        if not (weibo_enabled or toutiao_enabled or reddit_enabled or hn_enabled):
            messagebox.showwarning("提示", "请至少勾选一个渠道进行抓取。")
            return

        def run():
            try:
                if not BASEOPENSDK_AVAILABLE:
                    raise RuntimeError("BaseOpenSDK 未安装或不可用。")
                app_token = self.app_token_var.get().strip()
                table_id = self.table_id_var.get().strip()
                pbt = self.pbt_var.get().strip()
                if not app_token or not table_id or not pbt:
                    raise RuntimeError("缺少 AppToken / TableId / PBT 配置，请在界面填写并保存。")
                self.status_var.set("抓取中...")
                all_items: List[Dict] = []
                if weibo_enabled:
                    wb_items = scrape_items(limit=weibo_limit, headless=headless, channel="微博")
                    all_items.extend(wb_items or [])
                if toutiao_enabled:
                    tt_items = scrape_items(limit=toutiao_limit, headless=headless, channel="头条")
                    all_items.extend(tt_items or [])
                if reddit_enabled:
                    rd_items = scrape_items(limit=reddit_limit, headless=headless, channel="Reddit")
                    all_items.extend(rd_items or [])
                # Google Trends 抓取已移除
                if hn_enabled:
                    hn_items = scrape_items(limit=hn_limit, headless=headless, channel="Hacker News")
                    all_items.extend(hn_items or [])
                if not all_items:
                    self.status_var.set("未获取到数据")
                    messagebox.showwarning("提示", "未获取到热搜数据，可能需要登录或网络受限。")
                    return
                self.status_var.set("写入飞书中...")
                ok = write_to_feishu(all_items, app_token=app_token, table_id=table_id, pbt=pbt)
                self.status_var.set(f"成功写入 {ok} 条记录")
                messagebox.showinfo("完成", f"成功写入 {ok} 条记录到飞书多维表。")
            except Exception as e:
                self.status_var.set("错误")
                messagebox.showerror("错误", str(e))
            finally:
                self.set_busy(False)

        self.set_busy(True)
        threading.Thread(target=run, daemon=True).start()

    def on_save_config(self):
        # 保留兼容方法，改由弹窗调用
        cfg = {
            "app_token": self.app_token_var.get().strip(),
            "table_id": self.table_id_var.get().strip(),
            "pbt": self.pbt_var.get().strip(),
        }
        save_feishu_config(cfg)
        messagebox.showinfo("已保存", "已保存多维表参数设置。")

    # ===== 定时抓取逻辑 =====
    def _parse_time(self, s: str) -> tuple[int, int]:
        try:
            hh, mm = s.strip().split(":")
            h = int(hh)
            m = int(mm)
            if not (0 <= h <= 23 and 0 <= m <= 59):
                raise ValueError
            return h, m
        except Exception:
            raise ValueError("时间格式应为 HH:MM，范围 00:00–23:59")

    def _next_delay_seconds(self, freq: str, time_str: str, start_weekday: int | None = None) -> float:
        now = datetime.now()
        h, m = self._parse_time(time_str)
        today_run = now.replace(hour=h, minute=m, second=0, microsecond=0)

        if freq == "仅一次":
            # 若时间已过，安排到明天
            target = today_run if today_run > now else (today_run + timedelta(days=1))
            return (target - now).total_seconds()
        elif freq == "每天":
            target = today_run if today_run > now else (today_run + timedelta(days=1))
            return (target - now).total_seconds()
        elif freq == "每周":
            # 以开始当天为基准的周循环
            base_weekday = start_weekday if start_weekday is not None else now.weekday()
            days_delta = (base_weekday - now.weekday()) % 7
            target_date = now.date() + timedelta(days=days_delta)
            target = datetime.combine(target_date, today_run.time())
            if target <= now:
                target += timedelta(days=7)
            return (target - now).total_seconds()
        else:
            # 回退为每天
            target = today_run if today_run > now else (today_run + timedelta(days=1))
            return (target - now).total_seconds()

    # ===== 旧版单任务定时入口（保留但不再在界面使用） =====
    def on_start_schedule(self):
        if self.schedule_running:
            messagebox.showinfo("提示", "定时任务已在运行。")
            return
        try:
            freq = self.freq_var.get().strip()
            time_str = self.time_var.get().strip()
            self._parse_time(time_str)

            if not BASEOPENSDK_AVAILABLE:
                raise RuntimeError("BaseOpenSDK 未安装或不可用，无法写入飞书。")
            app_token = self.app_token_var.get().strip()
            table_id = self.table_id_var.get().strip()
            pbt = self.pbt_var.get().strip()
            if not app_token or not table_id or not pbt:
                raise RuntimeError("缺少 AppToken / TableId / PBT 配置，请在界面填写并保存。")
            headless = bool(self.headless_var.get())
            start_weekday = datetime.now().weekday()
            weibo_enabled = bool(self.weibo_enabled_var.get())
            toutiao_enabled = bool(self.toutiao_enabled_var.get())
            weibo_limit = max(1, int(self.weibo_limit_var.get()))
            toutiao_limit = max(1, int(self.toutiao_limit_var.get()))
            reddit_enabled = bool(self.reddit_enabled_var.get())
            reddit_limit = max(1, int(self.reddit_limit_var.get()))
            # 已移除 Google Trends 选项
            hn_enabled = bool(self.hn_enabled_var.get())
            hn_limit = max(1, int(self.hn_limit_var.get()))
            if not (weibo_enabled or toutiao_enabled or reddit_enabled or hn_enabled):
                raise RuntimeError("请至少勾选一个渠道进行抓取。")

            # 持久化配置
            conf = {
                "time": time_str,
                "freq": freq,
                "headless": headless,
                "start_weekday": start_weekday,
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "app_token": app_token,
                "table_id": table_id,
                "pbt": pbt,
                "weibo_enabled": weibo_enabled,
                "weibo_limit": weibo_limit,
                "toutiao_enabled": toutiao_enabled,
                "toutiao_limit": toutiao_limit,
                "reddit_enabled": reddit_enabled,
                "reddit_limit": reddit_limit,
                # Google Trends 配置已移除
                "hn_enabled": hn_enabled,
                "hn_limit": hn_limit,
            }
            try:
                with open(SCHEDULE_FILE, "w", encoding="utf-8") as f:
                    json.dump(conf, f, ensure_ascii=False, indent=2)
            except Exception:
                # 非致命，继续运行
                pass

            delay = self._next_delay_seconds(freq, time_str, start_weekday)
            self.schedule_conf = conf
            self.schedule_running = True
            self.status_var.set(f"定时已开启：{freq} {time_str}，下一次 {int(delay)} 秒后")

            def _fire():
                self._run_scheduled_once()
                # 仅一次则停止并清空
                if self.schedule_conf and self.schedule_conf.get("freq") == "仅一次":
                    self.on_stop_schedule(clear_file=True)
                else:
                    # 继续安排下一次
                    if self.schedule_running:
                        next_delay = self._next_delay_seconds(
                            self.schedule_conf.get("freq", "每天"),
                            self.schedule_conf.get("time", "08:00"),
                            self.schedule_conf.get("start_weekday")
                        )
                        self.status_var.set(f"上次已运行，下一次 {int(next_delay)} 秒后")
                        self.schedule_timer = threading.Timer(next_delay, _fire)
                        self.schedule_timer.daemon = True
                        self.schedule_timer.start()

            # 首次安排
            self.schedule_timer = threading.Timer(delay, _fire)
            self.schedule_timer.daemon = True
            self.schedule_timer.start()

        except Exception as e:
            messagebox.showerror("错误", str(e))

    def _run_scheduled_once(self):
        conf = self.schedule_conf or {}
        headless = bool(conf.get("headless", bool(self.headless_var.get())))
        app_token = conf.get("app_token") or self.app_token_var.get().strip()
        table_id = conf.get("table_id") or self.table_id_var.get().strip()
        pbt = conf.get("pbt") or self.pbt_var.get().strip()
        weibo_enabled = bool(conf.get("weibo_enabled", bool(self.weibo_enabled_var.get())))
        toutiao_enabled = bool(conf.get("toutiao_enabled", bool(self.toutiao_enabled_var.get())))
        weibo_limit = int(conf.get("weibo_limit", max(1, int(self.weibo_limit_var.get()))))
        toutiao_limit = int(conf.get("toutiao_limit", max(1, int(self.toutiao_limit_var.get()))))
        reddit_enabled = bool(conf.get("reddit_enabled", bool(self.reddit_enabled_var.get())))
        reddit_limit = int(conf.get("reddit_limit", max(1, int(self.reddit_limit_var.get()))))
        # 已移除 Google Trends 选项
        hn_enabled = bool(conf.get("hn_enabled", bool(self.hn_enabled_var.get())))
        hn_limit = int(conf.get("hn_limit", max(1, int(self.hn_limit_var.get()))))
        # 兼容旧配置（仅单渠道）
        if ("channel" in conf) and ("limit" in conf):
            if str(conf.get("channel")).strip() == "微博":
                weibo_enabled = True
                weibo_limit = int(conf.get("limit", weibo_limit))
                toutiao_enabled = False
                reddit_enabled = False
                # Google Trends 兼容逻辑已移除
                hn_enabled = False
            elif str(conf.get("channel")).strip() == "头条":
                toutiao_enabled = True
                toutiao_limit = int(conf.get("limit", toutiao_limit))
                weibo_enabled = False
                reddit_enabled = False
                # Google Trends 兼容逻辑已移除
                hn_enabled = False

        def run_job():
            try:
                self.status_var.set("定时抓取中...")
                all_items: List[Dict] = []
                if weibo_enabled:
                    wb_items = scrape_items(limit=weibo_limit, headless=headless, channel="微博")
                    all_items.extend(wb_items or [])
                if toutiao_enabled:
                    tt_items = scrape_items(limit=toutiao_limit, headless=headless, channel="头条")
                    all_items.extend(tt_items or [])
                if reddit_enabled:
                    rd_items = scrape_items(limit=reddit_limit, headless=headless, channel="Reddit")
                    all_items.extend(rd_items or [])
                # Google Trends 抓取已移除
                if hn_enabled:
                    hn_items = scrape_items(limit=hn_limit, headless=headless, channel="Hacker News")
                    all_items.extend(hn_items or [])
                if not all_items:
                    self.status_var.set("定时未获取到数据")
                    return
                self.status_var.set("定时写入飞书中...")
                ok = write_to_feishu(all_items, app_token=app_token, table_id=table_id, pbt=pbt)
                self.status_var.set(f"定时成功写入 {ok} 条记录")
            except Exception as e:
                self.status_var.set("定时任务错误")
                # 不弹窗打扰，状态提示即可
            finally:
                pass

        threading.Thread(target=run_job, daemon=True).start()

    def on_stop_schedule(self, clear_file: bool = True):
        # 停止当前定时
        try:
            if self.schedule_timer:
                self.schedule_timer.cancel()
        except Exception:
            pass
        self.schedule_timer = None
        self.schedule_running = False
        self.status_var.set("定时已停止")
        if clear_file:
            try:
                with open(SCHEDULE_FILE, "w", encoding="utf-8") as f:
                    f.write("{}")
            except Exception:
                pass

    def restore_schedule(self):
        # 程序启动时恢复
        try:
            if not os.path.exists(SCHEDULE_FILE):
                return
            with open(SCHEDULE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            time_str = data.get("time")
            freq = data.get("freq")
            if not time_str or not freq:
                return
            # 对于仅一次，若时间已过则不恢复
            now = datetime.now()
            h, m = self._parse_time(time_str)
            today_run = now.replace(hour=h, minute=m, second=0, microsecond=0)
            if freq == "仅一次" and today_run <= now:
                return
            # 恢复配置到界面
            self.time_var.set(time_str)
            self.freq_var.set(freq)
            self.headless_var.set(bool(data.get("headless", self.headless_var.get())))
            # 新配置：渠道独立
            if "weibo_enabled" in data:
                self.weibo_enabled_var.set(bool(data.get("weibo_enabled")))
            if "weibo_limit" in data:
                self.weibo_limit_var.set(int(data.get("weibo_limit", self.weibo_limit_var.get())))
            if "toutiao_enabled" in data:
                self.toutiao_enabled_var.set(bool(data.get("toutiao_enabled")))
            if "toutiao_limit" in data:
                self.toutiao_limit_var.set(int(data.get("toutiao_limit", self.toutiao_limit_var.get())))
            if "reddit_enabled" in data:
                self.reddit_enabled_var.set(bool(data.get("reddit_enabled")))
            if "reddit_limit" in data:
                self.reddit_limit_var.set(int(data.get("reddit_limit", self.reddit_limit_var.get())))
            # Google Trends 恢复逻辑已移除
            if "hn_enabled" in data:
                self.hn_enabled_var.set(bool(data.get("hn_enabled")))
            if "hn_limit" in data:
                self.hn_limit_var.set(int(data.get("hn_limit", self.hn_limit_var.get())))
            # 兼容旧配置：单渠道 + 单数量
            if ("channel" in data) and ("limit" in data):
                ch = str(data.get("channel")).strip()
                lim = max(1, int(data.get("limit", 30)))
                if ch == "微博":
                    self.weibo_enabled_var.set(True)
                    self.weibo_limit_var.set(lim)
                    self.toutiao_enabled_var.set(False)
                elif ch == "头条":
                    self.toutiao_enabled_var.set(True)
                    self.toutiao_limit_var.set(lim)
                    self.weibo_enabled_var.set(False)
                # 旧配置不支持新渠道，保持默认状态
            # 恢复多维表参数
            if data.get("app_token"):
                self.app_token_var.set(data.get("app_token"))
            if data.get("table_id"):
                self.table_id_var.set(data.get("table_id"))
            if data.get("pbt"):
                self.pbt_var.set(data.get("pbt"))
            self.schedule_conf = data
            # 启动定时
            self.on_start_schedule()
        except Exception:
            # 忽略恢复错误
            pass

    # ===== 多任务定时：新增/管理/运行 =====
    def _compute_next_run(self, freq: str, time_str: str, start_weekday: int | None = None) -> datetime:
        """返回下一次运行的实际时间戳（datetime）。"""
        now = datetime.now()
        h, m = self._parse_time(time_str)
        today_run = now.replace(hour=h, minute=m, second=0, microsecond=0)
        if freq == "仅一次":
            target = today_run if today_run > now else (today_run + timedelta(days=1))
        elif freq == "每天":
            target = today_run if today_run > now else (today_run + timedelta(days=1))
        elif freq == "每周":
            base_weekday = start_weekday if start_weekday is not None else now.weekday()
            days_delta = (base_weekday - now.weekday()) % 7
            target_date = now.date() + timedelta(days=days_delta)
            target = datetime.combine(target_date, today_run.time())
            if target <= now:
                target += timedelta(days=7)
        else:
            target = today_run if today_run > now else (today_run + timedelta(days=1))
        return target

    def _task_channels_summary(self, t: Dict) -> str:
        parts = []
        if t.get("weibo_enabled"): parts.append(f"微博({t.get('weibo_limit', 30)})")
        if t.get("toutiao_enabled"): parts.append(f"头条({t.get('toutiao_limit', 30)})")
        if t.get("reddit_enabled"): parts.append(f"Reddit({t.get('reddit_limit', 30)})")
        if t.get("hn_enabled"): parts.append(f"HackerNews({t.get('hn_limit', 30)})")
        return ", ".join(parts) or "(未选择渠道)"

    def _load_tasks(self) -> List[Dict]:
        try:
            if not os.path.exists(SCHEDULES_FILE):
                return []
            with open(SCHEDULES_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                return data
            return []
        except Exception:
            return []

    def _save_tasks(self) -> None:
        try:
            with open(SCHEDULES_FILE, "w", encoding="utf-8") as f:
                json.dump(self.tasks, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def add_schedule_from_ui(self):
        # 打开新增任务弹框，包含渠道、数量、保存到Excel/飞书、飞书参数、时间与频率
        self.open_add_schedule_dialog()

    def _start_task_thread(self, task: Dict):
        tid = task["id"]
        # 若已存在线程，先停止
        self._stop_task_thread(tid)
        stop_event = threading.Event()
        self.task_events[tid] = stop_event

        def runner():
            while not stop_event.is_set():
                # 计算下一次运行
                next_dt = self._compute_next_run(task.get("freq", "每天"), task.get("time", "08:00"), task.get("start_weekday"))
                task["next_run"] = next_dt.strftime("%Y-%m-%d %H:%M:%S")
                self._save_tasks()
                # 等待到时间点或被停止
                now = datetime.now()
                wait_s = max(0.0, (next_dt - now).total_seconds())
                # 使用 stop_event.wait 支持提前停止
                if stop_event.wait(wait_s):
                    break
                # 到点执行一次
                self._run_task_once(task)
                # 一次性任务执行完即完成
                if task.get("freq") == "仅一次":
                    task["status"] = "completed"
                    self._save_tasks()
                    break

        th = threading.Thread(target=runner, daemon=True)
        self.task_threads[tid] = th
        th.start()

    def _stop_task_thread(self, task_id: str):
        ev = self.task_events.get(task_id)
        if ev:
            try:
                ev.set()
            except Exception:
                pass
        self.task_events.pop(task_id, None)
        self.task_threads.pop(task_id, None)

    def _run_task_once(self, task: Dict):
        # 运行一次任务：抓取并写入飞书
        cfg = load_feishu_config()
        # 任务内可覆盖飞书参数
        app_token = task.get("app_token") or cfg.get("app_token") or self.app_token_var.get().strip()
        table_id = task.get("table_id") or cfg.get("table_id") or self.table_id_var.get().strip()
        pbt = task.get("pbt") or cfg.get("pbt") or self.pbt_var.get().strip()
        headless = bool(task.get("headless", bool(self.headless_var.get())))
        weibo_enabled = bool(task.get("weibo_enabled", False))
        toutiao_enabled = bool(task.get("toutiao_enabled", False))
        reddit_enabled = bool(task.get("reddit_enabled", False))
        hn_enabled = bool(task.get("hn_enabled", False))
        weibo_limit = int(task.get("weibo_limit", 30))
        toutiao_limit = int(task.get("toutiao_limit", 30))
        reddit_limit = int(task.get("reddit_limit", 30))
        hn_limit = int(task.get("hn_limit", 30))
        save_excel = bool(task.get("save_excel", False))
        save_feishu = bool(task.get("save_feishu", True))
        excel_dir = task.get("excel_dir") or "."

        def run_job():
            try:
                self.status_var.set(f"任务 {task.get('id')} 抓取中...")
                all_items: List[Dict] = []
                if weibo_enabled:
                    all_items.extend(scrape_items(limit=weibo_limit, headless=headless, channel="微博") or [])
                if toutiao_enabled:
                    all_items.extend(scrape_items(limit=toutiao_limit, headless=headless, channel="头条") or [])
                if reddit_enabled:
                    all_items.extend(scrape_items(limit=reddit_limit, headless=headless, channel="Reddit") or [])
                if hn_enabled:
                    all_items.extend(scrape_items(limit=hn_limit, headless=headless, channel="Hacker News") or [])
                if not all_items:
                    self.status_var.set(f"任务 {task.get('id')} 未获取到数据")
                    return
                # 保存到 Excel（可选）
                if save_excel:
                    try:
                        # 生成文件名（含时间戳）
                        ts = datetime.now().strftime("%Y%m%d_%H%M")
                        enabled_count = sum([weibo_enabled, toutiao_enabled, reddit_enabled, hn_enabled])
                        if enabled_count > 1:
                            out = f"hot_all_{ts}.xlsx"
                        elif weibo_enabled:
                            out = f"weibo_hot_{ts}.xlsx"
                        elif toutiao_enabled:
                            out = f"toutiao_hot_{ts}.xlsx"
                        elif reddit_enabled:
                            out = f"reddit_hot_{ts}.xlsx"
                        else:
                            out = f"hn_hot_{ts}.xlsx"
                        full_path = os.path.abspath(os.path.join(excel_dir, out))
                        save_to_excel(all_items, path=full_path)
                        self.status_var.set(f"任务 {task.get('id')} 已保存Excel: {full_path}")
                    except Exception as e:
                        self.status_var.set(f"任务 {task.get('id')} 保存Excel失败: {e}")

                # 写入飞书（可选）
                if save_feishu:
                    if not BASEOPENSDK_AVAILABLE or not app_token or not table_id or not pbt:
                        self.status_var.set("飞书参数缺失，未写入")
                    else:
                        self.status_var.set(f"任务 {task.get('id')} 写入飞书中...")
                        ok = write_to_feishu(all_items, app_token=app_token, table_id=table_id, pbt=pbt)
                        self.status_var.set(f"任务 {task.get('id')} 成功写入 {ok} 条记录")
                if not save_excel and not save_feishu:
                    self.status_var.set(f"任务 {task.get('id')} 未选择保存方式")
            except Exception:
                self.status_var.set(f"任务 {task.get('id')} 发生错误")

        threading.Thread(target=run_job, daemon=True).start()

    def open_schedules_manager_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("定时任务管理")
        dlg.transient(self.root)
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding=10)
        frm.pack(fill=tk.BOTH, expand=True)

        cols = ("id", "freq", "time", "next_run", "status", "channels")
        tv = ttk.Treeview(frm, columns=cols, show="headings", height=10)
        for c, txt in zip(cols, ["ID", "频率", "时间", "下一次", "状态", "渠道"]):
            tv.heading(c, text=txt)
            tv.column(c, width=100 if c != "channels" else 260, anchor=tk.W)
        tv.pack(fill=tk.BOTH, expand=True)

        # 填充当前任务
        self.tasks = self._load_tasks()
        tv.delete(*tv.get_children())
        for t in self.tasks:
            tv.insert("", tk.END, values=(t.get("id"), t.get("freq"), t.get("time"), t.get("next_run"), t.get("status"), self._task_channels_summary(t)))

        btns = ttk.Frame(frm)
        btns.pack(fill=tk.X, pady=8)

        def refresh_tv():
            self.tasks = self._load_tasks()
            tv.delete(*tv.get_children())
            for t in self.tasks:
                tv.insert("", tk.END, values=(t.get("id"), t.get("freq"), t.get("time"), t.get("next_run"), t.get("status"), self._task_channels_summary(t)))

        def get_selected_task() -> Dict | None:
            sel = tv.selection()
            if not sel:
                messagebox.showwarning("提示", "请先选择一个任务")
                return None
            vals = tv.item(sel[0], "values")
            tid = vals[0]
            for t in self.tasks:
                if t.get("id") == tid:
                    return t
            return None

        def on_view():
            t = get_selected_task()
            if not t:
                return
            details = (
                f"ID: {t.get('id')}\n"
                f"频率: {t.get('freq')}\n"
                f"时间: {t.get('time')}\n"
                f"下一次: {t.get('next_run')}\n"
                f"状态: {t.get('status')}\n"
                f"渠道: {self._task_channels_summary(t)}\n"
            )
            messagebox.showinfo("任务详情", details)

        def on_stop():
            t = get_selected_task()
            if not t:
                return
            self._stop_task_thread(t.get("id"))
            t["status"] = "stopped"
            self._save_tasks()
            refresh_tv()

        def on_run_once():
            t = get_selected_task()
            if not t:
                return
            self._run_task_once(t)

        def on_delete():
            t = get_selected_task()
            if not t:
                return
            self._stop_task_thread(t.get("id"))
            self.tasks = [x for x in self.tasks if x.get("id") != t.get("id")]
            self._save_tasks()
            refresh_tv()

        def on_edit():
            t = get_selected_task()
            if not t:
                return
            self.open_edit_schedule_dialog(t, refresh_tv, dlg)

        ttk.Button(btns, text="查看", command=on_view).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="编辑", command=on_edit).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="停止", command=on_stop).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="执行", command=on_run_once).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="删除", command=on_delete).pack(side=tk.LEFT, padx=4)

    def restore_schedules(self):
        # 加载所有任务并启动线程（仅对状态为 scheduled 的任务）
        self.tasks = self._load_tasks()
        for t in self.tasks:
            if t.get("status") == "scheduled":
                self._start_task_thread(t)

    # ===== 弹框：新增与编辑任务 =====
    def open_add_schedule_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("新增定时任务")
        dlg.transient(self.root)
        dlg.grab_set()
        container = ttk.Frame(dlg, padding=12)
        container.pack(fill=tk.BOTH, expand=True)

        # 行1：渠道选择与数量
        ttk.Label(container, text="渠道与数量").grid(row=0, column=0, sticky=tk.W)
        weibo_var = tk.BooleanVar(value=bool(self.weibo_enabled_var.get()))
        weibo_limit_var = tk.IntVar(value=int(self.weibo_limit_var.get()))
        toutiao_var = tk.BooleanVar(value=bool(self.toutiao_enabled_var.get()))
        toutiao_limit_var = tk.IntVar(value=int(self.toutiao_limit_var.get()))
        reddit_var = tk.BooleanVar(value=bool(self.reddit_enabled_var.get()))
        reddit_limit_var = tk.IntVar(value=int(self.reddit_limit_var.get()))
        hn_var = tk.BooleanVar(value=bool(self.hn_enabled_var.get()))
        hn_limit_var = tk.IntVar(value=int(self.hn_limit_var.get()))

        row = 1
        ttk.Checkbutton(container, text="微博", variable=weibo_var).grid(row=row, column=0, sticky=tk.W)
        ttk.Label(container, text="数量").grid(row=row, column=1, sticky=tk.W)
        ttk.Spinbox(container, from_=1, to=50, textvariable=weibo_limit_var, width=8).grid(row=row, column=2, sticky=tk.W)
        row += 1
        ttk.Checkbutton(container, text="头条", variable=toutiao_var).grid(row=row, column=0, sticky=tk.W)
        ttk.Label(container, text="数量").grid(row=row, column=1, sticky=tk.W)
        ttk.Spinbox(container, from_=1, to=50, textvariable=toutiao_limit_var, width=8).grid(row=row, column=2, sticky=tk.W)
        row += 1
        ttk.Checkbutton(container, text="Reddit", variable=reddit_var).grid(row=row, column=0, sticky=tk.W)
        ttk.Label(container, text="数量").grid(row=row, column=1, sticky=tk.W)
        ttk.Spinbox(container, from_=1, to=50, textvariable=reddit_limit_var, width=8).grid(row=row, column=2, sticky=tk.W)
        row += 1
        ttk.Checkbutton(container, text="Hacker News", variable=hn_var).grid(row=row, column=0, sticky=tk.W)
        ttk.Label(container, text="数量").grid(row=row, column=1, sticky=tk.W)
        ttk.Spinbox(container, from_=1, to=50, textvariable=hn_limit_var, width=8).grid(row=row, column=2, sticky=tk.W)
        # 无界面模式固定启用，不显示控件

        # 行：保存到Excel/飞书
        row += 1
        ttk.Label(container, text="保存到哪里").grid(row=row, column=0, sticky=tk.W)
        save_excel_var = tk.BooleanVar(value=True)
        save_feishu_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(container, text="保存到Excel", variable=save_excel_var).grid(row=row, column=1, sticky=tk.W)
        ttk.Checkbutton(container, text="写入飞书", variable=save_feishu_var).grid(row=row, column=2, sticky=tk.W)
        row += 1
        ttk.Label(container, text="Excel保存目录").grid(row=row, column=0, sticky=tk.W)
        excel_dir_var = tk.StringVar(value=os.getcwd())
        ttk.Entry(container, textvariable=excel_dir_var).grid(row=row, column=1, columnspan=2, sticky=tk.EW)
        def choose_dir():
            d = filedialog.askdirectory()
            if d:
                excel_dir_var.set(d)
        ttk.Button(container, text="选择…", command=choose_dir).grid(row=row, column=3, sticky=tk.W)

        # 行：飞书参数
        row += 1
        ttk.Label(container, text="飞书参数").grid(row=row, column=0, sticky=tk.W)
        cfg = load_feishu_config()
        app_token_var = tk.StringVar(value=cfg.get("app_token", self.app_token_var.get().strip()))
        table_id_var = tk.StringVar(value=cfg.get("table_id", self.table_id_var.get().strip()))
        pbt_var = tk.StringVar(value=cfg.get("pbt", self.pbt_var.get().strip()))
        row += 1
        ttk.Label(container, text="AppToken").grid(row=row, column=0, sticky=tk.W)
        ttk.Entry(container, textvariable=app_token_var).grid(row=row, column=1, columnspan=3, sticky=tk.EW)
        row += 1
        ttk.Label(container, text="TableId").grid(row=row, column=0, sticky=tk.W)
        ttk.Entry(container, textvariable=table_id_var).grid(row=row, column=1, columnspan=3, sticky=tk.EW)
        row += 1
        ttk.Label(container, text="PBT").grid(row=row, column=0, sticky=tk.W)
        ttk.Entry(container, textvariable=pbt_var).grid(row=row, column=1, columnspan=3, sticky=tk.EW)

        # 行：时间与频率
        row += 1
        ttk.Label(container, text="时间(HH:MM)").grid(row=row, column=0, sticky=tk.W)
        time_var = tk.StringVar(value=self.time_var.get().strip())
        ttk.Entry(container, textvariable=time_var, width=10).grid(row=row, column=1, sticky=tk.W)
        ttk.Label(container, text="频率").grid(row=row, column=2, sticky=tk.W)
        freq_var = tk.StringVar(value=self.freq_var.get().strip())
        ttk.Combobox(container, textvariable=freq_var, values=["仅一次", "每天", "每周"], state="readonly", width=8).grid(row=row, column=3, sticky=tk.W)

        # 操作按钮
        row += 1
        btns = ttk.Frame(container)
        btns.grid(row=row, column=0, columnspan=4, sticky=tk.E)
        def save_task():
            try:
                freq = freq_var.get().strip()
                time_str = time_var.get().strip()
                self._parse_time(time_str)
                weibo_enabled = bool(weibo_var.get())
                toutiao_enabled = bool(toutiao_var.get())
                reddit_enabled = bool(reddit_var.get())
                hn_enabled = bool(hn_var.get())
                if not (weibo_enabled or toutiao_enabled or reddit_enabled or hn_enabled):
                    messagebox.showwarning("提示", "请至少勾选一个渠道进行抓取。")
                    return
                task_id = uuid.uuid4().hex[:8]
                start_weekday = datetime.now().weekday()
                next_run = self._compute_next_run(freq, time_str, start_weekday)
                task = {
                    "id": task_id,
                    "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "status": "scheduled",
                    "freq": freq,
                    "time": time_str,
                    "start_weekday": start_weekday,
                    "next_run": next_run.strftime("%Y-%m-%d %H:%M:%S"),
                    "headless": True,
                    "weibo_enabled": weibo_enabled,
                    "weibo_limit": max(1, int(weibo_limit_var.get())),
                    "toutiao_enabled": toutiao_enabled,
                    "toutiao_limit": max(1, int(toutiao_limit_var.get())),
                    "reddit_enabled": reddit_enabled,
                    "reddit_limit": max(1, int(reddit_limit_var.get())),
                    "hn_enabled": hn_enabled,
                    "hn_limit": max(1, int(hn_limit_var.get())),
                    "save_excel": bool(save_excel_var.get()),
                    "excel_dir": excel_dir_var.get().strip() or ".",
                    "save_feishu": bool(save_feishu_var.get()),
                    "app_token": app_token_var.get().strip(),
                    "table_id": table_id_var.get().strip(),
                    "pbt": pbt_var.get().strip(),
                }
                self.tasks.append(task)
                self._save_tasks()
                self._start_task_thread(task)
                messagebox.showinfo("已添加", f"定时任务已添加（ID: {task_id}），下一次：{task['next_run']}")
                dlg.destroy()
            except Exception as e:
                messagebox.showerror("错误", str(e))
        ttk.Button(btns, text="保存", command=save_task).pack(side=tk.RIGHT, padx=6)
        ttk.Button(btns, text="取消", command=dlg.destroy).pack(side=tk.RIGHT)
        container.columnconfigure(1, weight=1)

    def open_edit_schedule_dialog(self, task: Dict, on_saved_refresh=None, parent=None):
        dlg = tk.Toplevel(parent or self.root)
        dlg.title(f"编辑任务 {task.get('id')}")
        dlg.transient(self.root)
        dlg.grab_set()
        container = ttk.Frame(dlg, padding=12)
        container.pack(fill=tk.BOTH, expand=True)

        # 预填变量
        weibo_var = tk.BooleanVar(value=bool(task.get("weibo_enabled", False)))
        weibo_limit_var = tk.IntVar(value=int(task.get("weibo_limit", 30)))
        toutiao_var = tk.BooleanVar(value=bool(task.get("toutiao_enabled", False)))
        toutiao_limit_var = tk.IntVar(value=int(task.get("toutiao_limit", 30)))
        reddit_var = tk.BooleanVar(value=bool(task.get("reddit_enabled", False)))
        reddit_limit_var = tk.IntVar(value=int(task.get("reddit_limit", 30)))
        hn_var = tk.BooleanVar(value=bool(task.get("hn_enabled", False)))
        hn_limit_var = tk.IntVar(value=int(task.get("hn_limit", 30)))
        # 无界面模式固定启用，不显示控件
        save_excel_var = tk.BooleanVar(value=bool(task.get("save_excel", False)))
        save_feishu_var = tk.BooleanVar(value=bool(task.get("save_feishu", True)))
        excel_dir_var = tk.StringVar(value=task.get("excel_dir") or os.getcwd())
        app_token_var = tk.StringVar(value=task.get("app_token") or self.app_token_var.get().strip())
        table_id_var = tk.StringVar(value=task.get("table_id") or self.table_id_var.get().strip())
        pbt_var = tk.StringVar(value=task.get("pbt") or self.pbt_var.get().strip())
        time_var = tk.StringVar(value=task.get("time") or self.time_var.get().strip())
        freq_var = tk.StringVar(value=task.get("freq") or self.freq_var.get().strip())

        # 布局（与新增一致）
        ttk.Label(container, text="渠道与数量").grid(row=0, column=0, sticky=tk.W)
        row = 1
        ttk.Checkbutton(container, text="微博", variable=weibo_var).grid(row=row, column=0, sticky=tk.W)
        ttk.Label(container, text="数量").grid(row=row, column=1, sticky=tk.W)
        ttk.Spinbox(container, from_=1, to=50, textvariable=weibo_limit_var, width=8).grid(row=row, column=2, sticky=tk.W)
        row += 1
        ttk.Checkbutton(container, text="头条", variable=toutiao_var).grid(row=row, column=0, sticky=tk.W)
        ttk.Label(container, text="数量").grid(row=row, column=1, sticky=tk.W)
        ttk.Spinbox(container, from_=1, to=50, textvariable=toutiao_limit_var, width=8).grid(row=row, column=2, sticky=tk.W)
        row += 1
        ttk.Checkbutton(container, text="Reddit", variable=reddit_var).grid(row=row, column=0, sticky=tk.W)
        ttk.Label(container, text="数量").grid(row=row, column=1, sticky=tk.W)
        ttk.Spinbox(container, from_=1, to=50, textvariable=reddit_limit_var, width=8).grid(row=row, column=2, sticky=tk.W)
        row += 1
        ttk.Checkbutton(container, text="Hacker News", variable=hn_var).grid(row=row, column=0, sticky=tk.W)
        ttk.Label(container, text="数量").grid(row=row, column=1, sticky=tk.W)
        ttk.Spinbox(container, from_=1, to=50, textvariable=hn_limit_var, width=8).grid(row=row, column=2, sticky=tk.W)
        # 无界面模式固定启用，不显示控件

        row += 1
        ttk.Label(container, text="保存到哪里").grid(row=row, column=0, sticky=tk.W)
        ttk.Checkbutton(container, text="保存到Excel", variable=save_excel_var).grid(row=row, column=1, sticky=tk.W)
        ttk.Checkbutton(container, text="写入飞书", variable=save_feishu_var).grid(row=row, column=2, sticky=tk.W)
        row += 1
        ttk.Label(container, text="Excel保存目录").grid(row=row, column=0, sticky=tk.W)
        ttk.Entry(container, textvariable=excel_dir_var).grid(row=row, column=1, columnspan=2, sticky=tk.EW)
        def choose_dir2():
            d = filedialog.askdirectory()
            if d:
                excel_dir_var.set(d)
        ttk.Button(container, text="选择…", command=choose_dir2).grid(row=row, column=3, sticky=tk.W)

        row += 1
        ttk.Label(container, text="飞书参数").grid(row=row, column=0, sticky=tk.W)
        row += 1
        ttk.Label(container, text="AppToken").grid(row=row, column=0, sticky=tk.W)
        ttk.Entry(container, textvariable=app_token_var).grid(row=row, column=1, columnspan=3, sticky=tk.EW)
        row += 1
        ttk.Label(container, text="TableId").grid(row=row, column=0, sticky=tk.W)
        ttk.Entry(container, textvariable=table_id_var).grid(row=row, column=1, columnspan=3, sticky=tk.EW)
        row += 1
        ttk.Label(container, text="PBT").grid(row=row, column=0, sticky=tk.W)
        ttk.Entry(container, textvariable=pbt_var).grid(row=row, column=1, columnspan=3, sticky=tk.EW)

        row += 1
        ttk.Label(container, text="时间(HH:MM)").grid(row=row, column=0, sticky=tk.W)
        ttk.Entry(container, textvariable=time_var, width=10).grid(row=row, column=1, sticky=tk.W)
        ttk.Label(container, text="频率").grid(row=row, column=2, sticky=tk.W)
        ttk.Combobox(container, textvariable=freq_var, values=["仅一次", "每天", "每周"], state="readonly", width=8).grid(row=row, column=3, sticky=tk.W)

        row += 1
        btns = ttk.Frame(container)
        btns.grid(row=row, column=0, columnspan=4, sticky=tk.E)
        def save_edit():
            try:
                new_time = time_var.get().strip()
                self._parse_time(new_time)
                new_freq = freq_var.get().strip()
                task["time"] = new_time
                task["freq"] = new_freq
                task["headless"] = True
                task["weibo_enabled"] = bool(weibo_var.get())
                task["weibo_limit"] = max(1, int(weibo_limit_var.get()))
                task["toutiao_enabled"] = bool(toutiao_var.get())
                task["toutiao_limit"] = max(1, int(toutiao_limit_var.get()))
                task["reddit_enabled"] = bool(reddit_var.get())
                task["reddit_limit"] = max(1, int(reddit_limit_var.get()))
                task["hn_enabled"] = bool(hn_var.get())
                task["hn_limit"] = max(1, int(hn_limit_var.get()))
                task["save_excel"] = bool(save_excel_var.get())
                task["excel_dir"] = excel_dir_var.get().strip() or "."
                task["save_feishu"] = bool(save_feishu_var.get())
                task["app_token"] = app_token_var.get().strip()
                task["table_id"] = table_id_var.get().strip()
                task["pbt"] = pbt_var.get().strip()
                # 重启线程以应用新配置
                self._start_task_thread(task)
                self._save_tasks()
                if on_saved_refresh:
                    on_saved_refresh()
                dlg.destroy()
            except Exception as e:
                messagebox.showerror("错误", str(e))
        ttk.Button(btns, text="保存", command=save_edit).pack(side=tk.RIGHT, padx=6)
        ttk.Button(btns, text="取消", command=dlg.destroy).pack(side=tk.RIGHT)
        container.columnconfigure(1, weight=1)

    # ===== 多维表参数弹窗 =====
    def open_feishu_settings_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("多维表参数设置")
        dlg.transient(self.root)
        dlg.grab_set()
        container = ttk.Frame(dlg, padding=16)
        container.pack(fill=tk.BOTH, expand=True)

        # 1. 解析链接区域
        ttk.Label(container, text="多维表链接（自动解析）").grid(row=0, column=0, sticky=tk.W)
        link_var = tk.StringVar()
        entry_link = ttk.Entry(container, textvariable=link_var, width=48)
        entry_link.grid(row=0, column=1, sticky=tk.EW, padx=8)
        
        def parse_link():
            url = link_var.get().strip()
            if not url:
                return
            try:
                # 简单解析逻辑
                # 格式: .../base/<app_token>?table=<table_id>...
                import re
                # 匹配 app_token (base/ 之后，? 之前)
                token_match = re.search(r"/base/([a-zA-Z0-9]+)", url)
                if token_match:
                    self.app_token_var.set(token_match.group(1))
                
                # 匹配 table_id (table= 之后，& 或结束之前)
                table_match = re.search(r"table=([a-zA-Z0-9]+)", url)
                if table_match:
                    self.table_id_var.set(table_match.group(1))
                
                if token_match or table_match:
                    messagebox.showinfo("解析成功", "已从链接提取 AppToken / TableId")
                else:
                    messagebox.showwarning("解析失败", "未从链接中找到有效信息，请检查链接格式")
            except Exception as e:
                messagebox.showerror("错误", f"解析出错: {e}")

        btn_parse = ttk.Button(container, text="解析", command=parse_link)
        btn_parse.grid(row=0, column=2, sticky=tk.W)

        ttk.Separator(container).grid(row=1, column=0, columnspan=3, sticky=tk.EW, pady=10)

        # 2. 详细参数区域
        ttk.Label(container, text="AppToken").grid(row=2, column=0, sticky=tk.W)
        entry_app = ttk.Entry(container, textvariable=self.app_token_var, width=48)
        entry_app.grid(row=2, column=1, sticky=tk.EW, padx=8)

        ttk.Label(container, text="TableId").grid(row=3, column=0, sticky=tk.W)
        entry_tbl = ttk.Entry(container, textvariable=self.table_id_var, width=48)
        entry_tbl.grid(row=3, column=1, sticky=tk.EW, padx=8)

        ttk.Label(container, text="PersonalBaseToken (PBT)").grid(row=4, column=0, sticky=tk.W)
        entry_pbt = ttk.Entry(container, textvariable=self.pbt_var, width=48)
        entry_pbt.grid(row=4, column=1, sticky=tk.EW, padx=8)

        help_text = (
            "说明：\n"
            "1. 复制多维表浏览器地址栏链接到上方，点击“解析”可自动填充 AppToken 和 TableId。\n"
            "2. PBT (PersonalBaseToken) 仍需手动获取（在多维表右上角“扩展脚本”中生成）。"
        )
        ttk.Label(container, text=help_text, foreground="#666", wraplength=480, justify=tk.LEFT).grid(
            row=5, column=0, columnspan=3, sticky=tk.W, pady=(10, 0)
        )

        btns = ttk.Frame(container)
        btns.grid(row=6, column=0, columnspan=3, sticky=tk.E, pady=(10, 0))
        def save_and_close():
            cfg = {
                "app_token": self.app_token_var.get().strip(),
                "table_id": self.table_id_var.get().strip(),
                "pbt": self.pbt_var.get().strip(),
            }
            save_feishu_config(cfg)
            messagebox.showinfo("已保存", "已保存多维表参数设置。")
            dlg.destroy()
        ttk.Button(btns, text="保存", command=save_and_close).pack(side=tk.RIGHT, padx=6)
        ttk.Button(btns, text="取消", command=dlg.destroy).pack(side=tk.RIGHT)

        container.columnconfigure(1, weight=1)


def main():
    root = tk.Tk()
    HotGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()