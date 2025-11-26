# 热榜抓取（微博/头条）并保存到 Excel

本项目提供脚本与图形界面，支持抓取微博热搜与今日头条热榜，默认抓取前 30 条，保存到对应的 Excel 文件（`weibo_hot_top30.xlsx` 或 `toutiao_hot_top30.xlsx`）。

## 环境要求
- Python 3.8 或更高版本（Windows/Mac/Linux 均可）
- 依赖库：`requests`、`beautifulsoup4`、`openpyxl`

## 安装依赖
在项目根目录执行：

```bash
python -m pip install -r requirements.txt
```

如果系统没有 `python` 命令，可尝试：

```bash
py -3 -m pip install -r requirements.txt
```

## 运行脚本
在项目根目录执行：

```bash
python weibo_hot.py
```

运行成功后，会在当前目录生成 `weibo_hot_top30.xlsx`，包含以下列：
- 排名
- 标题
- 链接
- 抓取时间

## 使用 Playwright（更稳妥）
Playwright 通过真实浏览器抓取，更能绕过页面动态加载与部分反爬。

安装 Playwright 并准备浏览器：

```bash
python -m pip install -r requirements.txt
python -m playwright install chromium
```

运行 Playwright 版脚本：

```bash
python weibo_hot_playwright.py
```

可选环境变量：
- `WEIBO_COOKIE`：已登录微博的 Cookie 字符串（提升成功率）
- `HTTPS_PROXY` / `HTTP_PROXY`：代理地址，如 `http://127.0.0.1:7890`
- `WEIBO_HEADLESS`：设为 `0` 以打开有界面浏览器（默认无界面）

## 图形界面（选择渠道、数量并一键操作）
运行（建议使用项目虚拟环境）：

```bash
\.venv\Scripts\activate
python weibo_hot_gui.py
```

或直接双击 `run_gui.bat`（强制使用本项目虚拟环境）。

功能：
- 选择“渠道”：`微博` / `头条` / `Reddit` / `Hacker News`
- 设置抓取数量（默认 30，范围 1–50）
- 勾选“无界面模式”以 headless 抓取
- 点击“抓取并保存到Excel”生成渠道对应的文件：
  - 微博：`weibo_hot_top{N}.xlsx`
  - 头条：`toutiao_hot_top{N}.xlsx`
  - Reddit：`reddit_hot_top{N}.xlsx`
  - Hacker News：`hn_hot_top{N}.xlsx`
  - 若同时勾选多个渠道，将合并保存为：`hot_all_top{总条数}.xlsx`
- 点击“抓取并写入飞书”写入多维表（需配置 `FEISHU_PBT`、`FEISHU_APP_TOKEN`、`FEISHU_TABLE_ID`）。写入字段包含：`排名`、`标题`、`链接`、`渠道`、`抓取时间`。
- 点击“设置多维表参数…”弹出配置窗口，填写并保存后将用于手动写入与定时任务。
  

多维表参数设置：
- `AppToken`：Base 文档链接中的 `/base/:app_token` 部分（示例：`VzyDbBfWjaJXoTsEhcfcYSRfnWd`）。
- `TableId`：链接参数 `table=tbl...`（示例：`tbl2UPfeJl47mlPO`）。
- `PersonalBaseToken (PBT)`：在对应 Base 文档中生成的授权码（示例以 `pt-` 开头）。
- 在界面填写并点击“保存设置”，手动写入和定时任务都会使用这里的参数。

### 定时抓取（支持渠道）
- 在 GUI 底部的“定时抓取”区域，设置：
  - 时间：`HH:MM` 24 小时制，例如 `08:00`
  - 频率：`仅一次`、`每天`、`每周`
- 点击“开始定时”后，程序会在后台按设定时间抓取并写入飞书（按当前渠道设置），多次运行会自动安排下一次。
- 点击“停止定时”可终止后台任务并清空本地 `schedule.json`。
- 程序启动时若检测到 `schedule.json` 存在且任务未过期，会自动恢复定时任务。
- 注意：定时任务仅执行“抓取并写入飞书”，需保证已安装 BaseOpenSDK 并配置 `FEISHU_PBT`、`FEISHU_APP_TOKEN`、`FEISHU_TABLE_ID`。

## 抓取说明
- 微博：优先调用 `https://weibo.com/ajax/side/hotSearch` JSON 接口，如不可用则回退解析 `https://s.weibo.com/top/summary?cate=realtimehot` 页面；已过滤广告项与非数字排名条目（如置顶）。
- 头条：在浏览器环境内请求 `https://www.toutiao.com/hot-event/hot-board/?origin=toutiao_pc`，若不可用则解析 `https://www.toutiao.com/hot-event/hotboard/?origin=toutiao_pc` 的 `window.__INITIAL_STATE__`；数据会标准化为统一字段并包含 `渠道=头条`。
 
 - Reddit：优先解析 `r/all` 与 `popular` 接口；
 - Hacker News：优先使用 Firebase API（`topstories` + `item/{id}`），回退解析首页列表。

## 常见问题
- 未获取到数据：可能为网络异常或触发反爬。可稍后重试，或在 `HEADERS` 中补充有效 Cookie / 调整 User-Agent，或使用代理网络。
- `pip`/`python` 不可用：请先安装 Python，并勾选“Add to PATH”。

## 扩展建议
- 定时任务：每日定时抓取并以日期命名文件。
- 保存更多字段：如话题分类、是否新上榜等。
- 并发抓取详情页：补充更丰富的元数据（需注意速率与反爬）。

## 写入飞书多维表
- 依赖：安装 BaseOpenSDK（Python）
  - 命令：`python -m pip install https://lf3-static.bytednsdoc.com/obj/eden-cn/lmeh7phbozvhoz/base-open-sdk/baseopensdk-0.0.13-py3-none-any.whl`
  - 注意：SDK 当前不支持 Python 3.13；本项目使用 Python 3.12。
- 运行写入脚本：
  - `python push_to_feishu.py`
- 可选环境变量（不设置也可使用脚本默认值）：
  - `FEISHU_PBT`：PersonalBaseToken（授权码）。示例：`pt-...`。
  - `FEISHU_APP_TOKEN`：Base 文档的 AppToken（从链接中 `/base/:app_token` 获取）。
  - `FEISHU_TABLE_ID`：目标表的 TableId（从链接中的 `table=tbl...` 获取）。
  - `HTTPS_PROXY` / `HTTP_PROXY`：如需代理访问。
- 字段写入：脚本按以下字段名写入记录：`排名`、`标题`、`链接`、`抓取时间`。确保表中存在这些字段（类型文本/数字均可）。