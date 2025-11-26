import os
import time
from datetime import datetime
from typing import List, Dict, Optional

import requests
from playwright.sync_api import sync_playwright
from baseopensdk import BaseClient
from baseopensdk.api.base.v1.model.app_table_record import AppTableRecord
from baseopensdk.api.base.v1.model.create_app_table_record_request import (
    CreateAppTableRecordRequest,
)
from feishu_utils import ensure_fields_exist

# ===== 配置区域（支持环境变量覆盖） =====
# 从链接中提取的 AppToken 与 TableId
APP_TOKEN = os.environ.get("FEISHU_APP_TOKEN", "")
TABLE_ID = os.environ.get("FEISHU_TABLE_ID", "")

# PersonalBaseToken（授权码），优先取环境变量
PBT = os.environ.get("FEISHU_PBT", "")  # 如果未设置，将尝试使用用户传入的参数

# 代理（如需要）
HTTPS_PROXY = os.environ.get("HTTPS_PROXY")
HTTP_PROXY = os.environ.get("HTTP_PROXY")
PROXIES = None
if HTTPS_PROXY or HTTP_PROXY:
    PROXIES = {"https": HTTPS_PROXY, "http": HTTP_PROXY}


def fetch_hot_top10() -> List[Dict]:
    """使用 Playwright 从浏览器环境抓取微博热搜前10（JSON优先、DOM回退）。"""
    from weibo_hot_playwright import fetch_top_via_api, fetch_top_via_dom
    items: List[Dict] = []
    with sync_playwright() as p:
        # 无界面运行；如需可通过环境变量 WEIBO_HEADLESS=0 切到有界面
        headless_env = os.environ.get("WEIBO_HEADLESS", "1").strip()
        headless = headless_env not in ("0", "false", "False")
        items = fetch_top_via_api(p.chromium, headless=headless)
        if not items:
            items = fetch_top_via_dom(p.chromium, headless=headless)
    return items



def create_record(client, fields, app_token, table_id):
    """使用 BaseOpenSDK 通过 PersonalBaseToken 写入一条记录。返回响应字典或 None。"""
    try:
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
        # BaseOpenSDK 返回对象，包含 raw 与 data；尝试转为 dict
        data = getattr(resp, 'data', None)
        if data is None:
            # 兜底：返回 raw 文本
            raw = getattr(resp, 'raw', None)
            if raw is not None:
                return {"raw_status": raw.status_code, "raw_text": str(raw.text)[:400]}
            return None
        # data.record 为 AppTableRecord
        rec = getattr(data, 'record', None)
        if rec is not None and getattr(rec, 'record_id', None):
            return {"record_id": rec.record_id}
        return {"data": data.__dict__}
    except Exception as e:
        print(f"SDK 写入异常: {e}")
        return None


def push_items_to_bitable(items: List[Dict], app_token: str = APP_TOKEN, table_id: str = TABLE_ID, pbt: str = PBT) -> int:
    """将抓取到的 items 批量写入飞书多维表。返回成功条数。"""
    if not items:
        print("未获取到热搜数据，写入终止。")
        return 0

    # 确保有配置
    if not app_token or not table_id or not pbt:
         print("缺少 AppToken / TableId / 授权码 (PBT)，无法推送。")
         return 0

    # 创建 client
    try:
        client = BaseClient.builder().app_token(app_token).personal_base_token(pbt).build()
    except Exception as e:
        print(f"初始化飞书客户端失败: {e}")
        return 0

    # 字段检查与自动创建
    field_types = {}
    if len(items) > 0:
        first_item = items[0]
        # 构造样本数据，包含所有可能写入的字段
        # 注意：这里我们尝试使用更丰富的字段类型
        sample_fields = {
            "排名": first_item.get("rank"),
            "标题": first_item.get("title"),
            "链接": {"text": first_item.get("title"), "link": first_item.get("link")}, # 尝试使用超链接类型
            "抓取时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        # 如果有额外字段
        if "hot_value" in first_item:
            sample_fields["热度"] = first_item["hot_value"]
        if "channel" in first_item:
            sample_fields["渠道"] = first_item["channel"]
            
        field_types = ensure_fields_exist(client, app_token, table_id, sample_fields)

    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success = 0
    for it in items:
        fields = {
            "排名": it.get("rank"),
            "标题": it.get("title"),
            "链接": it.get("link"), # 默认文本
            "抓取时间": ts,
        }
        
        # 如果是超链接类型(15)，则使用对象格式
        if field_types.get("链接") == 15:
            fields["链接"] = {"text": it.get("title"), "link": it.get("link")}
            
        if "hot_value" in it:
            fields["热度"] = it["hot_value"]
        if "channel" in it:
            fields["渠道"] = it["channel"]

        resp = create_record(client, fields, app_token, table_id)
        if resp is not None:
            success += 1
        # 简单的节流，避免触发接口限频（PBT 单文档 2qps）
        time.sleep(0.6)
    return success


def main() -> int:
    if not APP_TOKEN or not TABLE_ID or not PBT:
        print("缺少 AppToken / TableId / 授权码 (PBT)，请设置环境变量或在脚本中配置。")
        return 1

    items = fetch_hot_top10()
    if not items:
        print("未获取到热搜数据，可能受限于登录或网络。")
        return 1

    ok = push_items_to_bitable(items)
    print(f"成功写入 {ok} 条记录到飞书多维表。")
    return 0 if ok > 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())