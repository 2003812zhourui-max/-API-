"""通用 OMP Excel 导出模块

支持通过不同 templateId / templateConfig 导出各类报表：
  - 出库单 (delivery)
  - 入库单 (receiving/ASN)
  - 库存 (inventory)
  - 自定义模板

用法:
  python export_omp_generic.py --storage-state storage_state.json \\
      --start-time "2026-06-01 00:00:00" --end-time "2026-06-02 23:59:59" \\
      --template-type receiving --output receiving_20260602.xlsx

环境变量:
  OMP_TOKEN         Bearer token
  OMP_COOKIE         Cookie 字符串
  OMP_TRACK_KEY      手动指定 Track-Key（不填则自动生成）
"""
from __future__ import annotations

import argparse
import os
import sys
import time
from pathlib import Path
from typing import Any

# 确保项目根目录在 sys.path
PROJECT_DIR = Path(__file__).resolve().parent
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from omp_api import compact_json_text, create_omp_session, create_export_task, download_when_ready
from track_key import track_key_for_text

API_ORIGIN = "https://omp.xlwms.com"

# ── 预定义模板配置 ──────────────────────────────────────────
# 可以通过浏览器 F12 抓取其他 templateConfig 来扩展

# 出库单模板（与 export_delivery_once.py 一致）
TEMPLATE_DELIVERY = {
    "templateId": "2060429982465736704",
    "endpoint": "/gateway/omp/order/delivery/export",
    "templateConfig": {
        "groups": [
            {
                "groupKey": "BASIC", "groupName": "基础信息", "order": 0,
                "fields": [
                    {"fieldKey": 300003, "parentKey": "ORDER", "order": 0, "sheetIndex": 0, "sheetName": "出库单"},
                    {"fieldKey": 300008, "parentKey": "ORDER", "order": 1, "sheetIndex": 0, "sheetName": "出库单"},
                    {"fieldKey": 300021, "parentKey": "ORDER", "order": 2, "sheetIndex": 0, "sheetName": "出库单"},
                    {"fieldKey": 300020, "parentKey": "ORDER", "order": 3, "sheetIndex": 0, "sheetName": "出库单"},
                    {"fieldKey": 300022, "parentKey": "ORDER", "order": 4, "sheetIndex": 0, "sheetName": "出库单"},
                    {"fieldKey": 300023, "parentKey": "ORDER", "order": 5, "sheetIndex": 0, "sheetName": "出库单"},
                    {"fieldKey": 300024, "parentKey": "ORDER", "order": 6, "sheetIndex": 0, "sheetName": "出库单"},
                    {"fieldKey": 300019, "parentKey": "ORDER", "order": 7, "sheetIndex": 0, "sheetName": "出库单"},
                    {"fieldKey": 300026, "parentKey": "ORDER", "order": 8, "sheetIndex": 0, "sheetName": "出库单"},
                ],
            },
            {
                "groupKey": "PACKAGE", "groupName": "包裹信息", "order": 1,
                "fields": [
                    {"fieldKey": 300048, "parentKey": "PACKAGE", "order": 0, "sheetIndex": 0, "sheetName": "出库单"}
                ],
            },
            {
                "groupKey": "SKU", "groupName": "产品信息", "order": 2,
                "fields": [
                    {"fieldKey": 300061, "parentKey": "SKU", "order": 0, "sheetIndex": 0, "sheetName": "出库单"}
                ],
            },
            {
                "groupKey": "BASIC", "groupName": "基础信息", "order": 3,
                "fields": [
                    {"fieldKey": 310001, "parentKey": "ORDER", "order": 0, "sheetIndex": 1, "sheetName": "拣货库位明细"}
                ],
            },
            {
                "groupKey": "SKU", "groupName": "产品信息", "order": 4,
                "fields": [
                    {"fieldKey": 310006, "parentKey": "SKU", "order": 0, "sheetIndex": 1, "sheetName": "拣货库位明细"}
                ],
            },
        ]
    },
}

# 入库单模板（需要从浏览器 F12 抓取实际的 templateId 和 fieldKeys 后替换）
# 以下为框架模板，使用时需要替换实际的 templateId
TEMPLATE_RECEIVING = {
    "templateId": "REPLACE_ME_RECEIVING_TEMPLATE_ID",
    "endpoint": "/gateway/omp/order/receiving/export",
    "templateConfig": {
        "groups": [
            {
                "groupKey": "BASIC", "groupName": "基础信息", "order": 0,
                "fields": [
                    # 入库单号、ASN号、供应商、仓库、状态、创建时间、收货时间等
                    # 需要从浏览器开发者工具中抓取实际 fieldKey
                ],
            },
            {
                "groupKey": "SKU", "groupName": "产品信息", "order": 1,
                "fields": [],
            },
        ]
    },
}

# 库存模板
TEMPLATE_INVENTORY = {
    "templateId": "REPLACE_ME_INVENTORY_TEMPLATE_ID",
    "endpoint": "/gateway/omp/order/inventory/export",
    "templateConfig": {
        "groups": [
            {
                "groupKey": "BASIC", "groupName": "基础信息", "order": 0,
                "fields": [],
            },
        ]
    },
}


def get_template(template_type: str) -> dict[str, Any]:
    """获取预定义或自定义模板"""
    templates = {
        "delivery": TEMPLATE_DELIVERY,
        "receiving": TEMPLATE_RECEIVING,
        "inventory": TEMPLATE_INVENTORY,
    }
    if template_type in templates:
        return templates[template_type]
    raise ValueError(f"Unknown template type: {template_type}. Available: {list(templates.keys())}")


def build_payload(
    template: dict[str, Any],
    start_time: str,
    end_time: str,
    wh_code: str = "",
    status: str = "",
) -> dict[str, Any]:
    """构建导出请求 payload"""
    return {
        "appendixFlag": "",
        "categoryIdList": [],
        "countKind": "orderWeight",
        "endTime": end_time,
        "exportMode": 0,
        "expressFlag": "",
        "morePkgFlag": "",
        "noTypeNo": "",
        "orderSourceList": [],
        "platformCode": 1,
        "relatedReturnOrder": "",
        "startTime": start_time,
        "status": status,
        "storeName": "",
        "templateConfigJson": compact_json_text(template["templateConfig"]),
        "templateId": template["templateId"],
        "timeType": "createTime",
        "transitStatusList": [],
        "unitMark": 0,
        "whCode": wh_code,
        "withVas": "",
    }


def session_from_args(args: argparse.Namespace) -> Any:
    """从命令行参数创建 session"""
    if args.storage_state:
        return create_omp_session(args.storage_state, referer=f"{API_ORIGIN}/toolKit/taskCenter")

    token = os.environ.get("OMP_TOKEN", "").strip()
    if token.lower().startswith("bearer "):
        token = token[7:].strip()
    if not token:
        raise RuntimeError("Set OMP_TOKEN or pass --storage-state")
    return create_omp_session()


def main() -> None:
    parser = argparse.ArgumentParser(description="OMP 通用 Excel 导出工具")
    parser.add_argument("--template-type", default="delivery",
                        help="模板类型: delivery / receiving / inventory")
    parser.add_argument("--template-id", default="",
                        help="自定义 templateId（覆盖预定义模板）")
    parser.add_argument("--template-config", default="",
                        help="自定义 templateConfig JSON 文件路径（覆盖预定义模板）")
    parser.add_argument("--endpoint", default="",
                        help="自定义导出接口路径（覆盖预定义模板）")
    parser.add_argument("--start-time", default="2026-06-01 00:00:00")
    parser.add_argument("--end-time", default="2026-06-02 23:59:59")
    parser.add_argument("--wh-code", default="")
    parser.add_argument("--status", default="")
    parser.add_argument("--output", default="")
    parser.add_argument("--storage-state", default="")
    parser.add_argument("--attempts", type=int, default=30)
    parser.add_argument("--interval", type=int, default=10)
    args = parser.parse_args()

    # 加载模板
    template = get_template(args.template_type)

    # 覆盖自定义参数
    if args.template_id:
        template["templateId"] = args.template_id
    if args.endpoint:
        template["endpoint"] = args.endpoint
    if args.template_config:
        import json
        config_path = Path(args.template_config)
        if config_path.exists():
            template["templateConfig"] = json.loads(config_path.read_text(encoding="utf-8"))
        else:
            print(f"WARNING: template-config file not found: {config_path}")

    # 检查模板是否有效
    if "REPLACE_ME" in template["templateId"]:
        print("ERROR: 入库单/库存模板尚未配置实际的 templateId。")
        print("请通过浏览器 F12 抓取对应的 templateId 和 fieldKeys，")
        print("然后使用 --template-id 和 --template-config 参数指定。")
        print(f"\n当前模板类型: {args.template_type}")
        print(f"导出接口: {template['endpoint']}")
        sys.exit(1)

    session = session_from_args(args)
    payload = build_payload(template, args.start_time, args.end_time, args.wh_code, args.status)
    task_id = create_export_task(session, payload)
    print(f"Created export task: {task_id}", flush=True)

    output_path = Path(args.output or f"omp_export_{args.template_type}_{task_id}.xlsx").resolve()
    downloaded = download_when_ready(session, task_id, output_path, args.attempts, args.interval)
    print(f"Downloaded: {downloaded}", flush=True)


if __name__ == "__main__":
    main()
