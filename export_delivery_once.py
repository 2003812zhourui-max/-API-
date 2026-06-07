from __future__ import annotations

import argparse
import os
import sys
import time
from pathlib import Path
from typing import Any

import requests


PROJECT_DIR = Path(__file__).resolve().parent
API_ORIGIN = "https://omp.xlwms.com"

from track_key import compact_json_text, track_key_for_text  # noqa: E402
from silent_auth import create_session_from_storage_state, load_storage_state, local_storage_from_state  # noqa: E402


def build_payload(start_time: str, end_time: str, wh_code: str, status: str) -> dict[str, Any]:
    template_config = {
        "groups": [
            {
                "groupKey": "BASIC",
                "groupName": "基础信息",
                "order": 0,
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
                "groupKey": "PACKAGE",
                "groupName": "包裹信息",
                "order": 1,
                "fields": [
                    {"fieldKey": 300048, "parentKey": "PACKAGE", "order": 0, "sheetIndex": 0, "sheetName": "出库单"}
                ],
            },
            {
                "groupKey": "SKU",
                "groupName": "产品信息",
                "order": 2,
                "fields": [
                    {"fieldKey": 300061, "parentKey": "SKU", "order": 0, "sheetIndex": 0, "sheetName": "出库单"}
                ],
            },
            {
                "groupKey": "BASIC",
                "groupName": "基础信息",
                "order": 3,
                "fields": [
                    {"fieldKey": 310001, "parentKey": "ORDER", "order": 0, "sheetIndex": 1, "sheetName": "拣货库位明细"}
                ],
            },
            {
                "groupKey": "SKU",
                "groupName": "产品信息",
                "order": 4,
                "fields": [
                    {"fieldKey": 310006, "parentKey": "SKU", "order": 0, "sheetIndex": 1, "sheetName": "拣货库位明细"}
                ],
            },
        ]
    }

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
        "templateConfigJson": compact_json_text(template_config),
        "templateId": "2060429982465736704",
        "timeType": "createTime",
        "transitStatusList": [],
        "unitMark": 0,
        "whCode": wh_code,
        "withVas": "",
    }


def create_session(token: str) -> requests.Session:
    cookie = os.environ.get("OMP_COOKIE", "").strip() or "version=prod; prod=always"
    session = requests.Session()
    session.headers.update(
        {
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
            "Authorization": f"Bearer {token}",
            "Cookie": cookie,
            "Origin": API_ORIGIN,
            "Referer": f"{API_ORIGIN}/globalOrder/parcel",
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36"
            ),
            "lang": "zh",
            "version": "prod",
        }
    )
    return session


def session_from_args(args: argparse.Namespace) -> requests.Session:
    if args.storage_state:
        session, auth_values = create_session_from_storage_state(
            args.storage_state,
            wh_code=args.wh_code,
            auth_scheme="bearer",
            origin=API_ORIGIN,
            referer=f"{API_ORIGIN}/globalOrder/parcel",
        )
        if not auth_values:
            raise RuntimeError(f"No token found in storage_state: {args.storage_state}")
        local_storage = local_storage_from_state(load_storage_state(args.storage_state), origin=API_ORIGIN)
        omp_token = local_storage.get("omp-token", "").strip()
        if omp_token:
            session.headers["authorization"] = f"Bearer {omp_token}"
        return session

    token = os.environ.get("OMP_TOKEN", "").strip()
    if token.lower().startswith("bearer "):
        token = token[7:].strip()
    if not token:
        raise RuntimeError("Set OMP_TOKEN or pass --storage-state before running this script.")
    return create_session(token)


def create_export_task(session: requests.Session, payload: dict[str, Any]) -> str:
    body = compact_json_text(payload)
    track_key = os.environ.get("OMP_TRACK_KEY", "").strip() or track_key_for_text(body)
    response = session.post(
        f"{API_ORIGIN}/gateway/omp/order/delivery/export",
        data=body.encode("utf-8"),
        headers={
            "Content-Type": "application/json;charset=UTF-8",
            "Track-Key": track_key,
        },
        timeout=60,
    )
    if response.status_code >= 400:
        raise RuntimeError(f"create export HTTP {response.status_code}: {response.text[:500]}")
    result = response.json()
    if result.get("code") != 200 or not result.get("data"):
        raise RuntimeError(f"create export failed: {result}")
    return str(result["data"])


def download_when_ready(session: requests.Session, task_id: str, output_path: Path, attempts: int, interval: int) -> Path:
    url = f"{API_ORIGIN}/gateway/omp/system/taskCenter/task/export/download?taskId={task_id}"
    for attempt in range(1, attempts + 1):
        response = session.get(
            url,
            headers={"referer": f"{API_ORIGIN}/toolKit/taskCenter"},
            timeout=120,
        )
        response.raise_for_status()
        content_type = response.headers.get("content-type", "")
        content = response.content
        if content.startswith(b"PK") or "spreadsheetml.sheet" in content_type or "octet-stream" in content_type:
            output_path.write_bytes(content)
            return output_path

        if "application/json" in content_type:
            result = response.json()
            download_url = (result.get("data") or {}).get("downLoadUrl") or (result.get("data") or {}).get("downloadUrl")
            if download_url:
                file_response = session.get(download_url, timeout=120)
                file_response.raise_for_status()
                output_path.write_bytes(file_response.content)
                return output_path

        preview = response.text[:200].replace("\n", " ")
        print(f"download not ready ({attempt}/{attempts}): {preview}", flush=True)
        time.sleep(interval)
    raise TimeoutError(f"task {task_id} was not downloadable after {attempts} attempts")


def main() -> None:
    parser = argparse.ArgumentParser(description="Create and download one OMP delivery export.")
    parser.add_argument("--start-time", default="2026-06-01 00:00:00")
    parser.add_argument("--end-time", default="2026-06-02 23:59:59")
    parser.add_argument("--wh-code", default="")
    parser.add_argument("--status", default="")
    parser.add_argument("--output", default="")
    parser.add_argument("--storage-state", default="")
    parser.add_argument("--attempts", type=int, default=20)
    parser.add_argument("--interval", type=int, default=10)
    args = parser.parse_args()

    session = session_from_args(args)
    task_id = create_export_task(session, build_payload(args.start_time, args.end_time, args.wh_code, args.status))
    print(f"created task: {task_id}", flush=True)

    output_path = Path(args.output or f"delivery_export_{task_id}.xlsx").resolve()
    downloaded = download_when_ready(session, task_id, output_path, args.attempts, args.interval)
    print(f"downloaded: {downloaded}", flush=True)


if __name__ == "__main__":
    main()
