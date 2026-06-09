from __future__ import annotations

import argparse
import base64
import json
import os
import time
from pathlib import Path
from typing import Any

import requests

from silent_auth import create_session_from_storage_state, load_storage_state, local_storage_from_state
from track_key import compact_json_text, track_key_for_text


API_ORIGIN = "https://omp.xlwms.com"


def jwt_expiry(token: str) -> int | None:
    token = token.strip()
    if token.lower().startswith("bearer "):
        token = token[7:].strip()
    parts = token.split(".")
    if len(parts) < 2:
        return None
    try:
        payload = parts[1] + "=" * (-len(parts[1]) % 4)
        data = json.loads(base64.urlsafe_b64decode(payload.encode("utf-8")))
    except (ValueError, json.JSONDecodeError):
        return None
    exp = data.get("exp")
    return exp if isinstance(exp, int) else None


def assert_session_token_fresh(session: requests.Session) -> None:
    auth_value = session.headers.get("Authorization") or session.headers.get("authorization") or ""
    exp = jwt_expiry(auth_value)
    if exp and exp <= int(time.time()):
        expired_at = time.strftime("%Y-%m-%d %H:%M:%S UTC", time.gmtime(exp))
        raise RuntimeError(
            "OMP token expired at "
            f"{expired_at}. Refresh wms_storage_state.json by logging in again, "
            "or set OMP_TOKEN/OMP_COOKIE from a fresh F12 request."
        )


def normalize_status(status: str) -> str | int:
    status = str(status).strip()
    if not status:
        return ""
    try:
        return int(status)
    except ValueError:
        return status


def wh_code_from_storage_state(storage_state: str) -> str:
    local_storage = local_storage_from_state(load_storage_state(storage_state), origin=API_ORIGIN)
    try:
        wh_info = json.loads(local_storage.get("wh", "{}"))
    except json.JSONDecodeError:
        return ""
    return str(wh_info.get("whCode") or "")


def export_template_config() -> dict[str, Any]:
    basic_name = "\u57fa\u7840\u4fe1\u606f"
    package_name = "\u5305\u88f9\u4fe1\u606f"
    sku_name = "\u4ea7\u54c1\u4fe1\u606f"
    outbound_sheet = "\u51fa\u5e93\u5355"
    picking_detail_sheet = "\u62e3\u8d27\u5e93\u4f4d\u660e\u7ec6"

    return {
        "groups": [
            {
                "groupKey": "BASIC",
                "groupName": basic_name,
                "order": 0,
                "fields": [
                    {"fieldKey": 300003, "parentKey": "ORDER", "order": 0, "sheetIndex": 0, "sheetName": outbound_sheet},
                    {"fieldKey": 300008, "parentKey": "ORDER", "order": 1, "sheetIndex": 0, "sheetName": outbound_sheet},
                    {"fieldKey": 300021, "parentKey": "ORDER", "order": 2, "sheetIndex": 0, "sheetName": outbound_sheet},
                    {"fieldKey": 300020, "parentKey": "ORDER", "order": 3, "sheetIndex": 0, "sheetName": outbound_sheet},
                    {"fieldKey": 300022, "parentKey": "ORDER", "order": 4, "sheetIndex": 0, "sheetName": outbound_sheet},
                    {"fieldKey": 300023, "parentKey": "ORDER", "order": 5, "sheetIndex": 0, "sheetName": outbound_sheet},
                    {"fieldKey": 300024, "parentKey": "ORDER", "order": 6, "sheetIndex": 0, "sheetName": outbound_sheet},
                    {"fieldKey": 300019, "parentKey": "ORDER", "order": 7, "sheetIndex": 0, "sheetName": outbound_sheet},
                    {"fieldKey": 300026, "parentKey": "ORDER", "order": 8, "sheetIndex": 0, "sheetName": outbound_sheet},
                ],
            },
            {
                "groupKey": "PACKAGE",
                "groupName": package_name,
                "order": 1,
                "fields": [
                    {"fieldKey": 300048, "parentKey": "PACKAGE", "order": 0, "sheetIndex": 0, "sheetName": outbound_sheet}
                ],
            },
            {
                "groupKey": "SKU",
                "groupName": sku_name,
                "order": 2,
                "fields": [
                    {"fieldKey": 300061, "parentKey": "SKU", "order": 0, "sheetIndex": 0, "sheetName": outbound_sheet}
                ],
            },
            {
                "groupKey": "BASIC",
                "groupName": basic_name,
                "order": 3,
                "fields": [
                    {
                        "fieldKey": 310001,
                        "parentKey": "ORDER",
                        "order": 0,
                        "sheetIndex": 1,
                        "sheetName": picking_detail_sheet,
                    }
                ],
            },
            {
                "groupKey": "SKU",
                "groupName": sku_name,
                "order": 4,
                "fields": [
                    {
                        "fieldKey": 310006,
                        "parentKey": "SKU",
                        "order": 0,
                        "sheetIndex": 1,
                        "sheetName": picking_detail_sheet,
                    }
                ],
            },
        ]
    }


def build_payload(start_time: str, end_time: str, wh_code: str, status: str) -> dict[str, Any]:
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
        "status": normalize_status(status),
        "storeName": "",
        "templateConfigJson": compact_json_text(export_template_config()),
        "templateId": "2060429982465736704",
        "timeType": "createTime",
        "transitStatusList": [],
        "unitMark": 0,
        "whCode": wh_code,
        "withVas": "",
    }


def create_basic_session(token: str) -> requests.Session:
    if token.lower().startswith("bearer "):
        token = token[7:].strip()
    session = requests.Session()
    cookie = os.environ.get("OMP_COOKIE", "").strip()
    if cookie:
        session.headers["Cookie"] = cookie
    session.headers.update(
        {
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
            "Authorization": f"Bearer {token}",
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
            wh_code=getattr(args, "wh_code", ""),
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
    else:
        token = os.environ.get("OMP_TOKEN", "").strip()
        if not token:
            raise RuntimeError("Set OMP_TOKEN or pass --storage-state before running this script.")
        session = create_basic_session(token)

    env_token = os.environ.get("OMP_TOKEN", "").strip()
    if env_token:
        if env_token.lower().startswith("bearer "):
            env_token = env_token[7:].strip()
        session.headers["Authorization"] = f"Bearer {env_token}"

    env_cookie = os.environ.get("OMP_COOKIE", "").strip()
    if env_cookie:
        session.headers["Cookie"] = env_cookie

    assert_session_token_fresh(session)
    return session


def create_export_task(session: requests.Session, payload: dict[str, Any]) -> str:
    body = compact_json_text(payload)
    try:
        track_key = track_key_for_text(body)
    except Exception:
        track_key = os.environ.get("OMP_TRACK_KEY", "").strip()
    if not track_key:
        raise RuntimeError("Could not generate Track-Key.")

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
    output_path.parent.mkdir(parents=True, exist_ok=True)

    for attempt in range(1, attempts + 1):
        response = session.get(
            url,
            headers={"Referer": f"{API_ORIGIN}/toolKit/taskCenter"},
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
            data = result.get("data") or {}
            download_url = data.get("downLoadUrl") or data.get("downloadUrl") if isinstance(data, dict) else ""
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
    parser.add_argument("--wh-code", default="", help="Warehouse code, auto-detected from storage_state if empty")
    parser.add_argument("--status", default="100", help="Order status filter (100=all non-draft)")
    parser.add_argument("--output", default="")
    parser.add_argument("--storage-state", default="")
    parser.add_argument("--attempts", type=int, default=20)
    parser.add_argument("--interval", type=int, default=10)
    args = parser.parse_args()

    wh_code = args.wh_code or (wh_code_from_storage_state(args.storage_state) if args.storage_state else "")

    session = session_from_args(args)
    task_id = create_export_task(session, build_payload(args.start_time, args.end_time, wh_code, args.status))
    print(f"created task: {task_id}", flush=True)

    output_path = Path(args.output or f"delivery_export_{task_id}.xlsx").resolve()
    downloaded = download_when_ready(session, task_id, output_path, args.attempts, args.interval)
    print(f"downloaded: {downloaded}", flush=True)


if __name__ == "__main__":
    main()
