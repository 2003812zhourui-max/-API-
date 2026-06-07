"""共享 OMP API 模块 —— 认证、导出、下载、拣货人查询"""
from __future__ import annotations

import json
import os
import time as time_module
from collections import Counter
from pathlib import Path
from typing import Any
from urllib.parse import urlencode

import requests

from track_key import track_key_for_text

API_ORIGIN = "https://omp.xlwms.com"


def compact_json_text(value: Any) -> str:
    # NOTE: ensure_ascii=True is REQUIRED to match browser JSON.stringify behavior
    # for Track-Key signing. The browser's signer hashes the \uXXXX escaped form.
    return json.dumps(value, ensure_ascii=True, separators=(",", ":"))


def load_storage_state(path: str) -> dict[str, Any]:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def local_storage_from_state(state: dict[str, Any], origin: str = API_ORIGIN) -> dict[str, str]:
    result: dict[str, str] = {}
    for origin_entry in state.get("origins", []):
        if origin_entry.get("origin") != origin:
            continue
        for entry in origin_entry.get("localStorage", []):
            result[str(entry.get("name"))] = str(entry.get("value"))
    return result


def create_omp_session(
    storage_state: str = "",
    *,
    referer: str = "",
) -> requests.Session:
    """创建带认证的 OMP requests.Session"""
    session = requests.Session()

    if storage_state:
        state = load_storage_state(storage_state)
        for cookie in state.get("cookies", []):
            session.cookies.set(
                cookie["name"],
                cookie["value"],
                domain=cookie.get("domain"),
                path=cookie.get("path", "/"),
            )
        local_storage = local_storage_from_state(state)
        token = local_storage.get("omp-token") or local_storage.get("wms-token") or ""
    else:
        token = ""

    env_token = os.environ.get("OMP_TOKEN", "").strip()
    token = env_token or token
    if token.lower().startswith("bearer "):
        token = token[7:].strip()
    if not token:
        raise RuntimeError("Set OMP_TOKEN or pass --storage-state before fetching picker names.")

    cookie = os.environ.get("OMP_COOKIE", "").strip()
    if cookie:
        session.headers["Cookie"] = cookie

    if not referer:
        referer = f"{API_ORIGIN}/globalOrder/parcel"

    session.headers.update({
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        "Authorization": f"Bearer {token}",
        "Origin": API_ORIGIN,
        "Referer": referer,
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36"
        ),
        "lang": "zh",
        "version": "prod",
    })
    return session


def delivery_page_payload(
    wave_no: str,
    start_time: str,
    end_time: str,
    size: int,
    current: int = 1,
) -> dict[str, Any]:
    return {
        "appendixFlag": "",
        "categoryIdList": [],
        "countKind": "orderWeight",
        "current": current,
        "endTime": end_time,
        "expressFlag": "",
        "morePkgFlag": "",
        "noTypeNo": wave_no,
        "orderSourceList": [],
        "platformCode": 1,
        "relatedReturnOrder": "",
        "size": size,
        "startTime": start_time,
        "status": "",
        "storeName": "",
        "timeType": "createTime",
        "transitStatusList": [],
        "unitMark": 0,
        "waveNo": wave_no,
        "waveNoList": [wave_no],
        "whCode": "",
        "withVas": "",
    }


def fetch_delivery_records_by_wave(
    session: requests.Session,
    wave_no: str,
    start_time: str,
    end_time: str,
    page_size: int,
) -> list[dict[str, Any]]:
    """根据波次号查询出库单列表"""
    payload = delivery_page_payload(wave_no, start_time, end_time, page_size)
    body = compact_json_text(payload)
    try:
        track_key = track_key_for_text(body)
    except Exception:
        track_key = os.environ.get("OMP_TRACK_KEY", "").strip()
    if not track_key:
        raise RuntimeError("Could not generate Track-Key for delivery page request.")

    response = session.post(
        f"{API_ORIGIN}/gateway/omp/order/delivery/page",
        data=body.encode("utf-8"),
        headers={
            "Content-Type": "application/json;charset=UTF-8",
            "Track-Key": track_key,
            "Referer": f"{API_ORIGIN}/globalOrder/parcel",
        },
        timeout=60,
    )
    response.raise_for_status()
    result = response.json()
    if result.get("code") != 200:
        raise RuntimeError(f"delivery page failed for {wave_no}: {result}")
    records = (result.get("data") or {}).get("records") or []
    return records if isinstance(records, list) else []


def fetch_pick_name_from_order(session: requests.Session, record: dict[str, Any]) -> str:
    """从订单日志中提取拣货人姓名"""
    delivery_no = str(record.get("deliveryNo") or "").strip()
    customer_code = str(record.get("customerCode") or "").strip()
    wh_code = str(record.get("whCode") or "").strip()
    source_no = str(record.get("sourceNo") or "").strip()
    if not delivery_no or not customer_code or not wh_code:
        return ""

    query = urlencode({
        "billType": "delivery_packet",
        "billNo": delivery_no,
        "customerCode": customer_code,
        "whCode": wh_code,
    })
    response = session.get(
        f"{API_ORIGIN}/gateway/omp/order/log/progress?{query}",
        headers={
            "Referer": f"{API_ORIGIN}/globalOrder/parcelDetail/{customer_code}/{wh_code}/{delivery_no}/{source_no}"
        },
        timeout=60,
    )
    response.raise_for_status()
    result = response.json()
    if result.get("code") != 200:
        return ""

    events = result.get("data") or []
    for event in events:
        if event.get("bizType") == "delivery_pick" or event.get("bizDesc") == "拣货":
            return str(event.get("operateName") or event.get("operateAccount") or "").strip()
    return ""


def resolve_picker_names(
    wave_nos: list[str],
    storage_state: str,
    start_time: str,
    end_time: str,
    page_size: int = 20,
    sample_limit: int = 3,
    wave_limit: int = 0,
    sleep_seconds: float = 0.0,
) -> dict[str, str]:
    """批量解析波次号对应的拣货人姓名"""
    session = create_omp_session(storage_state)
    result: dict[str, str] = {}
    target_wave_nos = wave_nos[:wave_limit] if wave_limit > 0 else wave_nos

    for index, wave_no in enumerate(target_wave_nos, 1):
        if not wave_no:
            continue
        try:
            records = fetch_delivery_records_by_wave(session, wave_no, start_time, end_time, page_size)
            names: list[str] = []
            for record in records[:sample_limit]:
                name = fetch_pick_name_from_order(session, record)
                if name:
                    names.append(name)
            if names:
                result[wave_no] = Counter(names).most_common(1)[0][0]
            print(f"picker {index}/{len(target_wave_nos)} {wave_no}: {result.get(wave_no, '') or '-'}", flush=True)
        except Exception as exc:
            result[wave_no] = ""
            print(f"picker {index}/{len(target_wave_nos)} {wave_no}: failed: {exc}", flush=True)

        if sleep_seconds > 0:
            time_module.sleep(sleep_seconds)

    return result


def create_export_task(
    session: requests.Session,
    payload: dict[str, Any],
) -> str:
    """创建 OMP 导出任务，返回 taskId"""
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


def download_when_ready(
    session: requests.Session,
    task_id: str,
    output_path: Path,
    attempts: int = 30,
    interval: int = 10,
) -> Path:
    """轮询下载导出任务结果"""
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
        time_module.sleep(interval)
    raise TimeoutError(f"task {task_id} was not downloadable after {attempts} attempts")
