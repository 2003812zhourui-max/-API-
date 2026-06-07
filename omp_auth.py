"""OMP 自动登录模块

逆向自 OMP 前端 app.js 的 loginDual 函数。
登录一次获取 omp-token + wms-token，写入 storage_state.json。

用法:
  python omp_auth.py --account RPA001 --password yourpassword
  python omp_auth.py --account RPA001 --password yourpassword --output new_state.json
"""
from __future__ import annotations

import argparse
import json
import os
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

import requests

from track_key import track_key_for_text, compact_json_text

API_ORIGIN = "https://omp.xlwms.com"

# 固定的 deviceFingerprint（从 localStorage 提取，不变）
DEVICE_FINGERPRINT = "5db79f09"

# deviceInfo 模板
DEVICE_INFO = json.dumps({
    "userAgent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36",
    "platform": "Win32",
    "language": "zh-CN",
    "screenResolution": "1920x1080",
    "timezone": "Asia/Shanghai",
}, separators=(",", ":"))


def login_omp(login_account: str, password: str) -> dict[str, Any]:
    """登录 OMP，返回 {omp_token, wms_token, user_info}"""
    login_flow_id = str(uuid.uuid4())

    # 登录 OMP
    omp_body = {
        "loginAccount": login_account,
        "password": password,
        "loginFlowId": login_flow_id,
        "deviceFingerprint": DEVICE_FINGERPRINT,
        "deviceInfo": DEVICE_INFO,
        "businessType": "omp",
    }
    omp_result = _call_login("/gateway/omp/auth/login", omp_body, login_flow_id)

    # 登录 WMS
    wms_body = {
        "loginAccount": login_account,
        "password": password,
        "loginFlowId": login_flow_id,
        "deviceFingerprint": DEVICE_FINGERPRINT,
        "deviceInfo": DEVICE_INFO,
        "businessType": "wms",
    }
    wms_result = _call_login("/gateway/wms/auth/login", wms_body, login_flow_id)

    omp_token = (omp_result.get("data") or {}).get("token") or ""
    wms_token = (wms_result.get("data") or {}).get("token") or ""

    if not omp_token or not wms_token:
        # 检查是否需要 MFA 验证
        omp_data = omp_result.get("data") or {}
        wms_data = wms_result.get("data") or {}
        login_action = omp_data.get("loginAction") or wms_data.get("loginAction") or ""
        if login_action and login_action != "LOGIN_SUCCESS":
            return {
                "success": False,
                "step": login_action,
                "challengeId": omp_data.get("challengeId") or wms_data.get("challengeId"),
                "mfaChannel": omp_data.get("mfaChannel") or wms_data.get("mfaChannel"),
                "mfaMaskedTarget": omp_data.get("mfaMaskedTarget") or wms_data.get("mfaMaskedTarget"),
                "bindToken": omp_data.get("bindToken") or wms_data.get("bindToken"),
                "securitySessionToken": omp_data.get("securitySessionToken") or wms_data.get("securitySessionToken"),
                "needUpdatePassword": omp_data.get("needUpdatePassword") or wms_data.get("needUpdatePassword"),
                "msg": f"Login requires additional step: {login_action}",
            }

        error_msg = omp_result.get("msg") or wms_result.get("msg") or "Unknown error"
        return {"success": False, "step": "ERROR", "msg": error_msg}

    return {
        "success": True,
        "step": "LOGIN_SUCCESS",
        "omp_token": omp_token,
        "wms_token": wms_token,
        "omp_data": omp_result.get("data") or {},
        "wms_data": wms_result.get("data") or {},
    }


def _call_login(endpoint: str, body: dict[str, Any], login_flow_id: str) -> dict[str, Any]:
    """调用登录接口"""
    body_str = compact_json_text(body)
    track_key = track_key_for_text(body_str)

    headers = {
        "Content-Type": "application/json;charset=UTF-8",
        "Accept": "application/json, text/plain, */*",
        "Origin": API_ORIGIN,
        "Referer": f"{API_ORIGIN}/login",
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36"
        ),
        "X-Login-Flow-Id": login_flow_id,
        "X-Device-Fingerprint": DEVICE_FINGERPRINT,
        "X-Client-Type": "web",
        "Track-Key": track_key,
        "lang": "zh",
        "version": "prod",
    }

    response = requests.post(
        f"{API_ORIGIN}{endpoint}",
        data=body_str.encode("utf-8"),
        headers=headers,
        timeout=30,
    )
    response.raise_for_status()
    return response.json()


def update_storage_state(storage_path: str | Path, omp_token: str, wms_token: str) -> None:
    """更新 storage_state.json 中的 token"""
    path = Path(storage_path)
    state = json.loads(path.read_text(encoding="utf-8"))

    for origin_entry in state.get("origins", []):
        if origin_entry.get("origin") != API_ORIGIN:
            continue
        for entry in origin_entry.get("localStorage", []):
            name = entry.get("name", "")
            if name == "omp-token":
                entry["value"] = omp_token
            elif name == "wms-token":
                entry["value"] = wms_token

    path.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Updated tokens in: {path}")


def create_fresh_storage_state(output_path: str | Path, omp_token: str, wms_token: str) -> None:
    """用 token 创建全新的 storage_state.json"""
    state = {
        "cookies": [
            {
                "name": "version",
                "value": "prod",
                "domain": ".xlwms.com",
                "path": "/",
                "secure": False,
                "httpOnly": False,
                "sameSite": "Lax",
            },
            {
                "name": "prod",
                "value": "always",
                "domain": ".xlwms.com",
                "path": "/",
                "secure": False,
                "httpOnly": False,
                "sameSite": "Lax",
            },
        ],
        "origins": [
            {
                "origin": API_ORIGIN,
                "localStorage": [
                    {"name": "omp-token", "value": omp_token},
                    {"name": "wms-token", "value": wms_token},
                    {"name": "lx_device_fingerprint", "value": DEVICE_FINGERPRINT},
                    {"name": "language", "value": "zh"},
                    {"name": "version", "value": "prod"},
                ],
            }
        ],
    }
    Path(output_path).write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Created fresh storage_state at: {output_path}")


def main() -> None:
    parser = argparse.ArgumentParser(description="OMP 自动登录，获取 token")
    parser.add_argument("--account", required=True, help="登录账号")
    parser.add_argument("--password", required=True, help="登录密码")
    parser.add_argument("--output", default="", help="输出 storage_state.json 路径（默认更新项目中的 storage_state.json）")
    args = parser.parse_args()

    print(f"Logging in as: {args.account}")
    result = login_omp(args.account, args.password)

    if result["success"]:
        print(f"\n[SUCCESS] Login success!")
        print(f"OMP token: {result['omp_token'][:50]}...")
        print(f"WMS token: {result['wms_token'][:50]}...")

        # 更新 storage_state
        project_dir = Path(__file__).resolve().parent
        default_storage = project_dir / "storage_state.json"

        if args.output:
            output_path = Path(args.output)
        elif default_storage.exists():
            output_path = default_storage
            update_storage_state(output_path, result["omp_token"], result["wms_token"])
        else:
            output_path = project_dir / "storage_state.json"
            create_fresh_storage_state(output_path, result["omp_token"], result["wms_token"])
    else:
        step = result.get("step", "ERROR")
        msg = result.get("msg", "")
        print(f"\n[FAILED] Login failed: step={step}")
        if msg:
            print(f"   message: {msg}")
        if step in ("SEND_VERIFY_CODE", "BIND_CONTACT"):
            print("   MFA verification required — this account needs two-factor auth.")
            print(f"   Challenge ID: {result.get('challengeId', 'N/A')}")
            print(f"   MFA Channel: {result.get('mfaChannel', 'N/A')}")
            print(f"   MFA Target: {result.get('mfaMaskedTarget', 'N/A')}")


if __name__ == "__main__":
    main()
