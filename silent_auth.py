from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import requests


API_ORIGIN = "https://omp.xlwms.com"


def load_storage_state(path: str | Path) -> dict[str, Any]:
    storage_path = Path(path)
    if not storage_path.exists():
        raise FileNotFoundError(f"storage_state file does not exist: {storage_path}")
    data = json.loads(storage_path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise RuntimeError("storage_state JSON root must be an object.")
    return data


def local_storage_from_state(state: dict[str, Any], origin: str = API_ORIGIN) -> dict[str, str]:
    result: dict[str, str] = {}
    for origin_entry in state.get("origins", []):
        if not isinstance(origin_entry, dict) or origin_entry.get("origin") != origin:
            continue
        for entry in origin_entry.get("localStorage", []):
            if not isinstance(entry, dict):
                continue
            name = entry.get("name")
            value = entry.get("value")
            if isinstance(name, str) and isinstance(value, str):
                result[name] = value
    return result


def parse_json_text(value: str) -> dict[str, Any]:
    try:
        parsed = json.loads(value)
    except (TypeError, json.JSONDecodeError):
        return {}
    return parsed if isinstance(parsed, dict) else {}


def auth_header_candidates(token: str, auth_scheme: str = "auto") -> list[str]:
    token = token.strip()
    if not token:
        return []
    if token.lower().startswith("bearer "):
        return [token]
    if auth_scheme == "raw":
        return [token]
    if auth_scheme == "bearer":
        return [f"Bearer {token}"]
    return [f"Bearer {token}", token]


def create_session_from_storage_state(
    storage_state_path: str | Path,
    *,
    wh_code: str = "",
    auth_scheme: str = "auto",
    origin: str = API_ORIGIN,
    referer: str = API_ORIGIN,
) -> tuple[requests.Session, list[str]]:
    """Create a requests session from a Playwright storage_state file."""
    state = load_storage_state(storage_state_path)
    session = requests.Session()

    for cookie in state.get("cookies", []):
        if not isinstance(cookie, dict):
            continue
        name = cookie.get("name")
        value = cookie.get("value")
        if not isinstance(name, str) or not isinstance(value, str):
            continue
        session.cookies.set(
            name,
            value,
            domain=cookie.get("domain") if isinstance(cookie.get("domain"), str) else None,
            path=cookie.get("path") if isinstance(cookie.get("path"), str) else "/",
        )

    local_storage = local_storage_from_state(state, origin=origin)
    wh_info = parse_json_text(local_storage.get("wh", ""))
    token = local_storage.get("omp-token") or local_storage.get("wms-token") or ""
    auth_values = auth_header_candidates(token, auth_scheme=auth_scheme)

    resolved_wh_code = wh_code or str(wh_info.get("whCode") or "")
    tenant_code = str(wh_info.get("tenantCode") or "")
    language = local_storage.get("language", "zh")
    version = session.cookies.get("version") or local_storage.get("version", "")

    session.headers.update(
        {
            "accept": "application/json, text/plain, */*",
            "origin": origin,
            "referer": referer,
            "user-agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0 Safari/537.36"
            ),
            "x-requested-with": "XMLHttpRequest",
        }
    )
    if resolved_wh_code:
        session.headers["whcode"] = resolved_wh_code
    if tenant_code:
        session.headers["tenantcode"] = tenant_code
    if language:
        session.headers["lang"] = language
        session.headers["language"] = language
    if version:
        session.headers["version"] = version
    if auth_values:
        session.headers["authorization"] = auth_values[0]

    return session, auth_values
