from __future__ import annotations

import argparse
import json
import re
import subprocess
from typing import Any

import requests


API_ORIGIN = "https://omp.xlwms.com"
TARGET_PAGE = "https://omp.xlwms.com/wms/outbound/parcel"
DEFAULT_MARKER = 'function g(e){const t=0,n="",i=8;'

_TRACK_SIGN_JS_CACHE: str | None = None


def compact_json_text(value: Any) -> str:
    """Serialize JSON exactly like the frontend request body."""
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def absolute_api_url(url: str, api_origin: str = API_ORIGIN) -> str:
    if url.startswith(("http://", "https://")):
        return url
    if not url.startswith("/"):
        url = f"/{url}"
    return f"{api_origin}{url}"


def extract_js_function(source: str, marker: str = DEFAULT_MARKER) -> str:
    """Extract the minified Track-Key signer function from app.js."""
    start = source.find(marker)
    if start < 0:
        raise RuntimeError("Track-Key signer marker was not found in app.js.")

    depth = 0
    in_string = ""
    escape = False
    for index in range(start, len(source)):
        char = source[index]

        if in_string:
            if escape:
                escape = False
            elif char == "\\":
                escape = True
            elif char == in_string:
                in_string = ""
            continue

        if char in {'"', "'", "`"}:
            in_string = char
        elif char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return source[start : index + 1]

    raise RuntimeError("Track-Key signer function extraction failed.")


def load_track_sign_js(
    target_page: str = TARGET_PAGE,
    api_origin: str = API_ORIGIN,
    marker: str = DEFAULT_MARKER,
    *,
    force_reload: bool = False,
) -> str:
    """Download the current frontend app.js and extract the Track-Key signer."""
    global _TRACK_SIGN_JS_CACHE
    if _TRACK_SIGN_JS_CACHE and not force_reload:
        return _TRACK_SIGN_JS_CACHE

    headers = {"user-agent": "Mozilla/5.0"}
    html = requests.get(target_page, timeout=30, headers=headers).text
    matches = re.findall(r'src="([^"]*/js/app\.[^"]+\.js)"', html)
    if not matches:
        raise RuntimeError("Could not find app.js from target page.")

    app_js_url = absolute_api_url(matches[-1], api_origin=api_origin)
    app_js = requests.get(app_js_url, timeout=30, headers=headers).text
    _TRACK_SIGN_JS_CACHE = extract_js_function(app_js, marker=marker)
    return _TRACK_SIGN_JS_CACHE


def track_key_for_text(text: str) -> str:
    """Generate Track-Key for a compact JSON request body."""
    sign_function = load_track_sign_js()
    script = (
        "const fs=require('fs');"
        "const input=fs.readFileSync(0,'utf8');"
        f"const sign={sign_function};"
        "process.stdout.write(sign(input));"
    )
    try:
        result = subprocess.run(
            ["node", "-e", script],
            input=text,
            text=True,
            capture_output=True,
            check=True,
            timeout=10,
        )
    except FileNotFoundError as exc:
        raise RuntimeError("Node.js is required to generate Track-Key.") from exc
    except subprocess.SubprocessError as exc:
        raise RuntimeError(f"Track-Key generation failed: {exc}") from exc
    return result.stdout.strip()


def track_key_for_json(payload: Any) -> str:
    return track_key_for_text(compact_json_text(payload))


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate Track-Key for xlwms/OMP JSON bodies.")
    parser.add_argument("--json", default="", help="JSON string. If omitted, stdin is used.")
    parser.add_argument("--raw", action="store_true", help="Treat input as an already compacted raw body.")
    args = parser.parse_args()

    body = args.json
    if not body:
        body = input()
    if not args.raw:
        body = compact_json_text(json.loads(body))
    print(track_key_for_text(body))


if __name__ == "__main__":
    main()
