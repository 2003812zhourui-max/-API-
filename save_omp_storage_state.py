from __future__ import annotations

import argparse
import os
import time
from pathlib import Path

from playwright.sync_api import sync_playwright


API_ORIGIN = "https://omp.xlwms.com"


def token_is_present(page) -> bool:
    try:
        token = page.evaluate("() => localStorage.getItem('omp-token') || ''")
    except Exception:
        return False
    return bool(str(token).strip())


def main() -> None:
    parser = argparse.ArgumentParser(description="Login to OMP once and save Playwright storage_state.")
    parser.add_argument(
        "--output",
        default=r"C:\Users\Administrator\Downloads\PDF_D\WMS_PDF_Tool\wms_storage_state.json",
    )
    parser.add_argument("--timeout", type=int, default=600)
    args = parser.parse_args()

    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)

    proxy_server = os.environ.get("HTTPS_PROXY") or os.environ.get("HTTP_PROXY") or ""
    launch_kwargs = {"headless": False}
    if proxy_server:
        launch_kwargs["proxy"] = {"server": proxy_server}

    with sync_playwright() as p:
        try:
            browser = p.chromium.launch(channel="chrome", **launch_kwargs)
        except Exception:
            browser = p.chromium.launch(**launch_kwargs)
        context = browser.new_context()
        page = context.new_page()
        page.goto(f"{API_ORIGIN}/globalOrder/parcel", wait_until="domcontentloaded")

        print("Browser opened. Please log in to OMP in that browser window.", flush=True)
        print(f"Waiting up to {args.timeout} seconds for omp-token...", flush=True)

        deadline = time.time() + args.timeout
        while time.time() < deadline:
            if page.url.startswith(API_ORIGIN) and token_is_present(page):
                context.storage_state(path=str(output))
                print(f"saved storage_state: {output}", flush=True)
                browser.close()
                return
            time.sleep(2)

        browser.close()
        raise TimeoutError("Timed out waiting for OMP login. Please run again and finish login in the browser.")


if __name__ == "__main__":
    main()
