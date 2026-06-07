from __future__ import annotations

import argparse
from argparse import Namespace
from datetime import date, datetime, timedelta
from pathlib import Path

from clean_delivery_kpi import clean_workbook, kpi_window, parse_kpi_date
from export_delivery_once import build_payload, create_export_task, download_when_ready, session_from_args


def latest_completed_kpi_date(today: date | None = None) -> date:
    """Return the latest weekday KPI date that should already be complete."""
    current = today or date.today()
    candidate = current - timedelta(days=1)
    while candidate.weekday() >= 5:
        candidate -= timedelta(days=1)
    return candidate


def clean_args_from_pipeline_args(args: argparse.Namespace) -> Namespace:
    return Namespace(
        fetch_picker_names=args.fetch_picker_names,
        storage_state=args.storage_state,
        picker_start_time=args.picker_start_time,
        picker_end_time=args.picker_end_time,
        picker_page_size=args.picker_page_size,
        picker_sample_limit=args.picker_sample_limit,
        picker_wave_limit=args.picker_wave_limit,
        picker_sleep=args.picker_sleep,
        name=args.name,
    )


def run_delivery_kpi(args: argparse.Namespace) -> tuple[Path, Path]:
    kpi_day = parse_kpi_date(args.kpi_date) if args.kpi_date else latest_completed_kpi_date()
    start_at, end_at = kpi_window(kpi_day)
    start_time = start_at.strftime("%Y-%m-%d %H:%M:%S")
    end_time = end_at.strftime("%Y-%m-%d %H:%M:%S")

    output_dir = Path(args.output_dir).resolve()
    raw_dir = output_dir / "raw"
    kpi_dir = output_dir / "kpi"
    raw_dir.mkdir(parents=True, exist_ok=True)
    kpi_dir.mkdir(parents=True, exist_ok=True)

    export_args = Namespace(storage_state=args.storage_state, wh_code=args.wh_code)
    session = session_from_args(export_args)
    payload = build_payload(start_time, end_time, args.wh_code, args.status)
    task_id = create_export_task(session, payload)

    raw_path = raw_dir / f"delivery_kpi_raw_{kpi_day:%Y%m%d}_{task_id}.xlsx"
    cleaned_path = kpi_dir / f"delivery_kpi_{kpi_day:%Y%m%d}.xlsx"

    print(f"KPI date: {kpi_day:%Y-%m-%d}", flush=True)
    print(f"export window: {start_time} ~ {end_time}", flush=True)
    print(f"created task: {task_id}", flush=True)

    downloaded = download_when_ready(session, task_id, raw_path, args.attempts, args.interval)
    print(f"downloaded raw file: {downloaded}", flush=True)

    clean_workbook(downloaded, cleaned_path, kpi_day, clean_args_from_pipeline_args(args))
    print(f"cleaned KPI file: {cleaned_path}", flush=True)
    return downloaded, cleaned_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Export delivery Excel and build the delivery KPI workbook.")
    parser.add_argument("--kpi-date", default="", help="KPI date, YYYY-MM-DD. Default uses latest completed weekday.")
    parser.add_argument("--storage-state", default="", help="Optional Playwright storage_state JSON")
    parser.add_argument("--wh-code", default="")
    parser.add_argument("--status", default="")
    parser.add_argument("--output-dir", default="output")
    parser.add_argument("--attempts", type=int, default=30)
    parser.add_argument("--interval", type=int, default=10)
    parser.add_argument("--name", default="", help="Fallback picker name when picker name cannot be fetched")
    parser.add_argument("--fetch-picker-names", action="store_true", help="Fetch picker name from OMP progress logs")
    parser.add_argument("--picker-start-time", default="", help="Override wave search start time")
    parser.add_argument("--picker-end-time", default="", help="Override wave search end time")
    parser.add_argument("--picker-page-size", type=int, default=20)
    parser.add_argument("--picker-sample-limit", type=int, default=3)
    parser.add_argument("--picker-wave-limit", type=int, default=0, help="Only fetch the first N wave picker names; 0 means all")
    parser.add_argument("--picker-sleep", type=float, default=0.0)
    args = parser.parse_args()

    run_delivery_kpi(args)


if __name__ == "__main__":
    main()
