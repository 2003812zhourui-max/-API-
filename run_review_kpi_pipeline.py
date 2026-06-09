from __future__ import annotations

import argparse
from argparse import Namespace
from datetime import datetime
from pathlib import Path

from clean_review_kpi import clean_review_workbook, create_time_window
from clean_delivery_kpi import parse_kpi_date
from export_delivery_once import build_payload, create_export_task, download_when_ready, session_from_args, wh_code_from_storage_state


def default_output_dir(kpi_day) -> Path:
    return Path.cwd() / "output" / "kpi" / kpi_day.strftime("%Y%m%d") / "review"


def download_delivery_excel(
    *,
    kpi_date: str,
    storage_state: str,
    output_path: Path,
    wh_code: str,
    status: str,
    attempts: int,
    interval: int,
    start_time: str,
    end_time: str,
) -> Path:
    args = Namespace(storage_state=storage_state)
    session = session_from_args(args)
    resolved_wh_code = wh_code or wh_code_from_storage_state(storage_state)
    payload = build_payload(start_time, end_time, resolved_wh_code, status)
    task_id = create_export_task(session, payload)
    print(f"download step: export task created: {task_id}", flush=True)

    downloaded = download_when_ready(session, task_id, output_path, attempts=attempts, interval=interval)
    print(f"download step: file downloaded: {downloaded}", flush=True)
    return downloaded


def build_review_args(args: argparse.Namespace) -> Namespace:
    return Namespace(
        name=args.name,
        fetch_reviewer_names=True,
        storage_state=args.storage_state,
        order_start_time=args.order_start_time,
        order_end_time=args.order_end_time,
        reviewer_page_size=args.reviewer_page_size,
        reviewer_sample_limit=args.reviewer_sample_limit,
        reviewer_wave_limit=args.reviewer_wave_limit,
        reviewer_sleep=args.reviewer_sleep,
    )


def run_pipeline(args: argparse.Namespace) -> None:
    kpi_day = parse_kpi_date(args.kpi_date)
    output_dir = Path(args.output_dir).resolve() if args.output_dir else default_output_dir(kpi_day).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    create_start, create_end = create_time_window(kpi_day)
    export_start = args.export_start_time or create_start.strftime("%Y-%m-%d %H:%M:%S")
    export_end = args.export_end_time or create_end.strftime("%Y-%m-%d %H:%M:%S")

    print(f"KPI date: {args.kpi_date}", flush=True)
    print(f"export window: {export_start} ~ {export_end}", flush=True)

    if args.input:
        raw_path = Path(args.input).resolve()
        print(f"download step: skipped, using input file: {raw_path}", flush=True)
    else:
        storage_state = Path(args.storage_state).resolve()
        if not storage_state.exists():
            raise FileNotFoundError(f"storage_state not found: {storage_state}")
        args.storage_state = str(storage_state)
        raw_path = output_dir / f"delivery_raw_{kpi_day:%Y%m%d}_{datetime.now():%H%M%S}.xlsx"
        raw_path = download_delivery_excel(
            kpi_date=args.kpi_date,
            storage_state=args.storage_state,
            output_path=raw_path,
            wh_code=args.wh_code,
            status=args.status,
            attempts=args.attempts,
            interval=args.interval,
            start_time=export_start,
            end_time=export_end,
        )

    review_path = output_dir / f"review_kpi_{kpi_day:%Y%m%d}.xlsx"
    print("filter/clean step: start review KPI", flush=True)
    clean_review_workbook(raw_path, review_path, kpi_day, build_review_args(args))

    print("done", flush=True)
    print(f"raw file: {raw_path}", flush=True)
    print(f"review KPI file: {review_path}", flush=True)
    print("wave result sheet: 复核波次号去重统计", flush=True)
    print("person result sheet: 复核人员汇总", flush=True)


def main() -> None:
    parser = argparse.ArgumentParser(description="Download, filter, clean, and resolve wave/reviewer data for review KPI.")
    parser.add_argument("--kpi-date", required=True, help="KPI date, YYYY-MM-DD")
    parser.add_argument("--input", default="", help="Use an existing delivery Excel instead of downloading")
    parser.add_argument("--storage-state", default="storage_state.json")
    parser.add_argument("--output-dir", default="")
    parser.add_argument("--name", default="未匹配", help="Fallback name when reviewer lookup fails")

    parser.add_argument("--export-start-time", default="", help="Override export start time")
    parser.add_argument("--export-end-time", default="", help="Override export end time")
    parser.add_argument("--wh-code", default="")
    parser.add_argument("--status", default="100")
    parser.add_argument("--attempts", type=int, default=30)
    parser.add_argument("--interval", type=int, default=10)

    parser.add_argument("--order-start-time", default="", help="Override OMP order search start time for name lookup")
    parser.add_argument("--order-end-time", default="", help="Override OMP order search end time for name lookup")
    parser.add_argument("--reviewer-page-size", type=int, default=100)
    parser.add_argument("--reviewer-sample-limit", type=int, default=3)
    parser.add_argument("--reviewer-wave-limit", type=int, default=0, help="Only fetch first N waves; 0 means all")
    parser.add_argument("--reviewer-sleep", type=float, default=0.0)
    args = parser.parse_args()

    run_pipeline(args)


if __name__ == "__main__":
    main()
