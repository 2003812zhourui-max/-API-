"""KPI 流水线 —— 一键执行：导出 → 下载 → 清洗 → 输出所有 KPI 报表

流程:
  1. 导出出库单 Excel
  2. 下载导出结果
  3. 清洗 → 出库单波次KPI (clean_delivery_kpi.py)
  4. 清洗 → 拣货人绩效KPI (clean_picker_performance_kpi.py)
  5. 清洗 → 出库时效KPI (clean_delivery_leadtime_kpi.py)

用法:
  python run_kpi_pipeline.py --kpi-date 2026-06-08 --storage-state storage_state.json --fetch-picker-names
  python run_kpi_pipeline.py --kpi-date 2026-06-08 --input delivery_20260608.xlsx --skip-export --fetch-picker-names
"""
from __future__ import annotations

import argparse
import os
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

# 确保项目根目录在 sys.path
PROJECT_DIR = Path(__file__).resolve().parent
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))


def run_export(args: argparse.Namespace, kpi_day: date, output_dir: Path) -> Path:
    """步骤 1-2: 导出出库单并下载"""
    from export_delivery_once import (
        build_payload,
        create_export_task,
        download_when_ready,
        session_from_args,
    )

    start_at, end_at = _kpi_window(kpi_day)
    start_time = start_at.strftime("%Y-%m-%d %H:%M:%S")
    end_time = end_at.strftime("%Y-%m-%d %H:%M:%S")

    print(f"\n{'='*60}")
    print(f"STEP 1: Export delivery orders")
    print(f"  Window: {start_time} ~ {end_time}")
    print(f"{'='*60}")

    export_args = argparse.Namespace(
        storage_state=args.storage_state,
        wh_code=args.wh_code,
        status=args.status or "",
    )
    session = session_from_args(export_args)

    payload = build_payload(start_time, end_time, args.wh_code or "", args.status or "")
    task_id = create_export_task(session, payload)
    print(f"  Export task created: {task_id}")

    output_path = output_dir / f"delivery_{kpi_day:%Y%m%d}.xlsx"
    downloaded = download_when_ready(session, task_id, output_path, args.export_attempts, args.export_interval)
    print(f"  Downloaded: {downloaded}")
    return downloaded


def run_kpi_delivery(args: argparse.Namespace, input_path: Path, output_dir: Path, kpi_day: date) -> Path:
    """步骤 3: 出库单波次 KPI"""
    from clean_delivery_kpi import clean_workbook, parse_kpi_date

    print(f"\n{'='*60}")
    print(f"STEP 3: Clean delivery KPI (wave summary + picker names)")
    print(f"{'='*60}")

    output_path = output_dir / f"{input_path.stem}_kpi_cleaned.xlsx"

    clean_args = argparse.Namespace(
        fetch_picker_names=args.fetch_picker_names,
        storage_state=args.storage_state,
        name=args.name,
        picker_start_time=args.picker_start_time,
        picker_end_time=args.picker_end_time,
        picker_page_size=args.picker_page_size,
        picker_sample_limit=args.picker_sample_limit,
        picker_wave_limit=args.picker_wave_limit,
        picker_sleep=args.picker_sleep,
    )
    clean_workbook(input_path, output_path, kpi_day, clean_args)
    return output_path


def run_kpi_picker_performance(args: argparse.Namespace, input_path: Path, output_dir: Path, kpi_day: date) -> Path:
    """步骤 4: 拣货人绩效 KPI"""
    from clean_picker_performance_kpi import clean_picker_performance

    print(f"\n{'='*60}")
    print(f"STEP 4: Picker performance KPI")
    print(f"{'='*60}")

    output_path = output_dir / f"{input_path.stem}_picker_performance.xlsx"

    perf_args = argparse.Namespace(
        fetch_picker_names=args.fetch_picker_names,
        storage_state=args.storage_state,
        picker_start_time=args.picker_start_time,
        picker_end_time=args.picker_end_time,
        picker_page_size=args.picker_page_size,
        picker_sample_limit=args.picker_sample_limit,
        picker_wave_limit=args.picker_wave_limit,
        picker_sleep=args.picker_sleep,
    )
    clean_picker_performance(input_path, output_path, kpi_day, perf_args)
    return output_path


def run_kpi_leadtime(args: argparse.Namespace, input_path: Path, output_dir: Path, kpi_day: date) -> Path:
    """步骤 5: 出库时效 KPI"""
    from clean_delivery_leadtime_kpi import clean_leadtime

    print(f"\n{'='*60}")
    print(f"STEP 5: Delivery lead time KPI")
    print(f"{'='*60}")

    output_path = output_dir / f"{input_path.stem}_leadtime_kpi.xlsx"
    clean_leadtime(input_path, output_path, kpi_day)
    return output_path


def _kpi_window(kpi_day: date) -> tuple[datetime, datetime]:
    """计算 KPI 时间窗口"""
    start_at = datetime.combine(kpi_day - timedelta(days=1), datetime.strptime("22:30:00", "%H:%M:%S").time())
    if kpi_day.weekday() == 4:
        end_at = datetime.combine(kpi_day, datetime.strptime("23:59:59", "%H:%M:%S").time())
    else:
        end_at = datetime.combine(kpi_day, datetime.strptime("22:30:00", "%H:%M:%S").time())
    return start_at, end_at


def main() -> None:
    parser = argparse.ArgumentParser(description="OMP 出库 KPI 流水线")
    parser.add_argument("--kpi-date", required=True, help="KPI date, YYYY-MM-DD")
    parser.add_argument("--input", default="", help="Skip export, use existing Excel file")
    parser.add_argument("--skip-export", action="store_true", help="Skip export step (requires --input)")
    parser.add_argument("--output-dir", default="", help="Output directory (default: ./output)")
    parser.add_argument("--name", default="", help="Fallback picker name")

    # 导出参数
    parser.add_argument("--storage-state", default="", help="Playwright storage_state JSON")
    parser.add_argument("--wh-code", default="")
    parser.add_argument("--status", default="")
    parser.add_argument("--export-attempts", type=int, default=30)
    parser.add_argument("--export-interval", type=int, default=10)

    # 拣货人查询参数
    parser.add_argument("--fetch-picker-names", action="store_true", help="Fetch picker names from OMP API")
    parser.add_argument("--picker-start-time", default="")
    parser.add_argument("--picker-end-time", default="")
    parser.add_argument("--picker-page-size", type=int, default=20)
    parser.add_argument("--picker-sample-limit", type=int, default=3)
    parser.add_argument("--picker-wave-limit", type=int, default=0)
    parser.add_argument("--picker-sleep", type=float, default=0.0)

    # 跳过特定步骤
    parser.add_argument("--skip-delivery-kpi", action="store_true")
    parser.add_argument("--skip-picker-kpi", action="store_true")
    parser.add_argument("--skip-leadtime-kpi", action="store_true")

    args = parser.parse_args()

    kpi_day = datetime.strptime(args.kpi_date, "%Y-%m-%d").date()

    # 输出目录
    if args.output_dir:
        output_dir = Path(args.output_dir).resolve()
    else:
        output_dir = PROJECT_DIR / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    # 步骤 1-2: 导出
    if args.input:
        input_path = Path(args.input).resolve()
        if not input_path.exists():
            print(f"ERROR: Input file not found: {input_path}")
            sys.exit(1)
        print(f"Using existing input: {input_path}")
    elif args.skip_export:
        # 尝试找最近导出的文件
        candidates = sorted(output_dir.glob("delivery_*.xlsx"), reverse=True)
        if candidates:
            input_path = candidates[0]
            print(f"Using latest export: {input_path}")
        else:
            print("ERROR: No existing export found. Run with --storage-state to export first.")
            sys.exit(1)
    else:
        input_path = run_export(args, kpi_day, output_dir)

    results: dict[str, Path] = {"export": input_path}

    # 步骤 3: 出库单 KPI
    if not args.skip_delivery_kpi:
        try:
            results["delivery_kpi"] = run_kpi_delivery(args, input_path, output_dir, kpi_day)
        except Exception as e:
            print(f"WARNING: Delivery KPI failed: {e}")

    # 步骤 4: 拣货人绩效 KPI
    if not args.skip_picker_kpi:
        try:
            results["picker_kpi"] = run_kpi_picker_performance(args, input_path, output_dir, kpi_day)
        except Exception as e:
            print(f"WARNING: Picker performance KPI failed: {e}")

    # 步骤 5: 出库时效 KPI
    if not args.skip_leadtime_kpi:
        try:
            results["leadtime_kpi"] = run_kpi_leadtime(args, input_path, output_dir, kpi_day)
        except Exception as e:
            print(f"WARNING: Lead time KPI failed: {e}")

    print(f"\n{'='*60}")
    print(f"PIPELINE COMPLETE")
    print(f"KPI date: {kpi_day:%Y-%m-%d}")
    print(f"Output directory: {output_dir}")
    for step, path in results.items():
        print(f"  {step}: {path.name}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
