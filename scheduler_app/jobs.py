"""任务定义 —— 每个任务就是一个函数，注册到 JOBS 列表即可。"""
import logging
from argparse import Namespace
from datetime import datetime, timedelta
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = PROJECT_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

logger = logging.getLogger(__name__)


# ── 任务 1: 出库单导出 ────────────────────────────────────
def export_delivery():
    from export_delivery_once import (
        build_payload,
        create_export_task,
        download_when_ready,
        session_from_args,
    )

    storage_state = PROJECT_DIR / "storage_state.json"
    if not storage_state.exists():
        raise FileNotFoundError(f"未找到 storage_state.json，请先导出 Playwright 状态文件到 {storage_state}")

    args = Namespace(storage_state=str(storage_state), wh_code="")
    session = session_from_args(args)

    yesterday = datetime.now() - timedelta(days=1)
    start = yesterday.strftime("%Y-%m-%d 00:00:00")
    end = yesterday.strftime("%Y-%m-%d 23:59:59")

    payload = build_payload(start, end, "", "")
    task_id = create_export_task(session, payload)
    logger.info(f"导出任务已创建: {task_id}")

    output_path = OUTPUT_DIR / f"delivery_{yesterday:%Y%m%d}.xlsx"
    download_when_ready(session, task_id, output_path, attempts=30, interval=10)
    logger.info(f"已下载到: {output_path}")


# ── 任务注册表 ────────────────────────────────────────────
JOBS = [
    {
        "name": "export_delivery",
        "display": "出库单导出",
        "description": "从领星 OMP 导出前一天的出库单（约 4-5 MB Excel）",
        "schedule_display": "每天 09:00",
        "cron": "0 9 * * *",
        "func": export_delivery,
    },
    # 后续任务示例（取消注释即可启用）:
    #
    # {
    #     "name": "clean_excel",
    #     "display": "Excel 清洗",
    #     "description": "对导出的 Excel 做字段映射和格式化",
    #     "schedule_display": "每天 09:30（导出完成后）",
    #     "cron": "30 9 * * *",
    #     "func": clean_excel,
    # },
    # {
    #     "name": "feishu_upload",
    #     "display": "飞书上传",
    #     "description": "将清洗后的数据上传到飞书多维表格",
    #     "schedule_display": "每天 10:00",
    #     "cron": "0 10 * * *",
    #     "func": feishu_upload,
    # },
]
