"""领星 OMP 自动化调度平台 — Flask + APScheduler + SQLite"""
from __future__ import annotations

import logging
import sqlite3
import sys
import traceback
from datetime import datetime
from io import StringIO
from pathlib import Path
from threading import Thread

from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from flask import Flask, jsonify, render_template_string, request

# 确保项目根目录在 sys.path 中，以便导入同目录模块
PROJECT_DIR = Path(__file__).resolve().parent.parent
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from scheduler_app.jobs import JOBS  # noqa: E402

# ── 目录 & 数据库 ──────────────────────────────────────────
DATA_DIR = PROJECT_DIR / "data"
LOG_DIR = PROJECT_DIR / "logs"
OUTPUT_DIR = PROJECT_DIR / "output"
for d in [DATA_DIR, LOG_DIR, OUTPUT_DIR]:
    d.mkdir(exist_ok=True)

DB_PATH = DATA_DIR / "scheduler.db"

app = Flask(__name__)
scheduler = BackgroundScheduler()

# ── 日志 ───────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.FileHandler(LOG_DIR / "scheduler.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger("scheduler")


# ── 数据库 ─────────────────────────────────────────────────
def get_db() -> sqlite3.Connection:
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    with get_db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS job_runs (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                job_name    TEXT    NOT NULL,
                status      TEXT    NOT NULL DEFAULT 'running',
                start_time  TEXT    NOT NULL,
                end_time    TEXT,
                message     TEXT,
                log         TEXT
            )
        """)


# ── 任务执行（带日志捕获）─────────────────────────────────
def run_job_with_logging(job_name: str, job_func) -> int:
    """在线程中执行任务，捕获日志和异常，写入数据库。返回 run_id。"""
    db = get_db()
    cursor = db.execute(
        "INSERT INTO job_runs (job_name, status, start_time, message) VALUES (?, 'running', ?, ?)",
        (job_name, datetime.now().isoformat(), "执行中…"),
    )
    run_id = cursor.lastrowid
    db.commit()
    db.close()

    # 捕获任务产生的日志
    log_buffer = StringIO()
    handler = logging.StreamHandler(log_buffer)
    handler.setLevel(logging.INFO)
    handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    task_logger = logging.getLogger(f"job.{job_name}")
    task_logger.addHandler(handler)
    task_logger.setLevel(logging.INFO)
    task_logger.propagate = False

    status = "success"
    message = "完成"
    try:
        job_func()
    except Exception:
        status = "failed"
        message = str(sys.exc_info()[1])
        task_logger.error("任务异常: %s", message)
        task_logger.error(traceback.format_exc())
    finally:
        task_logger.removeHandler(handler)
        end_time = datetime.now().isoformat()
        log_text = log_buffer.getvalue()
        db = get_db()
        db.execute(
            "UPDATE job_runs SET status=?, end_time=?, message=?, log=? WHERE id=?",
            (status, end_time, message, log_text, run_id),
        )
        db.commit()
        db.close()

    return run_id


# ── 路由: 仪表盘页面 ──────────────────────────────────────
_DASHBOARD_HTML = r"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>领星自动化调度平台</title>
<link href="https://cdn.bootcdn.net/ajax/libs/twitter-bootstrap/5.3.0/css/bootstrap.min.css" rel="stylesheet">
<link href="https://cdn.bootcdn.net/ajax/libs/bootstrap-icons/1.10.5/font/bootstrap-icons.css" rel="stylesheet">
<style>
  :root { --bs-body-bg: #f5f6fa; }
  body { background: var(--bs-body-bg); }
  .header { background: linear-gradient(135deg, #1a73e8, #0d47a1); color: #fff; padding: 20px 0; margin-bottom: 24px; }
  .header h4 { margin: 0; font-weight: 600; }
  .stat-card { border-radius: 12px; border: none; box-shadow: 0 1px 3px rgba(0,0,0,.08); }
  .stat-card .number { font-size: 2rem; font-weight: 700; }
  .job-card { border-radius: 12px; border: none; box-shadow: 0 1px 3px rgba(0,0,0,.08); margin-bottom: 16px; transition: box-shadow .2s; }
  .job-card:hover { box-shadow: 0 4px 16px rgba(0,0,0,.12); }
  .badge-status { font-size: .8rem; }
  .run-history { max-height: 300px; overflow-y: auto; }
  .log-viewer { background: #1e1e1e; color: #d4d4d4; font-family: Consolas,monospace; font-size: 13px; padding: 12px; border-radius: 8px; max-height: 400px; overflow-y: auto; white-space: pre-wrap; }
  .spinner-sm { width: 1rem; height: 1rem; border-width: .15em; }
</style>
</head>
<body>

<div class="header">
  <div class="container">
    <div class="d-flex justify-content-between align-items-center">
      <div>
        <h4><i class="bi bi-gear-wide-connected me-2"></i>领星自动化调度平台</h4>
        <small class="opacity-75">OMP Task Scheduler</small>
      </div>
      <div>
        <button class="btn btn-outline-light btn-sm me-2" onclick="refreshAll()"><i class="bi bi-arrow-clockwise me-1"></i>刷新</button>
        <span id="clock" class="small opacity-75"></span>
      </div>
    </div>
  </div>
</div>

<div class="container">

  <!-- 统计卡片 -->
  <div class="row mb-4" id="stats-row">
    <div class="col-md-3 mb-2">
      <div class="card stat-card p-3 text-center">
        <div class="text-muted small">任务总数</div>
        <div class="number text-primary" id="stat-total">-</div>
      </div>
    </div>
    <div class="col-md-3 mb-2">
      <div class="card stat-card p-3 text-center">
        <div class="text-muted small">今日成功</div>
        <div class="number text-success" id="stat-success">-</div>
      </div>
    </div>
    <div class="col-md-3 mb-2">
      <div class="card stat-card p-3 text-center">
        <div class="text-muted small">今日失败</div>
        <div class="number text-danger" id="stat-failed">-</div>
      </div>
    </div>
    <div class="col-md-3 mb-2">
      <div class="card stat-card p-3 text-center">
        <div class="text-muted small">下次执行</div>
        <div class="number text-info" id="stat-next" style="font-size:1.2rem">-</div>
      </div>
    </div>
  </div>

  <!-- 任务列表 -->
  <div id="jobs-container">加载中…</div>

</div>

<!-- 日志模态框 -->
<div class="modal fade" id="logModal" tabindex="-1">
  <div class="modal-dialog modal-lg modal-dialog-scrollable">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title">执行日志</h5>
        <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
      </div>
      <div class="modal-body">
        <div class="log-viewer" id="log-content">加载中…</div>
      </div>
    </div>
  </div>
</div>

<script src="https://cdn.bootcdn.net/ajax/libs/twitter-bootstrap/5.3.0/js/bootstrap.bundle.min.js"></script>
<script>
const API = '/api';

function fmtTime(iso) {
  if (!iso) return '-';
  const d = new Date(iso);
  return d.toLocaleString('zh-CN', {month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',second:'2-digit'});
}

function badge(status) {
  const map = { success: 'bg-success', failed: 'bg-danger', running: 'bg-warning text-dark' };
  const cls = map[status] || 'bg-secondary';
  return `<span class="badge ${cls} badge-status">${status}</span>`;
}

async function refreshAll() {
  try {
    const resp = await fetch(API + '/jobs');
    const jobs = await resp.json();
    renderJobs(jobs);
    updateClock();
  } catch (e) {
    document.getElementById('jobs-container').innerHTML =
      `<div class="alert alert-warning">无法连接调度服务: ${e.message}</div>`;
  }
}

function renderJobs(jobs) {
  const container = document.getElementById('jobs-container');
  if (!jobs.length) {
    container.innerHTML = '<div class="text-muted text-center py-5">暂无注册任务，请在 jobs.py 中添加</div>';
    return;
  }

  let total = jobs.length, success = 0, failed = 0;
  const today = new Date().toISOString().slice(0, 10);

  const cards = jobs.map(j => {
    const lr = j.last_run;
    const isToday = lr && lr.start_time && lr.start_time.startsWith(today);
    if (isToday && lr.status === 'success') success++;
    if (isToday && lr.status === 'failed') failed++;

    const isRunning = lr && lr.status === 'running';
    const statusIcon = isRunning ? '<span class="spinner-border spinner-sm text-warning me-1"></span>' : '';

    return `
    <div class="card job-card">
      <div class="card-body">
        <div class="d-flex justify-content-between align-items-start">
          <div>
            <h5 class="card-title mb-1">
              ${statusIcon}
              ${j.display}
              ${j.enabled ? '<span class="badge bg-success ms-1" style="font-size:.65rem">启用</span>' : '<span class="badge bg-secondary ms-1" style="font-size:.65rem">停用</span>'}
            </h5>
            <p class="text-muted small mb-2">${j.description}</p>
            <div class="d-flex gap-3 flex-wrap small text-muted">
              <span><i class="bi bi-clock me-1"></i>调度: ${j.schedule}</span>
              <span><i class="bi bi-arrow-right-circle me-1"></i>下次: ${j.next_run ? fmtTime(j.next_run) : '未调度'}</span>
              ${lr ? `<span><i class="bi bi-arrow-left-circle me-1"></i>上次: ${fmtTime(lr.start_time)} ${badge(lr.status)} ${lr.message||''}</span>` : '<span class="text-muted">尚无执行记录</span>'}
            </div>
          </div>
          <div class="d-flex gap-1">
            <button class="btn btn-primary btn-sm" onclick="triggerJob('${j.name}', this)" ${isRunning?'disabled':''}>
              <i class="bi bi-play-fill me-1"></i>立即执行
            </button>
            <button class="btn btn-outline-secondary btn-sm" onclick="loadHistory('${j.name}', this)" title="查看历史">
              <i class="bi bi-list-ul"></i>
            </button>
          </div>
        </div>
        <!-- 执行历史（动态插入） -->
        <div class="run-history mt-3" id="history-${j.name}" style="display:none"></div>
      </div>
    </div>`;
  }).join('');

  container.innerHTML = cards;

  // 更新统计
  document.getElementById('stat-total').textContent = total;
  document.getElementById('stat-success').textContent = success;
  document.getElementById('stat-failed').textContent = failed;
  const nexts = jobs.filter(j => j.next_run).map(j => j.next_run).sort();
  document.getElementById('stat-next').textContent = nexts.length ? fmtTime(nexts[0]) : '无';
}

async function triggerJob(name, btn) {
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner-border spinner-sm me-1"></span>执行中';
  try {
    const resp = await fetch(API + '/jobs/' + name + '/run', { method: 'POST' });
    const data = await resp.json();
    if (data.ok) {
      btn.classList.remove('btn-primary');
      btn.classList.add('btn-success');
      btn.innerHTML = '<i class="bi bi-check-lg me-1"></i>已触发';
      setTimeout(() => refreshAll(), 2000);
    } else {
      alert('触发失败: ' + JSON.stringify(data));
      btn.disabled = false;
      btn.innerHTML = '<i class="bi bi-play-fill me-1"></i>立即执行';
    }
  } catch (e) {
    alert('请求失败: ' + e.message);
    btn.disabled = false;
    btn.innerHTML = '<i class="bi bi-play-fill me-1"></i>立即执行';
  }
}

async function loadHistory(name, btn) {
  const el = document.getElementById('history-' + name);
  if (el.style.display === 'block') { el.style.display = 'none'; return; }

  try {
    const resp = await fetch(API + '/jobs/' + name + '/history');
    const rows = await resp.json();
    if (!rows.length) {
      el.innerHTML = '<p class="text-muted small mb-0">暂无记录</p>';
    } else {
      el.innerHTML = `
      <table class="table table-sm small mb-0">
        <thead><tr><th>时间</th><th>状态</th><th>结果</th><th>耗时</th><th></th></tr></thead>
        <tbody>
        ${rows.map(r => {
          const dur = r.end_time ? ((new Date(r.end_time) - new Date(r.start_time))/1000).toFixed(1) + 's' : '-';
          return `<tr>
            <td>${fmtTime(r.start_time)}</td>
            <td>${badge(r.status)}</td>
            <td class="text-truncate" style="max-width:200px">${r.message||''}</td>
            <td>${dur}</td>
            <td><button class="btn btn-outline-secondary btn-sm py-0" onclick="viewLog(${r.id})">日志</button></td>
          </tr>`;
        }).join('')}
        </tbody>
      </table>`;
    }
    el.style.display = 'block';
  } catch (e) {
    alert('加载历史失败: ' + e.message);
  }
}

async function viewLog(runId) {
  const modal = new bootstrap.Modal(document.getElementById('logModal'));
  document.getElementById('log-content').textContent = '加载中…';
  modal.show();
  try {
    const resp = await fetch(API + '/runs/' + runId + '/log');
    const data = await resp.json();
    document.getElementById('log-content').textContent = data.log || '(无日志)';
  } catch (e) {
    document.getElementById('log-content').textContent = '加载失败: ' + e.message;
  }
}

function updateClock() {
  document.getElementById('clock').textContent = new Date().toLocaleString('zh-CN');
}

setInterval(updateClock, 1000);
setInterval(refreshAll, 30000);  // 每 30 秒自动刷新
refreshAll();
</script>
</body>
</html>
"""


@app.route("/")
def dashboard():
    return render_template_string(_DASHBOARD_HTML)


# ── API ────────────────────────────────────────────────────

@app.route("/api/jobs")
def api_jobs():
    jobs_data = []
    for job_def in JOBS:
        job_name = job_def["name"]
        aps_job = scheduler.get_job(job_name)
        next_run = aps_job.next_run_time.isoformat() if aps_job and aps_job.next_run_time else None

        db = get_db()
        last_run = db.execute(
            "SELECT * FROM job_runs WHERE job_name=? ORDER BY id DESC LIMIT 1",
            (job_name,),
        ).fetchone()
        db.close()

        jobs_data.append({
            "name": job_name,
            "display": job_def.get("display", job_name),
            "description": job_def.get("description", ""),
            "schedule": job_def.get("schedule_display", ""),
            "enabled": aps_job is not None,
            "next_run": next_run,
            "last_run": dict(last_run) if last_run else None,
        })
    return jsonify(jobs_data)


@app.route("/api/jobs/<job_name>/run", methods=["POST"])
def api_trigger_job(job_name):
    job_def = next((j for j in JOBS if j["name"] == job_name), None)
    if not job_def:
        return jsonify({"error": "任务不存在"}), 404

    Thread(target=run_job_with_logging, args=(job_name, job_def["func"]), daemon=True).start()
    return jsonify({"ok": True, "message": f"任务 {job_name} 已触发"})


@app.route("/api/jobs/<job_name>/history")
def api_job_history(job_name):
    limit = request.args.get("limit", 20, type=int)
    db = get_db()
    rows = db.execute(
        "SELECT * FROM job_runs WHERE job_name=? ORDER BY id DESC LIMIT ?",
        (job_name, limit),
    ).fetchall()
    db.close()
    return jsonify([dict(r) for r in rows])


@app.route("/api/runs/<int:run_id>/log")
def api_run_log(run_id):
    db = get_db()
    row = db.execute("SELECT * FROM job_runs WHERE id=?", (run_id,)).fetchone()
    db.close()
    if not row:
        return jsonify({"error": "未找到记录"}), 404
    return jsonify(dict(row))


# ── 启动 ───────────────────────────────────────────────────
def start() -> None:
    init_db()
    logger.info("数据库初始化完成")

    for job_def in JOBS:
        if "cron" in job_def:
            scheduler.add_job(
                run_job_with_logging,
                trigger=CronTrigger.from_crontab(job_def["cron"]),
                args=[job_def["name"], job_def["func"]],
                id=job_def["name"],
                name=job_def.get("display", job_def["name"]),
                replace_existing=True,
            )
            logger.info("已注册定时任务: %s → %s", job_def["name"], job_def["cron"])

    scheduler.start()
    logger.info("调度器已启动")


def main() -> None:
    start()
    logger.info("Web 面板: http://localhost:5000")
    app.run(host="0.0.0.0", port=5000, debug=False)


if __name__ == "__main__":
    main()
