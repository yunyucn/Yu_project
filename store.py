"""任务存储：一张表记任务，一张表记上传的文件。

本机单机运行、任务量小，用 SQLite 就够，不需要额外服务。
每个连接用完即关；写操作用一把进程内锁串行化。
"""
import json
import os
import sqlite3
import threading
import uuid
from datetime import datetime
from contextlib import contextmanager

import config

_LOCK = threading.Lock()

SCHEMA = """
CREATE TABLE IF NOT EXISTS tasks (
    id            TEXT PRIMARY KEY,
    created_at    TEXT NOT NULL,
    updated_at    TEXT,
    company_name  TEXT,
    company_code  TEXT,
    report_period TEXT,
    status        TEXT NOT NULL,
    step          INTEGER DEFAULT 0,
    total_steps   INTEGER DEFAULT 11,
    step_name     TEXT,
    verdict       TEXT,
    headline      TEXT,
    message       TEXT,
    workdir       TEXT,
    run_dir       TEXT,
    xlsm_path     TEXT,
    xlsm_name     TEXT,
    log_tail      TEXT,
    params_json   TEXT,
    options_json  TEXT
);
CREATE TABLE IF NOT EXISTS task_files (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    task_id     TEXT NOT NULL,
    orig_name   TEXT,
    stored_path TEXT,
    kind        TEXT,
    kind_cn     TEXT,
    period      TEXT,
    company     TEXT,
    code        TEXT,
    blocking    INTEGER DEFAULT 0,
    why         TEXT
);
CREATE INDEX IF NOT EXISTS idx_files_task ON task_files(task_id);
CREATE TABLE IF NOT EXISTS task_runs (
    id TEXT PRIMARY KEY,
    task_id TEXT NOT NULL,
    created_at TEXT NOT NULL,
    finished_at TEXT,
    status TEXT NOT NULL,
    run_dir TEXT,
    message TEXT,
    exit_code INTEGER,
    duration_seconds REAL
);
"""


@contextmanager
def _conn():
    connection = sqlite3.connect(config.DB_PATH, timeout=30)
    connection.row_factory = sqlite3.Row
    try:
        with connection:
            yield connection
    finally:
        connection.close()


def init():
    config.ensure_dirs()
    with _LOCK, _conn() as c:
        c.executescript(SCHEMA)
        # 老库是 CREATE TABLE IF NOT EXISTS 建出来的，加字段只能靠 ALTER。
        cols = {r["name"] for r in c.execute("PRAGMA table_info(tasks)")}
        if "run_dir" not in cols:
            c.execute("ALTER TABLE tasks ADD COLUMN run_dir TEXT")
        if "run_id" not in cols:
            c.execute("ALTER TABLE tasks ADD COLUMN run_id TEXT")


def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def new_id():
    return datetime.now().strftime("%Y%m%d_%H%M%S_") + uuid.uuid4().hex[:6]


def create_task(company_name=None, company_code=None, report_period=None,
                workdir=None):
    tid = new_id()
    with _LOCK, _conn() as c:
        c.execute(
            "INSERT INTO tasks (id, created_at, updated_at, company_name, "
            "company_code, report_period, status, workdir) "
            "VALUES (?,?,?,?,?,?,?,?)",
            (tid, _now(), _now(), company_name, company_code, report_period,
             "uploaded", workdir))
    return tid


def update_task(task_id, **fields):
    if not fields:
        return
    fields["updated_at"] = _now()
    cols = ", ".join("%s = ?" % k for k in fields)
    vals = list(fields.values()) + [task_id]
    with _LOCK, _conn() as c:
        c.execute("UPDATE tasks SET %s WHERE id = ?" % cols, vals)


def get_task(task_id):
    with _conn() as c:
        row = c.execute("SELECT * FROM tasks WHERE id = ?", (task_id,)).fetchone()
    return dict(row) if row else None


def list_tasks(limit=200):
    with _conn() as c:
        rows = c.execute(
            "SELECT * FROM tasks ORDER BY created_at DESC LIMIT ?", (limit,)
        ).fetchall()
    return [dict(r) for r in rows]


def add_file(task_id, orig_name, stored_path, kind=None, kind_cn=None,
             period=None, company=None, code=None, blocking=0, why=None):
    with _LOCK, _conn() as c:
        c.execute(
            "INSERT INTO task_files (task_id, orig_name, stored_path, kind, "
            "kind_cn, period, company, code, blocking, why) "
            "VALUES (?,?,?,?,?,?,?,?,?,?)",
            (task_id, orig_name, stored_path, kind, kind_cn, period, company,
             code, int(blocking), why))


def files_for(task_id):
    with _conn() as c:
        rows = c.execute(
            "SELECT * FROM task_files WHERE task_id = ? ORDER BY id", (task_id,)
        ).fetchall()
    return [dict(r) for r in rows]


def set_options(task_id, options):
    update_task(task_id, options_json=json.dumps(options, ensure_ascii=False))


def get_options(task):
    raw = task.get("options_json")
    if not raw:
        return {}
    try:
        return json.loads(raw)
    except ValueError:
        return {}


def pending_count():
    with _conn() as c:
        row = c.execute(
            "SELECT COUNT(*) AS n FROM tasks WHERE status IN ('queued','running')"
        ).fetchone()
    return row["n"] if row else 0


def claim_task(task_id, options):
    run_id = uuid.uuid4().hex
    with _LOCK, _conn() as connection:
        result = connection.execute(
            "UPDATE tasks SET status='queued', step=0, step_name='排队中', "
            "run_id=?, run_dir=NULL, xlsm_path=NULL, xlsm_name=NULL, verdict=NULL, "
            "message=NULL, log_tail=NULL, updated_at=?, options_json=?, "
            "company_name=?, company_code=?, report_period=? "
            "WHERE id=? AND status NOT IN ('queued','running','uploading')",
            (run_id, _now(), json.dumps(options, ensure_ascii=False),
             options.get('company_name'), options.get('company_code'),
             options.get('report_period'), task_id))
        if result.rowcount != 1:
            return None
        connection.execute(
            "INSERT INTO task_runs (id,task_id,created_at,status) VALUES (?,?,?,?)",
            (run_id, task_id, _now(), 'queued'))
    return run_id


def update_run(run_id, **fields):
    if not run_id or not fields:
        return
    columns = ', '.join('%s=?' % key for key in fields)
    with _LOCK, _conn() as connection:
        connection.execute('UPDATE task_runs SET %s WHERE id=?' % columns,
                           list(fields.values()) + [run_id])


def recover_interrupted():
    message = '服务曾中断，请重新生成；若文件尚未识别，请重新上传。'
    with _LOCK, _conn() as connection:
        rows = connection.execute(
            "SELECT id,run_id,workdir,run_dir FROM tasks WHERE status IN ('queued','running','uploading')").fetchall()
        connection.execute(
            "UPDATE tasks SET status='failed', message=?, step_name='已中断', "
            "xlsm_path=NULL, xlsm_name=NULL, updated_at=? "
            "WHERE status IN ('queued','running','uploading')", (message, _now()))
        connection.execute(
            "UPDATE task_runs SET status='failed', message=?, finished_at=? "
            "WHERE status IN ('queued','running')", (message, _now()))
    return [dict(row) for row in rows]
