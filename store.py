"""任务存储：一张表记任务，一张表记上传的文件。

本机单机运行、任务量小，用 SQLite 就够，不需要额外服务。
每个连接用完即关；写操作用一把进程内锁串行化。
"""
import json
import hashlib
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
    options_json  TEXT,
    deleted_at    TEXT,
    archived_status TEXT,
    files_revision INTEGER DEFAULT 0
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
    why         TEXT,
    sha256      TEXT,
    file_size   INTEGER,
    added_at    TEXT,
    active      INTEGER DEFAULT 1
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
    duration_seconds REAL,
    files_revision INTEGER,
    files_snapshot_json TEXT,
    params_json TEXT,
    output_sha256 TEXT,
    is_latest INTEGER DEFAULT 0,
    deleted_at TEXT
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
        for table, fields in {
            "tasks": {
                "run_dir": "TEXT", "run_id": "TEXT", "deleted_at": "TEXT",
                "archived_status": "TEXT", "files_revision": "INTEGER DEFAULT 0",
            },
            "task_files": {
                "sha256": "TEXT", "file_size": "INTEGER", "added_at": "TEXT",
                "active": "INTEGER DEFAULT 1",
            },
            "task_runs": {
                "files_revision": "INTEGER", "files_snapshot_json": "TEXT",
                "params_json": "TEXT", "output_sha256": "TEXT",
                "is_latest": "INTEGER DEFAULT 0", "deleted_at": "TEXT",
            },
        }.items():
            cols = {r["name"] for r in c.execute("PRAGMA table_info(%s)" % table)}
            for name, definition in fields.items():
                if name not in cols:
                    c.execute("ALTER TABLE %s ADD COLUMN %s %s" %
                              (table, name, definition))
        c.execute("UPDATE tasks SET files_revision=0 WHERE files_revision IS NULL")
        c.execute("UPDATE task_files SET active=1 WHERE active IS NULL")
        c.execute("UPDATE task_files SET added_at=COALESCE(added_at, ?)", (_now(),))
        rows = c.execute("SELECT id, stored_path FROM task_files "
                         "WHERE sha256 IS NULL AND stored_path IS NOT NULL").fetchall()
        for row in rows:
            try:
                digest = _sha256(row["stored_path"])
                size = os.path.getsize(row["stored_path"])
            except OSError:
                continue
            c.execute("UPDATE task_files SET sha256=?, file_size=? WHERE id=?",
                      (digest, size, row["id"]))
        try:
            c.execute("CREATE INDEX IF NOT EXISTS idx_runs_task_created "
                      "ON task_runs(task_id, created_at DESC)")
            c.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_task_files_sha "
                      "ON task_files(task_id, sha256) WHERE active=1 AND sha256 IS NOT NULL")
        except sqlite3.IntegrityError:
            # 极旧数据库可能已经有重复记录；查询去重仍可工作，后续追加不再制造新重复。
            pass


def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _sha256(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


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
    return list_tasks_filtered(limit=limit, include_deleted=False)


def list_tasks_filtered(limit=200, include_deleted=False):
    with _conn() as c:
        where = "" if include_deleted else "WHERE deleted_at IS NULL"
        rows = c.execute("SELECT * FROM tasks %s ORDER BY created_at DESC LIMIT ?" % where,
                         (limit,)).fetchall()
    return [dict(r) for r in rows]


def add_file(task_id, orig_name, stored_path, kind=None, kind_cn=None,
             period=None, company=None, code=None, blocking=0, why=None):
    digest = _sha256(stored_path)
    size = os.path.getsize(stored_path)
    with _LOCK, _conn() as c:
        duplicate = c.execute("SELECT id FROM task_files WHERE task_id=? AND sha256=? "
                              "AND active=1", (task_id, digest)).fetchone()
        if duplicate:
            return False
        c.execute(
            "INSERT INTO task_files (task_id, orig_name, stored_path, kind, "
            "kind_cn, period, company, code, blocking, why, sha256, file_size, added_at, active) "
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,1)",
            (task_id, orig_name, stored_path, kind, kind_cn, period, company,
             code, int(blocking), why, digest, size, _now()))
    return True


def file_by_sha(task_id, digest):
    with _conn() as c:
        row = c.execute("SELECT * FROM task_files WHERE task_id=? AND sha256=? "
                        "AND active=1", (task_id, digest)).fetchone()
    return dict(row) if row else None


def update_file(task_id, stored_path, **fields):
    if not fields:
        return
    fields["active"] = 1
    cols = ", ".join("%s = ?" % key for key in fields)
    with _LOCK, _conn() as c:
        c.execute("UPDATE task_files SET %s WHERE task_id=? AND stored_path=?" % cols,
                  list(fields.values()) + [task_id, stored_path])


def files_for(task_id):
    with _conn() as c:
        rows = c.execute(
            "SELECT * FROM task_files WHERE task_id = ? AND active=1 ORDER BY id", (task_id,)
        ).fetchall()
    return [dict(r) for r in rows]


def runs_for(task_id):
    with _conn() as c:
        rows = c.execute(
            "SELECT * FROM task_runs WHERE task_id=? AND deleted_at IS NULL "
            "ORDER BY created_at DESC", (task_id,)).fetchall()
    return [dict(r) for r in rows]


def get_run(run_id):
    with _conn() as c:
        row = c.execute("SELECT * FROM task_runs WHERE id=? AND deleted_at IS NULL",
                        (run_id,)).fetchone()
    return dict(row) if row else None


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
        task_row = connection.execute(
            "SELECT files_revision FROM tasks WHERE id=? AND deleted_at IS NULL",
            (task_id,)).fetchone()
        if not task_row:
            return None
        snapshot = [dict(row) for row in connection.execute(
            "SELECT id,orig_name,stored_path,kind,kind_cn,period,company,code,sha256,file_size "
            "FROM task_files WHERE task_id=? AND active=1 ORDER BY id", (task_id,)).fetchall()]
        result = connection.execute(
            "UPDATE tasks SET status='queued', step=0, step_name='排队中', "
            "run_id=?, run_dir=NULL, xlsm_path=NULL, xlsm_name=NULL, verdict=NULL, "
            "message=NULL, log_tail=NULL, updated_at=?, options_json=?, "
            "company_name=?, company_code=?, report_period=? "
            "WHERE id=? AND deleted_at IS NULL AND status NOT IN ('queued','running','uploading')",
            (run_id, _now(), json.dumps(options, ensure_ascii=False),
             options.get('company_name'), options.get('company_code'),
             options.get('report_period'), task_id))
        if result.rowcount != 1:
            return None
        connection.execute("UPDATE task_runs SET is_latest=0 WHERE task_id=?", (task_id,))
        connection.execute(
            "INSERT INTO task_runs (id,task_id,created_at,status,files_revision,"
            "files_snapshot_json,params_json,is_latest) VALUES (?,?,?,?,?,?,?,1)",
            (run_id, task_id, _now(), 'queued', task_row["files_revision"] or 0,
             json.dumps(snapshot, ensure_ascii=False), json.dumps(options, ensure_ascii=False)))
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


def bump_files_revision(task_id):
    with _LOCK, _conn() as c:
        c.execute("UPDATE tasks SET files_revision=COALESCE(files_revision,0)+1, "
                  "updated_at=? WHERE id=?", (_now(), task_id))


def delete_task(task_id):
    """永久删除任务记录，返回任务目录；运行中的任务拒绝删除。"""
    with _LOCK, _conn() as c:
        row = c.execute("SELECT status, workdir FROM tasks WHERE id=?", (task_id,)).fetchone()
        if not row:
            return False, "任务不存在", None
        if row["status"] in ("queued", "running", "uploading"):
            return False, "任务正在运行，不能删除", None
        c.execute("DELETE FROM task_runs WHERE task_id=?", (task_id,))
        c.execute("DELETE FROM task_files WHERE task_id=?", (task_id,))
        c.execute("DELETE FROM tasks WHERE id=?", (task_id,))
    return True, None, row["workdir"]
