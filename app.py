"""A股估值模型填充 —— 本机 Web 版。

用户上传年报 PDF，系统调 pipeline/ 里的流水线，产出填好数的 .xlsm
和一张结论卡。

这一层只做四件事：收文件、认料、排队、把结论和文件递回去。
抽取、映射、写数、判定全部由 pipeline/ 完成。

服务端不需要 Excel、不需要 Windows：写工作簿是纯 Python 改 XML。产出的
.xlsm 里公式一个不少，但缓存值是空的，办公人员用 Excel 打开时 Excel 会
自己重算一遍（文件里写了 fullCalcOnLoad=1）。

启动：python app.py  然后打开 http://127.0.0.1:8000
Linux 部署：HOST=0.0.0.0 PORT=8000 python app.py
"""
import io
import os
import re
import logging
import math
import threading
import time
import uuid
import shutil
from datetime import date

from flask import (Flask, Response, abort, g, jsonify, render_template, request,
                   send_file, url_for)
from werkzeug.exceptions import HTTPException

import config
import prefill
import runner
import store
import worker
import runtime

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = config.MAX_UPLOAD_MB * 1024 * 1024
logger = logging.getLogger(__name__)
_upload_slot = threading.BoundedSemaphore(1)
_initialized = False

BAD_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f]')

CN_STATUS = {"uploading": "读取中", "uploaded": "待确认", "queued": "排队中", "running": "运行中",
             "done": "已完成", "review": "待复核", "stopped": "已停止", "failed": "失败"}
CN_VERDICT = {"full": "完整可用", "degraded": "降级可用",
              "facts_only": "只出事实层", "blocked": "被拦下"}


def status_cn(s):
    return CN_STATUS.get(s, s or "—")


def verdict_cn(v):
    return CN_VERDICT.get(v, v or "—")


app.jinja_env.globals.update(status_cn=status_cn, verdict_cn=verdict_cn,
                             max_mb=config.MAX_UPLOAD_MB)


def safe_name(name):
    """保留中文，只挡掉路径分隔符和控制字符。

    不能用 werkzeug 的 secure_filename：它会把中文整段丢掉，
    「兴源环境_AN2025.pdf」会变成「AN2025.pdf」。
    """
    name = os.path.basename(name or "").strip()
    name = BAD_CHARS.sub("_", name)
    return name or "unnamed.pdf"


def task_dir(task_id):
    return os.path.join(config.TASKS_DIR, task_id)


@app.before_request
def begin_request():
    g.request_id = uuid.uuid4().hex[:12]
    g.started = time.monotonic()


@app.after_request
def end_request(response):
    response.headers['X-Request-ID'] = g.request_id
    response.headers['X-Content-Type-Options'] = 'nosniff'
    if request.path.startswith('/api/'):
        response.headers['Cache-Control'] = 'no-store'
    if request.method != 'GET' or response.status_code >= 400:
        logger.info('request_finished', extra={'request_id': g.request_id,
                    'task_id': getattr(g, 'task_id', None), 'method': request.method,
                    'path': request.path, 'status': response.status_code,
                    'duration_seconds': round(time.monotonic() - g.started, 3)})
    return response


@app.errorhandler(Exception)
def request_error(error):
    if isinstance(error, HTTPException):
        messages = {413: '文件总大小不能超过 %s MB。' % config.MAX_UPLOAD_MB,
                    404: '任务或文件不存在。', 400: '请求内容不正确，请重试。',
                    405: '请求方式不支持。'}
        return jsonify(error=messages.get(error.code, '请求未完成，请重试。')) , error.code
    logger.exception('request_failed', extra={'request_id': getattr(g, 'request_id', None),
                     'task_id': getattr(g, 'task_id', None)})
    task_id = getattr(g, 'task_id', None)
    if task_id:
        try:
            store.update_task(task_id, status='failed', message='读取文件失败，请重新上传。')
        except Exception:
            logger.exception('upload_status_write_failed')
    return jsonify(error='操作未完成，请重试或联系维护人员。',
                   request_id=getattr(g, 'request_id', None), task_id=task_id), 500


@app.route('/healthz')
def health():
    try:
        store.pending_count()
        healthy = worker.is_alive()
    except Exception:
        logger.exception('health_failed')
        healthy = False
    return jsonify(status='ok' if healthy else 'unavailable'), 200 if healthy else 503


@app.route("/")
def index():
    return render_template("index.html",
                           tasks=store.list_tasks(),
                           template=config.find_template(),
                           queue_size=worker.queue_size())


@app.route("/tasks/<task_id>")
def task_page(task_id):
    task = store.get_task(task_id)
    if not task:
        abort(404)
    options = store.get_options(task)
    return render_template("task.html", task=task,
                           files=store.files_for(task_id),
                           options=options)


@app.route("/api/meta")
def api_meta():
    tpl = config.find_template()
    return jsonify({
        "template_ok": tpl is not None,
        "analyst_fields": config.analyst_fields(),
        "max_upload_mb": config.MAX_UPLOAD_MB,
    })


@app.route("/api/tasks", methods=["POST"])
def api_create_task():
    if not _upload_slot.acquire(blocking=False):
        return jsonify(error='正在读取另一批文件，请稍后再上传。'), 409
    try:
        return create_uploaded_task()
    finally:
        _upload_slot.release()


def create_uploaded_task():
    """上传 PDF -> 落盘 -> 认料 -> 建任务（等用户确认参数）。"""
    uploads = [u for u in request.files.getlist("files") if u and u.filename]
    if not uploads:
        return jsonify({"error": "没有收到文件。"}), 400
    if shutil.disk_usage(config.DATA_DIR).free < (request.content_length or 0) + 100 * 1024 * 1024:
        return jsonify(error='存储空间不足，请联系维护人员。'), 507

    task_id = store.create_task()
    g.task_id = task_id
    store.update_task(task_id, status='uploading')
    logger.info('upload_started', extra={'task_id': task_id, 'request_id': g.request_id})
    workdir = task_dir(task_id)
    reports_dir = os.path.join(workdir, config.REPORTS_DIRNAME)
    os.makedirs(reports_dir, exist_ok=True)
    store.update_task(task_id, workdir=workdir)

    saved, used = [], set()
    for u in uploads:
        name = safe_name(u.filename)
        if not name.lower().endswith(".pdf"):
            continue
        stem, ext = os.path.splitext(name)
        n = 2
        while name in used:
            name = "%s_%d%s" % (stem, n, ext)
            n += 1
        used.add(name)
        path = os.path.join(reports_dir, name)
        header = u.stream.read(1024)
        u.stream.seek(0)
        if b'%PDF-' not in header:
            store.update_task(task_id, status='failed', message='文件不是可识别的 PDF，请重新上传。')
            return jsonify(error='文件不是可识别的 PDF，请重新上传。', task_id=task_id), 400
        u.save(path)
        saved.append(path)

    if not saved:
        store.update_task(task_id, status="failed", message="上传的文件里没有 PDF。")
        return jsonify({"error": "上传的文件里没有 PDF。"}), 400

    try:
        recs = runner.detect_files(saved, workdir)
    except Exception as exc:
        logger.exception('file_detection_failed', extra={'task_id': task_id})
        store.update_task(task_id, status='failed', message='文件识别失败，请检查 PDF 后重新上传。')
        return jsonify(error='文件识别失败，请检查 PDF 后重新上传。', task_id=task_id), 422

    newest_annual = None
    newest_any = None
    for rec in recs:
        store.add_file(task_id, os.path.basename(rec["pdf"]), rec["pdf"],
                       kind=rec.get("kind"), kind_cn=rec.get("kind_cn"),
                       period=rec.get("period"), company=rec.get("company"),
                       code=rec.get("code"),
                       blocking=1 if rec.get("blocking") else 0,
                       why=rec.get("why"))
        if rec.get("blocking") or not rec.get("period"):
            continue
        if newest_any is None or rec["period"] > newest_any["period"]:
            newest_any = rec
        if rec.get("kind") == "annual":
            if newest_annual is None or rec["period"] > newest_annual["period"]:
                newest_annual = rec

    # 没有年报时不拦：季报/半年报照样能填「报告期历史数据」。但要说清楚这次
    # 产出是什么 —— 年度数据整片留空，绝对估值和相对估值都算不出可用结果。
    # 报告期取最近一份报告的期间，不编造一个年末日期：半年报就是半年报。
    interim_only = newest_annual is None
    interim_note = None
    if interim_only:
        kinds = sorted({r.get("kind_cn") or "认不出" for r in recs})
        interim_note = (
            "这批文件里没有年报（只有%s）。可以生成，但「年度数据」会整片留空，"
            "只有季报/半年报进「报告期历史数据」。要让绝对估值和相对估值出结果，"
            "还需要年报。" % ("、".join(kinds) if kinds else "认不出的文件"))

    shown = newest_annual or newest_any or {}
    company_name = shown.get("company")
    company_code = shown.get("code")
    report_period = shown.get("period")

    # 总股本、流通股本、股票代码其实都写在财报里，上传时就替用户算出来。
    # 预填是加分项：它自己挂了也只是留空，不能让整个上传失败。
    try:
        pre = prefill.prefill(recs, workdir, config.cache_dir())
    except Exception:
        logger.exception('prefill_failed', extra={'task_id': task_id})
        pre = {"total_shares_mn": None, "float_shares_mn": None,
               "stock_code": None, "source": {}}
    if not company_code:
        company_code = pre.get("stock_code")

    store.update_task(task_id, company_name=company_name,
                      company_code=company_code, report_period=report_period)
    store.set_options(task_id, {
        "company_name": company_name,
        "company_code": company_code,
        "report_period": report_period,
        "market": {
            "history": {}, "allow_missing_history": True,
            "total_shares_mn": pre.get("total_shares_mn"),
            "float_shares_mn": pre.get("float_shares_mn"),
        },
        "market_source": pre.get("source") or {},
        # 各年财务数（营收/归母净利/净资产/经营现金流），用来把 6 个年末收盘价
        # 换算成历史 PE/PS/PB/PCF。键统一成字符串，存进 JSON 后仍是它。
        "financials": {str(k): v for k, v in (pre.get("financials") or {}).items()},
        "interim_only": interim_only,
        "interim_note": interim_note,
        "analyst": {},
        "allow_incomplete_history": True,
        # 写入后要跑验收：文件「能打开」和数字「对得上」是两回事。
        "skip_accept": False,
    })
    store.update_task(task_id, status='uploaded')
    logger.info('upload_finished', extra={'task_id': task_id})
    return jsonify({"task_id": task_id,
                    "redirect": url_for("task_page", task_id=task_id)})


@app.route("/api/tasks/<task_id>")
def api_task(task_id):
    task = store.get_task(task_id)
    if not task:
        return jsonify({"error": "任务不存在"}), 404
    public = {key: task.get(key) for key in ('id', 'run_id', 'status', 'step', 'total_steps',
              'step_name', 'message', 'company_name', 'report_period', 'xlsm_name')}
    terminal = task['status'] in ('done', 'review', 'failed', 'stopped')
    artifacts = runner.collect_artifacts(task.get('run_dir') or '') if terminal else {}
    public['has_model'] = bool(terminal and artifacts.get('xlsm_path'))
    warning = runner.artifact_warning(artifacts) if public['has_model'] else None
    if public['has_model']:
        public['xlsm_name'] = artifacts['xlsm_name']
        public['download_label'] = '下载中间文件' if artifacts['intermediate'] else '下载模型文件'
    if task['status'] in ('done', 'review') and not public['has_model']:
        public.update(status='failed', message='交付文件不可用，请重新生成。')
    elif public['has_model'] and artifacts['intermediate']:
        public.update(status='stopped', message=warning)
    elif public['has_model'] and task['status'] in ('done', 'review'):
        if warning or artifacts.get('needs_excel_recalc'):
            public.update(status='review', message=warning or task.get('message')
                          or '模型已填充，请用 Excel 打开并复核。')
    elif public['has_model']:
        public['message'] = '生成未完成，已有文件仍可下载查看。'
    public['has_card'] = bool(config.SHOW_REPORT_CARD and terminal and task.get('run_dir') and os.path.isfile(
        os.path.join(task['run_dir'], '请先看这个.html')))
    if task['status'] in ('done', 'review', 'failed', 'stopped'):
        public['summary'] = (runner.summarize(task.get('run_dir'))
                             if config.SHOW_REPORT_CARD and public['has_model'] else {})
        public['summary'].update(
            level='ok' if public['status'] == 'done' else
                  ('warn' if public['status'] == 'review' else 'bad'),
            title=('生成未完成，已有文件可下载' if public['has_model'] and public['status'] in ('failed', 'stopped')
                   else ('模型已生成，待复核' if public['status'] == 'review' else
                         ('模型已生成' if public['has_model'] else '没有生成模型文件'))),
            message=public.get('message'))
    return jsonify(public)


@app.route("/api/tasks/<task_id>/start", methods=["POST"])
def api_start(task_id):
    task = store.get_task(task_id)
    if not task:
        return jsonify({"error": "任务不存在"}), 404
    if task["status"] in ("queued", "running", "uploading"):
        return jsonify({"error": "这个任务已经在队列里了。"}), 409

    body = request.get_json(silent=True) or {}
    if not isinstance(body, dict):
        return jsonify(error='提交内容不正确。'), 400
    if shutil.disk_usage(config.DATA_DIR).free < 100 * 1024 * 1024:
        return jsonify(error='存储空间不足，请联系维护人员。'), 507
    files = store.files_for(task_id)
    if not files or any(item.get('blocking') for item in files):
        return jsonify(error='文件不可用，请重新上传。'), 400
    for key in ('company_name', 'company_code', 'report_period'):
        if body.get(key) is not None and not isinstance(body[key], str):
            return jsonify(error='公司信息格式不正确。'), 400
    for key in ('history', 'prices', 'analyst'):
        if body.get(key) is not None and not isinstance(body[key], dict):
            return jsonify(error='提交内容格式不正确。'), 400
    if body.get('report_period'):
        try:
            date.fromisoformat(body['report_period'])
        except ValueError:
            return jsonify(error='报告期格式应为 YYYY-MM-DD。'), 400
    if body.get('company_code') and not re.fullmatch(r'\d{6}', body['company_code'].strip()):
        return jsonify(error='股票代码应为六位数字。'), 400
    options = store.get_options(task)
    options["company_name"] = (body.get("company_name") or "").strip() or None
    options["company_code"] = (body.get("company_code") or "").strip() or None
    options["report_period"] = (body.get("report_period") or "").strip() or None
    options["allow_incomplete_history"] = bool(
        body.get("allow_incomplete_history", True))

    market = {"history": {}, "allow_missing_history": True}
    for key in ("total_shares_mn", "float_shares_mn", "close_price",
                "hkd_cny", "usd_cny"):
        v = body.get(key)
        if v is None or str(v).strip() == "":
            market[key] = None
        else:
            try:
                market[key] = float(str(v).replace(",", "").strip())
                if not math.isfinite(market[key]) or market[key] <= 0:
                    raise ValueError()
            except ValueError:
                return jsonify({"error": "行情数据不是数字：%s" % key}), 400
    options["market"] = market

    # 历史倍数：办公人员只填 6 个年末收盘价，PE/PS/PB/PCF 由「总市值 ÷ 当年
    # 财务数」算出。显式传 history 时以它为准（便于脚本核对或人工覆盖）。
    history = body.get("history")
    if isinstance(history, dict) and history:
        for year, vals in history.items():
            if not isinstance(vals, dict):
                continue
            clean = {}
            for k, v in vals.items():
                try:
                    clean[k] = float(str(v).replace(",", "").strip())
                    if not math.isfinite(clean[k]):
                        raise ValueError()
                except (TypeError, ValueError):
                    return jsonify({"error": "历史倍数不是数字：%s %s"
                                    % (year, k)}), 400
            if clean:
                market["history"][str(year)] = clean
    else:
        market["history"] = prefill.history_from_prices(
            body.get("prices") or {}, options.get("financials") or {},
            market.get("total_shares_mn"))
        # 没单独填收盘价时，用报告期那一年的年末收盘价 —— 它们本来就是同一个数。
        if market.get("close_price") is None:
            period_year = str(options.get("report_period") or "")[:4]
            price = (body.get("prices") or {}).get(period_year)
            try:
                market["close_price"] = float(str(price).replace(",", "").strip())
            except (TypeError, ValueError):
                pass

    analyst = {}
    for k, v in (body.get("analyst") or {}).items():
        if v is None or str(v).strip() == "":
            continue
        try:
            analyst[k] = float(str(v).replace("%", "").strip())
            if not math.isfinite(analyst[k]):
                raise ValueError()
        except ValueError:
            return jsonify({"error": "假设不是数字：%s" % k}), 400
    options["analyst"] = analyst

    run_id = worker.enqueue(task_id, options)
    if not run_id:
        return jsonify(error='这个任务已经在队列里了。'), 409
    return jsonify(status='queued', run_id=run_id)


@app.route("/api/tasks/<task_id>/card")
def api_card(task_id):
    if not config.SHOW_REPORT_CARD:
        abort(404)
    task = store.get_task(task_id)
    if not task:
        abort(404)
    if task['status'] not in ('done', 'review', 'failed', 'stopped'):
        return Response('此任务没有可交付报告，请返回任务页查看状态。',
                        mimetype='text/plain; charset=utf-8'), 409
    artifacts = runner.collect_artifacts(task.get('run_dir') or '')
    run_dir = task.get("run_dir") or ""
    path = os.path.join(run_dir, "请先看这个.html")
    if not os.path.exists(path):
        return Response("还没有结论卡：任务还没跑完，或还没开始。",
                        mimetype="text/plain; charset=utf-8")
    with io.open(path, encoding="utf-8") as fh:
        content = fh.read()
    if task['status'] != 'done' or artifacts['needs_excel_recalc'] or runner.artifact_warning(artifacts):
        import html
        notice = (task.get('message') if task['status'] in ('failed', 'stopped')
                  else runner.artifact_warning(artifacts)) or '模型已填充，请用 Excel 打开并复核。'
        content = content.replace('<body>', '<body><p role="status" style="padding:12px;background:#fff3cd">'
                                  + html.escape(notice) + '</p>', 1)
    return Response(content, mimetype="text/html; charset=utf-8")


@app.route("/api/tasks/<task_id>/download")
def api_download(task_id):
    task = store.get_task(task_id)
    if not task:
        abort(404)
    artifacts = runner.collect_artifacts(task.get('run_dir') or '')
    path = artifacts['xlsm_path']
    if not path:
        return jsonify({"error": "这个任务没有产出 xlsm。"}), 404
    return send_file(path, as_attachment=True,
                     download_name=artifacts['xlsm_name'])


def selfcheck():
    print("=" * 60)
    print("A股估值模型填充 · 本机 Web 版")
    print("=" * 60)
    tpl = config.find_template()
    print("[OK]   原件模板：%s" % os.path.basename(tpl) if tpl
          else "[缺失] 没找到原件模板，请放进「原件」文件夹")
    for mod, why in (("flask", "网页"), ("pdfplumber", "解析 PDF"),
                     ("openpyxl", "读参数表")):
        try:
            __import__(mod)
            print("[OK]   %s 可用" % mod)
        except Exception:
            print("[缺失] %s 不可用（%s 会失败）" % (mod, why))
    print("服务端不需要 Excel：产出 .xlsm 时公式全保留，缓存值留空，")
    print("办公人员用 Excel 打开时自动重算。")
    print("=" * 60)


def initialize():
    global _initialized
    if _initialized:
        return
    config.ensure_dirs()
    runtime.acquire_instance()
    runtime.configure_logging()
    store.init()
    interrupted = [task for task in store.list_tasks(limit=100000)
                   if task['status'] in ('queued', 'running', 'uploading')]
    runtime.recover_processes([path for task in interrupted
                              for path in (task.get('workdir'), task.get('run_dir'))])
    for task in store.recover_interrupted():
        logger.warning('task_interrupted_on_restart', extra={'task_id': task['id'],
                       'run_id': task.get('run_id')})
    worker.start()
    _initialized = True


if __name__ == "__main__":
    from waitress import serve
    initialize()
    selfcheck()
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "8000"))
    print("打开 http://%s:%d" % ("127.0.0.1" if host == "0.0.0.0" else host, port))
    logger.info('service_started')
    try:
        serve(app, host=host, port=port, threads=4,
              max_request_body_size=config.MAX_UPLOAD_MB * 1024 * 1024)
    finally:
        runtime.stop_processes()
        logger.info('service_stopped')
