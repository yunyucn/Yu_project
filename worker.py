"""单进程串行生成队列，结果与运行记录分开保存。"""
import io
import json
import logging
import os
import queue
import threading
import time
import traceback

import config
import runner
import store

_q = queue.Queue()
_thread = None
_lock = threading.Lock()
logger = logging.getLogger(__name__)


def start():
    global _thread
    with _lock:
        if _thread and _thread.is_alive():
            return
        _thread = threading.Thread(target=_loop, name='pipeline-worker', daemon=True)
        _thread.start()


def is_alive():
    return bool(_thread and _thread.is_alive())


def enqueue(task_id, options):
    with _lock:
        run_id = store.claim_task(task_id, options)
        if run_id:
            _q.put(task_id)
            logger.info('task_queued', extra={'task_id': task_id, 'run_id': run_id})
        return run_id


def queue_size():
    return _q.qsize()


def _loop():
    while True:
        task_id = _q.get()
        started = time.monotonic()
        task = None
        try:
            task = store.get_task(task_id)
            if task and task['status'] == 'queued':
                _run(task_id)
        except Exception as exc:
            logger.exception('task_failed', extra={'task_id': task_id,
                             'run_id': (task or {}).get('run_id')})
            message = ('生成超时，请重试或联系维护人员。' if isinstance(exc, TimeoutError)
                       else '生成失败，请重试或将任务编号发给维护人员。')
            try:
                store.update_task(task_id, status='failed', message=message,
                                  step_name='已停止', xlsm_path=None, xlsm_name=None,
                                  log_tail=traceback.format_exc()[-4000:])
            except Exception:
                logger.exception('task_status_write_failed', extra={'task_id': task_id})
        finally:
            try:
                current = store.get_task(task_id)
                if current and task and current.get('run_id') == task.get('run_id'):
                    store.update_run(task['run_id'], status=current['status'],
                                     run_dir=current.get('run_dir'), message=current.get('message'),
                                     finished_at=store._now(),
                                     duration_seconds=round(time.monotonic() - started, 3))
                    logger.info('task_finished', extra={'task_id': task_id,
                                'run_id': current['run_id'], 'status': current['status'],
                                'duration_seconds': round(time.monotonic() - started, 3)})
            except Exception:
                logger.exception('run_record_write_failed', extra={'task_id': task_id})
            _q.task_done()


def _run(task_id):
    task = store.get_task(task_id)
    if not task:
        return
    files = store.files_for(task_id)
    options = store.get_options(task)
    context = {'task_id': task_id, 'run_id': task.get('run_id')}
    pdfs = [item['stored_path'] for item in files
            if item.get('stored_path') and os.path.isfile(item['stored_path'])]
    if not pdfs or any(item.get('blocking') for item in files):
        store.update_task(task_id, status='stopped', message='文件不可用，请重新上传。')
        return
    company = {key: options.get(key) or task.get(key)
               for key in ('company_name', 'company_code', 'report_period')}
    company = {'name': company['company_name'], 'code': company['company_code'],
               'report_period': company['report_period']}
    stem = runner.short_company(company['name']) or '估值模型'
    label = runner.period_label(company['report_period'])
    name = runner.safe_component('_'.join(value for value in (stem, label, '估值模型') if value))
    run_dir = runner.result_dir(task['workdir'], stem, label)
    os.makedirs(run_dir, exist_ok=False)
    store.update_task(task_id, run_dir=run_dir, status='running', step=0,
                      step_name='启动中', total_steps=config.TOTAL_STEPS)
    store.update_run(task.get('run_id'), status='running', run_dir=run_dir)
    logger.info('task_started', extra=context)
    params = runner.build_params(company, options.get('market'), options.get('analyst'), pdfs, options)
    params['_web_task_id'] = task_id
    params['_web_run_id'] = task.get('run_id')
    params_path = os.path.join(run_dir, 'params.json')
    with io.open(params_path, 'w', encoding='utf-8') as handle:
        json.dump(params, handle, ensure_ascii=False, indent=1)
    template = config.find_template()
    if not template:
        raise RuntimeError('原件模板不存在或哈希不匹配')
    stage_started = time.monotonic()
    previous_step = None

    def on_progress(step, total, step_name):
        nonlocal stage_started, previous_step
        now = time.monotonic()
        if previous_step is not None:
            logger.info('stage_finished', extra=dict(context, stage=previous_step,
                        duration_seconds=round(now - stage_started, 3)))
        previous_step, stage_started = step, now
        logger.info('stage_started', extra=dict(context, stage=step, stage_name=step_name))
        store.update_task(task_id, step=step, total_steps=total, step_name=step_name)

    returncode, tail = runner.run_pipeline(
        params_path, template, run_dir, name, on_progress=on_progress,
        log_path=os.path.join(run_dir, '_web_run.log'), cache_dir=config.cache_dir(),
        allow_no_annual=bool(options.get('interim_only')))
    logger.info('pipeline_exited', extra=dict(context, exit_code=returncode,
                stage=previous_step, duration_seconds=round(time.monotonic() - stage_started, 3)))
    store.update_run(task.get('run_id'), exit_code=returncode)
    artifacts = runner.collect_artifacts(run_dir)
    has_output = bool(artifacts['xlsm_path'])
    fields = dict(log_tail=tail[-4000:], verdict=artifacts['verdict'],
                  xlsm_path=artifacts['xlsm_path'], xlsm_name=artifacts['xlsm_name'],
                  step_name='已停止')
    if artifacts.get('xlsm_path'):
        try:
            fields['run_output_sha256'] = config._sha256(artifacts['xlsm_path'])
        except OSError:
            fields['run_output_sha256'] = None
    if has_output and returncode in (0, 2) and not artifacts['intermediate']:
        review = (returncode != 0 or artifacts['passed'] is not True
                  or artifacts['accepted'] is not True
                  or runner.artifact_warning(artifacts)
                  or artifacts['needs_excel_recalc'] or artifacts['verdict'] != 'full')
        fields.update(status='review' if review else 'done', step=config.TOTAL_STEPS,
                      step_name='待复核' if review else '完成',
                      message=('模型文件已生成，请用 Excel 打开并复核。' if review
                               else '模型已生成。'))
        if returncode == 2 or artifacts['passed'] is not True or artifacts['accepted'] is not True:
            fields['message'] = '模型文件已生成，但部分检查未通过，仍可下载复核。'
        fields['message'] = runner.artifact_warning(artifacts) or fields['message']
    else:
        fields.update(status='stopped' if returncode == 1 else 'failed',
                      message='生成未完成或校验未通过，请重试或联系维护人员。')
        if has_output:
            fields['message'] = ('生成未完成，已有文件仍可下载查看。' if not artifacts['intermediate']
                                 else runner.artifact_warning(artifacts))
    with open(os.path.join(run_dir, 'audit.json'), 'w', encoding='utf-8') as handle:
        json.dump(dict(context, exit_code=returncode, status=fields['status'],
                       template_sha256=config._sha256(template),
                       params_sha256=config._sha256(params_path),
                       output_sha256=config._sha256(fields['xlsm_path']) if fields['xlsm_path'] else None,
                       readback_passed=artifacts['passed'], acceptance_passed=artifacts['accepted'],
                       needs_excel_recalc=artifacts['needs_excel_recalc']),
                  handle, ensure_ascii=False, indent=2)
    output_sha256 = fields.pop('run_output_sha256', None)
    store.update_task(task_id, **fields)
    store.update_run(task.get('run_id'), status=fields['status'], message=fields['message'],
                     finished_at=store._now(), output_sha256=output_sha256,
                     params_json=json.dumps(params, ensure_ascii=False))
