"""应用日志、单实例锁与有超时保护的子进程。"""
import atexit
from collections import deque
from datetime import datetime, timezone
import json
import logging
from logging.handlers import RotatingFileHandler
import os
import subprocess
import threading
import time

import psutil

import config

logger = logging.getLogger(__name__)
_processes = set()
_lock = threading.Lock()
_instance_file = None


class JsonFormatter(logging.Formatter):
    def format(self, record):
        data = {'time': datetime.now(timezone.utc).isoformat(),
                'level': record.levelname, 'logger': record.name,
                'event': record.getMessage()}
        for key in ('task_id', 'run_id', 'request_id', 'stage', 'stage_name', 'status',
                    'duration_seconds', 'exit_code', 'method', 'path', 'log_path'):
            if hasattr(record, key):
                data[key] = getattr(record, key)
        if record.exc_info:
            data['exception'] = self.formatException(record.exc_info)
        return json.dumps(data, ensure_ascii=False)


def configure_logging():
    root = logging.getLogger()
    if any(getattr(handler, '_valuation_log', False) for handler in root.handlers):
        return
    os.makedirs(config.LOG_DIR, exist_ok=True)
    formatter = JsonFormatter()
    handler = RotatingFileHandler(os.path.join(config.LOG_DIR, 'app.jsonl'),
                                  maxBytes=5 * 1024 * 1024, backupCount=5, encoding='utf-8')
    handler._valuation_log = True
    handler.setFormatter(formatter)
    root.addHandler(handler)
    console = logging.StreamHandler()
    console.setFormatter(formatter)
    root.addHandler(console)
    root.setLevel(logging.INFO)


def acquire_instance():
    global _instance_file
    if _instance_file is not None:
        return
    config.ensure_dirs()
    handle = open(os.path.join(config.DATA_DIR, 'service.lock'), 'a+b')
    try:
        handle.seek(0)
        if not handle.read(1):
            handle.write(b'0')
            handle.flush()
        handle.seek(0)
        if os.name == 'nt':
            import msvcrt
            msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(handle, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError:
        handle.close()
        raise RuntimeError('已有服务使用此数据目录，请勿重复启动。') from None
    _instance_file = handle


def terminate_tree(process):
    try:
        parent = psutil.Process(process.pid)
        children = parent.children(recursive=True)
        for child in reversed(children):
            try:
                child.kill()
            except psutil.NoSuchProcess:
                pass
        parent.kill()
        psutil.wait_procs(children, timeout=3)
    except psutil.NoSuchProcess:
        pass


def stop_processes():
    with _lock:
        processes = list(_processes)
    for process in processes:
        try:
            terminate_tree(process)
        except Exception:
            logger.exception('process_shutdown_failed')


atexit.register(stop_processes)


def recover_processes(directories):
    for directory in set(path for path in directories if path):
        if not os.path.isdir(directory):
            continue
        for filename in os.listdir(directory):
            if not filename.endswith('.process.json'):
                continue
            path = os.path.join(directory, filename)
            try:
                with open(path, encoding='utf-8') as handle:
                    record = json.load(handle)
                process = psutil.Process(record['pid'])
                if (process.create_time() == record['created_at']
                        and process.cmdline() == record['command']):
                    terminate_tree(process)
                    logger.warning('interrupted_process_stopped', extra={'log_path': path})
            except psutil.NoSuchProcess:
                pass
            except Exception:
                logger.exception('interrupted_process_recovery_failed', extra={'log_path': path})
                raise
            os.remove(path)


def run_command(command, log_path, timeout_seconds, on_line=None, context=None):
    started = time.monotonic()
    tail = deque(maxlen=200)
    timed_out = threading.Event()
    os.makedirs(os.path.dirname(os.path.abspath(log_path)), exist_ok=True)
    environment = dict(os.environ, PYTHONIOENCODING='utf-8', PYTHONUNBUFFERED='1')
    context = dict(context or {}, log_path=str(log_path))
    logger.info('process_started', extra=context)
    with open(log_path, 'a', encoding='utf-8', buffering=1) as logfile:
        process = subprocess.Popen(command, stdout=subprocess.PIPE,
                                   stderr=subprocess.STDOUT, env=environment)
        with _lock:
            _processes.add(process)
        marker = log_path + '.process.json'

        def expire():
            timed_out.set()
            try:
                terminate_tree(process)
            except Exception:
                logger.exception('process_timeout_cleanup_failed', extra=context)

        timer = threading.Timer(timeout_seconds, expire)
        timer.daemon = True
        timer.start()
        try:
            try:
                tracked = psutil.Process(process.pid)
                with open(marker, 'w', encoding='utf-8') as handle:
                    json.dump({'pid': process.pid, 'created_at': tracked.create_time(),
                               'command': list(command)}, handle)
            except psutil.NoSuchProcess:
                pass
            for raw in process.stdout:
                line = raw.decode('utf-8', 'replace')
                logfile.write('%s %s' % (datetime.now(timezone.utc).isoformat(), line))
                logfile.flush()
                tail.append(line.rstrip('\r\n'))
                if on_line:
                    on_line(line)
            process.wait()
            if timed_out.is_set():
                logfile.write('\nTIMEOUT after %s seconds\n' % timeout_seconds)
                raise TimeoutError('子进程超过 %s 秒：%s' % (timeout_seconds, log_path))
        except BaseException:
            terminate_tree(process)
            process.wait()
            raise
        finally:
            timer.cancel()
            process.stdout.close()
            with _lock:
                _processes.discard(process)
            if os.path.exists(marker):
                os.remove(marker)
            logger.info('process_finished', extra=dict(context, exit_code=process.poll(),
                        duration_seconds=round(time.monotonic() - started, 3)))
    return process.returncode, '\n'.join(tail)
