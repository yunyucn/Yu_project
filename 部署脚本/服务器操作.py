"""由部署管理.py 临时上传后在 Linux 服务器执行，不需要本地运行。"""
import argparse
import hashlib
import json
import os
from pathlib import Path, PurePosixPath
import re
import shutil
import socket
import subprocess
import sys
import tarfile
import time
import urllib.request


def run(arguments, capture=False, check=True, timeout=1800):
    result = subprocess.run(arguments, text=True, encoding='utf-8', errors='replace',
                            stdout=subprocess.PIPE if capture else None,
                            stderr=subprocess.STDOUT if capture else None, timeout=timeout)
    if check and result.returncode != 0:
        if capture:
            print(result.stdout, flush=True)
        raise RuntimeError('命令失败：%s（退出码 %s）' % (' '.join(arguments), result.returncode))
    return result.stdout.strip() if capture else result.returncode


def prepare_root(root, project, create=False):
    if root.is_symlink() or root.resolve() != root:
        raise RuntimeError('服务器目录不能是符号链接或包含路径跳转。')
    marker = root / '.deployment-owner'
    if marker.exists():
        if marker.read_text(encoding='utf-8').strip() != project:
            raise RuntimeError('此目录属于其他项目，拒绝修改。')
    elif create:
        if root.exists() and any(root.iterdir()):
            raise RuntimeError('目标目录已有其他内容，拒绝覆盖。请更换独立服务器目录。')
        root.mkdir(parents=True, exist_ok=True)
        root.chmod(0o700)
        marker.write_text(project, encoding='utf-8')
    else:
        raise RuntimeError('此项目还没有部署，请先执行部署。')
    for name in ('data', 'releases', 'backups'):
        path = root / name
        if path.is_symlink():
            raise RuntimeError('数据/发布/备份目录不能是符号链接。')
        if create:
            path.mkdir(exist_ok=True)


def compose(root, project, *arguments, release=None, check=True, capture=False):
    directory = release or current_release(root)
    return run(['docker', 'compose', '-p', project, '-f', str(directory / 'compose.json'),
                *arguments], check=check, capture=capture)


def current_release(root):
    marker = root / 'current.txt'
    if not marker.is_file():
        raise RuntimeError('没有当前版本，请先部署。')
    version = marker.read_text(encoding='utf-8').strip()
    if not re.fullmatch(r'[0-9]{8}_[0-9]{6}_[a-f0-9]{8}', version):
        raise RuntimeError('当前版本记录不合法。')
    directory = root / 'releases' / version
    if directory.resolve().parent != (root / 'releases').resolve():
        raise RuntimeError('版本目录越界。')
    if not (directory / 'compose.json').is_file():
        raise RuntimeError('当前版本配置不存在。')
    return directory


def check_port(port, project):
    own_port = False
    identities = run(['docker', 'ps', '-q'], capture=True).splitlines()
    if identities:
        containers = json.loads(run(['docker', 'inspect', *identities], capture=True))
        for container in containers:
            labels = container.get('Config', {}).get('Labels') or {}
            for bindings in (container.get('NetworkSettings', {}).get('Ports') or {}).values():
                for binding in bindings or []:
                    if binding.get('HostPort') == str(port):
                        if labels.get('com.docker.compose.project') != project:
                            raise RuntimeError('端口 %s 已被其他容器占用，不会停止它。' % port)
                        own_port = True
    if not own_port:
        with socket.socket() as probe:
            try:
                probe.bind(('0.0.0.0', port))
            except OSError:
                raise RuntimeError('端口 %s 已被占用，请先人工核实，不会强制释放。' % port) from None


def check_project_owner(root, project):
    identities = run(['docker', 'ps', '-a', '-q', '--filter',
                      'label=com.docker.compose.project=' + project], capture=True).splitlines()
    if not identities:
        return
    containers = json.loads(run(['docker', 'inspect', *identities], capture=True))
    for container in containers:
        labels = container.get('Config', {}).get('Labels') or {}
        configuration = labels.get('com.docker.compose.project.config_files', '')
        paths = [Path(value) for value in configuration.split(',') if value]
        if not paths or any(not path.resolve().is_relative_to(root / 'releases') for path in paths):
            raise RuntimeError('Docker 项目名已被其他目录使用，请更改配置中的 Docker项目名。')


def unpack(archive_path, digest, destination):
    checksum = hashlib.sha256()
    with archive_path.open('rb') as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b''):
            checksum.update(chunk)
    if checksum.hexdigest() != digest:
        raise RuntimeError('上传文件校验失败，不解压、不修改现有服务。')
    destination.mkdir(parents=True, exist_ok=False)
    with tarfile.open(archive_path, 'r:gz') as archive:
        for member in archive.getmembers():
            relative = PurePosixPath(member.name)
            if (relative.is_absolute() or '..' in relative.parts or '\\' in member.name
                    or not relative.parts or relative.parts[0] == 'data'
                    or member.issym() or member.islnk() or not (member.isfile() or member.isdir())):
                raise RuntimeError('压缩包包含非法路径或文件类型。')
        for member in archive.getmembers():
            target = destination.joinpath(*PurePosixPath(member.name).parts)
            if member.isdir():
                target.mkdir(parents=True, exist_ok=True)
                continue
            target.parent.mkdir(parents=True, exist_ok=True)
            with archive.extractfile(member) as source, target.open('wb') as output:
                shutil.copyfileobj(source, output)
            target.chmod(0o644)


def write_compose(root, project, release, port, bind):
    value = {'services': {'valuation': {
        'build': {'context': str(release)},
        'image': '%s:%s' % (project, release.name),
        'ports': ['%s:%s:8000' % (bind, port)],
        'volumes': [{'type': 'bind', 'source': str(root / 'data'), 'target': '/app/data'}],
        'environment': {'HOST': '0.0.0.0', 'PORT': '8000', 'TZ': 'Asia/Shanghai'},
        'restart': 'unless-stopped', 'init': True, 'stop_grace_period': '45s',
        'healthcheck': {'test': ['CMD', 'python', '-c',
            "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8000/healthz', timeout=5)"],
            'interval': '15s', 'timeout': '10s', 'retries': 4, 'start_period': '30s'},
        'logging': {'driver': 'json-file', 'options': {'max-size': '10m', 'max-file': '3'}}}}}
    (release / 'compose.json').write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')


def await_health(root, project, port):
    for attempt in range(60):
        try:
            with urllib.request.urlopen('http://127.0.0.1:%s/healthz' % port, timeout=3) as response:
                if response.status == 200 and json.load(response).get('status') == 'ok':
                    print('健康检查通过。', flush=True)
                    return
        except Exception:
            pass
        time.sleep(2)
    compose(root, project, 'ps', check=False)
    compose(root, project, 'logs', '--tail', '100', '--no-color', check=False)
    raise RuntimeError('健康检查未通过。已保留文件和日志，请排查；不会删除数据或自动回退数据库。')


def backup_data(root, label):
    backup = root / 'backups' / ('data_%s.tar.gz' % label)
    with tarfile.open(backup, 'x:gz') as archive:
        archive.add(root / 'data', arcname='data')
    backup.chmod(0o600)
    print('数据备份：%s' % backup, flush=True)
    return backup


def deploy(arguments, root):
    project = arguments.项目
    previous = current_release(root) if (root / 'current.txt').is_file() else None
    check_port(arguments.端口, project)
    release = root / 'releases' / arguments.版本
    unpack(Path(arguments.压缩包), arguments.摘要, release)
    write_compose(root, project, release, arguments.端口, arguments.监听)
    compose(root, project, 'config', '--quiet', release=release)
    print('构建新版本；此阶段不停止旧服务。', flush=True)
    compose(root, project, 'build', release=release)
    was_running = bool(previous and compose(root, project, 'ps', '--status', 'running', '-q', capture=True))
    if previous:
        compose(root, project, 'stop', release=previous)
    try:
        if any((root / 'data').iterdir()):
            backup_data(root, arguments.版本)
    except BaseException:
        if was_running:
            compose(root, project, 'up', '-d', release=previous, check=False)
        raise
    if previous:
        (root / 'previous.txt').write_text(previous.name, encoding='utf-8')
    pending = root / 'current.next'
    pending.write_text(release.name, encoding='utf-8')
    pending.replace(root / 'current.txt')
    compose(root, project, 'up', '-d', '--no-build')
    await_health(root, project, arguments.端口)
    (root / 'last-success.txt').write_text(release.name, encoding='utf-8')
    print('部署完成；旧版本和备份仍保留，不自动清理其他镜像或容器。', flush=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('操作', choices=('检查', '部署', '启动', '停止', '状态', '日志', '备份'))
    parser.add_argument('--目录', required=True)
    parser.add_argument('--项目', required=True)
    parser.add_argument('--端口', type=int, required=True)
    parser.add_argument('--监听', choices=('127.0.0.1', '0.0.0.0'), required=True)
    parser.add_argument('--版本', required=True)
    parser.add_argument('--压缩包')
    parser.add_argument('--摘要')
    arguments = parser.parse_args()
    if not sys.platform.startswith('linux'):
        raise RuntimeError('服务器端脚本只支持 Linux。')
    if sys.version_info < (3, 9):
        raise RuntimeError('服务器需要 Python 3.9 或以上。')
    if not re.fullmatch(r'/(opt|srv)/[A-Za-z0-9_-]+', arguments.目录):
        raise ValueError('服务器目录不合法。')
    if not re.fullmatch(r'[a-z0-9][a-z0-9_-]*', arguments.项目):
        raise ValueError('项目名不合法。')
    if not re.fullmatch(r'[0-9]{8}_[0-9]{6}_[a-f0-9]{8}', arguments.版本):
        raise ValueError('版本编号不合法。')
    if not 1 <= arguments.端口 <= 65535:
        raise ValueError('端口不合法。')
    root = Path(arguments.目录)
    for command in (['docker', 'version'], ['docker', 'compose', 'version']):
        run(command, timeout=30)
    check_project_owner(root, arguments.项目)
    print('Python：%s\n服务器目录：%s\n操作：%s' % (sys.version.split()[0], root, arguments.操作), flush=True)
    if arguments.操作 == '检查':
        check_port(arguments.端口, arguments.项目)
        print('端口检查通过。此检查不安装软件、不修改防火墙。', flush=True)
        if root.exists():
            prepare_root(root, arguments.项目)
        return
    prepare_root(root, arguments.项目, create=arguments.操作 == '部署')
    import fcntl
    with (root / '.operation.lock').open('a') as lock:
        try:
            fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            raise RuntimeError('本项目另一个操作正在运行，请等待完成。') from None
        if arguments.操作 == '部署':
            if not arguments.压缩包 or not arguments.摘要:
                raise ValueError('部署必须提供压缩包和校验摘要。')
            deploy(arguments, root)
        elif arguments.操作 == '启动':
            check_port(arguments.端口, arguments.项目)
            compose(root, arguments.项目, 'up', '-d', '--no-build')
            await_health(root, arguments.项目, arguments.端口)
        elif arguments.操作 == '停止':
            compose(root, arguments.项目, 'stop')
            print('仅停止本项目，数据库、上传文件和结果均保留。')
        elif arguments.操作 == '状态':
            compose(root, arguments.项目, 'ps', '-a')
        elif arguments.操作 == '日志':
            compose(root, arguments.项目, 'logs', '--tail', '150', '--no-color')
        elif arguments.操作 == '备份':
            was_running = bool(compose(root, arguments.项目, 'ps', '--status', 'running', '-q', capture=True))
            compose(root, arguments.项目, 'stop')
            try:
                backup_data(root, arguments.版本)
            finally:
                if was_running:
                    compose(root, arguments.项目, 'up', '-d', '--no-build')
                    await_health(root, arguments.项目, arguments.端口)


if __name__ == '__main__':
    try:
        main()
    except Exception as error:
        print('服务器操作失败：%s' % error, file=sys.stderr, flush=True)
        sys.exit(1)
