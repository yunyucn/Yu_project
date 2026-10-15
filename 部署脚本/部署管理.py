"""在本机操作：默认预览，只有 --执行 才连接服务器。"""
import argparse
import base64
from datetime import datetime
import getpass
import hashlib
import json
import os
from pathlib import Path
import re
import shlex
import sys
import tarfile
import tempfile
import time
import uuid


HERE = Path(__file__).resolve().parent
ACTIONS = ('检查', '部署', '启动', '停止', '状态', '日志', '备份')
EXCLUDED = {'data', '__pycache__', '.git', '.venv', 'venv', 'node_modules',
            '.pytest_cache', '.env', '已信任主机', '部署记录'}
ROOT_FILES = {'app.py', 'config.py', 'prefill.py', 'runner.py', 'runtime.py',
              'store.py', 'worker.py', 'Dockerfile', '.dockerignore',
              'requirements.txt', 'README.md', '启动.sh'}
ROOT_DIRS = {'engine', 'pipeline', 'static', 'templates', '原件'}


def load_config(path):
    settings = json.loads(path.read_text(encoding='utf-8-sig'))
    for name in ('SSH端口', '访问端口'):
        settings[name] = int(settings[name])
        if not 1 <= settings[name] <= 65535:
            raise ValueError('%s 必须在 1—65535 之间。' % name)
    for name in ('SSH连接超时秒', '远程操作超时秒'):
        settings[name] = int(settings[name])
        if settings[name] <= 0:
            raise ValueError('%s 必须为正数。' % name)
    if not re.fullmatch(r'/(opt|srv)/[A-Za-z0-9_-]+', settings['服务器目录']):
        raise ValueError('服务器目录必须为 /opt/独立目录名 或 /srv/独立目录名。')
    if not re.fullmatch(r'[a-z0-9][a-z0-9_-]*', settings['Docker项目名']):
        raise ValueError('Docker 项目名只允许小写字母、数字、横线和下划线。')
    if settings['监听地址'] not in ('0.0.0.0', '127.0.0.1'):
        raise ValueError('监听地址只支持 0.0.0.0 或 127.0.0.1。')
    if not settings['服务器地址'] or not settings['用户名']:
        raise ValueError('服务器地址和用户名不能为空。')
    settings['本地项目'] = (path.parent / settings['项目目录']).resolve()
    return settings


def fingerprint(key):
    return 'SHA256:' + base64.b64encode(hashlib.sha256(key.asbytes()).digest()).decode().rstrip('=')


def connect(settings, private_key):
    try:
        import paramiko
    except ImportError:
        raise RuntimeError('请先安装脚本依赖：python -m pip install -r requirements.txt') from None

    known_hosts = HERE / '已信任主机'
    client = paramiko.SSHClient()
    client.load_system_host_keys()
    if known_hosts.exists():
        client.load_host_keys(str(known_hosts))

    class ConfirmHostKey(paramiko.MissingHostKeyPolicy):
        def missing_host_key(self, ssh_client, hostname, key):
            print('\n首次连接主机：%s\nSSH 指纹：%s' % (hostname, fingerprint(key)))
            print('请通过服务器控制台核对 SSH 指纹；不确定时不要信任。')
            if input('确认无误后输入“信任”：').strip() != '信任':
                raise RuntimeError('未信任服务器指纹，已取消连接。')
            ssh_client.get_host_keys().add(hostname, key.get_name(), key)
            ssh_client.save_host_keys(str(known_hosts))
            if os.name != 'nt':
                known_hosts.chmod(0o600)

    client.set_missing_host_key_policy(ConfirmHostKey())
    arguments = dict(hostname=settings['服务器地址'], port=settings['SSH端口'],
                     username=settings['用户名'], timeout=settings['SSH连接超时秒'],
                     auth_timeout=settings['SSH连接超时秒'],
                     banner_timeout=settings['SSH连接超时秒'],
                     look_for_keys=False, allow_agent=False)
    if private_key:
        key_path = private_key.resolve()
        if not key_path.is_file():
            raise ValueError('指定的 SSH 私钥文件不存在。')
        arguments['key_filename'] = str(key_path)
        arguments['passphrase'] = getpass.getpass('私钥口令（没有则直接回车）：') or None
    else:
        arguments['password'] = getpass.getpass('SSH 登录密码（输入不会显示，不保存）：')
    try:
        client.connect(**arguments)
        client.get_transport().set_keepalive(30)
        return client
    except BaseException:
        client.close()
        raise


def remote(client, command, timeout, transcript):
    stdin, stdout, stderr = client.exec_command(command, timeout=timeout)
    stdin.close()
    channel = stdout.channel
    deadline = time.monotonic() + timeout
    try:
        while True:
            for ready, receive in ((channel.recv_ready, channel.recv),
                                   (channel.recv_stderr_ready, channel.recv_stderr)):
                if ready():
                    data = receive(32768)
                    transcript.write(data)
                    transcript.flush()
                    print(data.decode('utf-8', errors='replace'), end='', flush=True)
            if channel.exit_status_ready() and not channel.recv_ready() and not channel.recv_stderr_ready():
                code = channel.recv_exit_status()
                if code != 0:
                    raise RuntimeError('服务器操作失败（退出码 %s），请查看部署记录。' % code)
                return
            if time.monotonic() > deadline:
                raise TimeoutError('等待超时；远程操作可能仍在继续，请先查看状态和日志，不要重复部署。')
            time.sleep(.05)
    finally:
        stdout.close()
        stderr.close()
        channel.close()


def package_project(source, destination):
    required = ROOT_FILES - {'README.md', '启动.sh'}
    absent = [name for name in required if not (source / name).is_file()]
    if absent or not all((source / name).is_dir() for name in ROOT_DIRS):
        raise ValueError('项目目录不完整，请确认“项目目录”指向 Yu_project 仓库根目录：%s' % absent)
    if not any((source / '原件').glob('*.xlsm')):
        raise ValueError('项目原件目录中缺少 .xlsm 模板。')
    count = 0
    with tarfile.open(destination, 'w:gz') as archive:
        for folder, directories, filenames in os.walk(source, followlinks=False):
            base = Path(folder)
            relative = base.relative_to(source)
            directories[:] = sorted(name for name in directories
                                    if name not in EXCLUDED and not (base / name).is_symlink()
                                    and (relative.parts or name in ROOT_DIRS))
            for name in sorted(filenames):
                path = base / name
                if name in EXCLUDED or path.is_symlink():
                    continue
                if not relative.parts and name not in ROOT_FILES:
                    continue
                if name.endswith(('.pyc', '.log', '.tmp-xlsm', '.key', '.pem')):
                    continue
                archive.add(path, arcname=path.relative_to(source).as_posix(), recursive=False)
                count += 1
    digest = hashlib.sha256()
    with destination.open('rb') as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b''):
            digest.update(chunk)
    print('已打包 %s 个文件，%.1f MB；未打包本机任务、缓存和日志。' %
          (count, destination.stat().st_size / 1024 / 1024))
    return digest.hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('操作', choices=ACTIONS, nargs='?', default='检查')
    parser.add_argument('--配置', type=Path, default=HERE / '部署配置.json')
    parser.add_argument('--密钥', type=Path, help='可选 SSH 私钥；不指定则安全输入登录密码')
    parser.add_argument('--执行', action='store_true', help='实际连接服务器；不加此参数只预览')
    arguments = parser.parse_args()
    settings = load_config(arguments.配置.resolve())
    print('操作：%s\n服务器：%s@%s:%s\n本地项目：%s\n服务器目录：%s\n端口：%s:8000' %
          (arguments.操作, settings['用户名'], settings['服务器地址'], settings['SSH端口'],
           settings['本地项目'], settings['服务器目录'], settings['访问端口']))
    if not arguments.执行:
        print('\n仅预览：没有连接服务器、上传文件或执行 Docker。加 --执行 后才会实际操作。')
        return
    if arguments.操作 in ('部署', '启动') and settings['监听地址'] == '0.0.0.0':
        print('\n注意：应用没有账号登录，HTTP 也不加密。请先在服务器防火墙/安全组中，')
        print('仅允许你们办公电脑的公网 IP 访问端口 %s；脚本不自动修改防火墙。' % settings['访问端口'])
    if input('确认后输入“%s”（其他输入取消）：' % arguments.操作).strip() != arguments.操作:
        print('已取消。')
        return
    identity = datetime.now().strftime('%Y%m%d_%H%M%S') + '_' + uuid.uuid4().hex[:8]
    logs = HERE / '部署记录'
    logs.mkdir(exist_ok=True)
    log_path = logs / ('%s_%s.log' % (identity, arguments.操作))
    with tempfile.TemporaryDirectory(prefix='valuation-deploy-') as temporary:
        digest = None
        archive = Path(temporary) / 'project.tar.gz'
        if arguments.操作 == '部署':
            digest = package_project(settings['本地项目'], archive)
        client = connect(settings, arguments.密钥)
        try:
            with client.open_sftp() as sftp, log_path.open('wb') as transcript:
                sftp.get_channel().settimeout(settings['远程操作超时秒'])
                temporary_remote = '/tmp/valuation-deploy-' + identity
                sftp.mkdir(temporary_remote, mode=0o700)
                helper = temporary_remote + '/server.py'
                upload = temporary_remote + '/project.tar.gz'
                try:
                    sftp.put(str(HERE / '服务器操作.py'), helper)
                    sftp.chmod(helper, 0o600)
                    if digest:
                        sftp.put(str(archive), upload)
                        sftp.chmod(upload, 0o600)
                    command = ['python3', '-u', helper, arguments.操作,
                               '--目录', settings['服务器目录'], '--项目', settings['Docker项目名'],
                               '--端口', str(settings['访问端口']), '--监听', settings['监听地址'],
                               '--版本', identity]
                    if digest:
                        command += ['--压缩包', upload, '--摘要', digest]
                    remote(client, shlex.join(command), settings['远程操作超时秒'], transcript)
                finally:
                    for path in (upload, helper):
                        try:
                            sftp.remove(path)
                        except OSError:
                            pass
                    try:
                        sftp.rmdir(temporary_remote)
                    except OSError:
                        pass
        finally:
            client.close()
            print('\n本机部署记录：%s' % log_path)
    if arguments.操作 in ('部署', '启动'):
        print('服务器健康检查通过。访问地址：http://%s:%s/' %
              (settings['服务器地址'], settings['访问端口']))
        print('外部访问是否可达，还取决于服务器防火墙/安全组。')


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print('\n已中断客户端；远端任务可能仍在运行，请先查看状态。', file=sys.stderr)
        sys.exit(130)
    except Exception as error:
        print('\n操作未完成：%s' % error, file=sys.stderr)
        sys.exit(1)
