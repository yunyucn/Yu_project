"""Web 应用的路径与常量。

所有路径都相对本文件解析，所以整个项目目录可以整体搬到别的机器（Windows
或 Linux），不需要改任何绝对路径。流水线、模板都随项目一起走。

服务端不依赖 Excel、不依赖 Windows：写工作簿由 pipeline/scripts/com_excel.py
（纯 Python，直接改 xlsm 里的 XML）完成，公式的计算结果交给办公人员用 Excel
打开时算（文件里已写 fullCalcOnLoad=1）。
"""
import io
import json
import os
import re

WEB_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(WEB_DIR)

# 流水线就在本项目里，不是旁边另一个 skill 目录：部署时只搬一个目录。
SKILL_DIR = os.path.join(WEB_DIR, "pipeline")
SCRIPTS_DIR = os.path.join(SKILL_DIR, "scripts")
CONFIG_DIR = os.path.join(SKILL_DIR, "config")
# 原件模板随项目发布。放在项目自己的「原件」下，服务端就不需要再有人
# 手工把模板拷到某个约定的位置。
TEMPLATE_DIR = os.path.join(WEB_DIR, "原件")

PIPELINE = os.path.join(SCRIPTS_DIR, "run_pipeline.py")
DETECT = os.path.join(SCRIPTS_DIR, "detect.py")
PARSE_PDF = os.path.join(SCRIPTS_DIR, "parse_pdf.py")
MAPPING = os.path.join(CONFIG_DIR, "mapping.json")
ANALYST_INPUTS = os.path.join(CONFIG_DIR, "analyst_inputs.json")
PDF_PAGES = os.path.join(CONFIG_DIR, "pdf_pages.json")
VERSION_FILE = os.path.join(SKILL_DIR, "VERSION")

DATA_DIR = os.path.join(WEB_DIR, "data")
TASKS_DIR = os.path.join(DATA_DIR, "tasks")
DB_PATH = os.path.join(DATA_DIR, "webapp.db")

REPORTS_DIRNAME = "财报"
# 结果落在「估值结果」下的「公司_期间_v版本_时间戳」文件夹里，办公人员
# 一眼能认出哪份是自己的。
OUT_DIRNAME = "估值结果"

TOTAL_STEPS = 11
SHOW_REPORT_CARD = False
MAX_UPLOAD_MB = 300
LOG_DIR = os.path.join(DATA_DIR, 'logs')
PIPELINE_TIMEOUT_SECONDS = int(os.environ.get('PIPELINE_TIMEOUT_SECONDS', '1800'))
PREFILL_TIMEOUT_SECONDS = int(os.environ.get('PREFILL_TIMEOUT_SECONDS', '900'))
DETECT_TIMEOUT_SECONDS = int(os.environ.get('DETECT_TIMEOUT_SECONDS', '120'))


def ensure_dirs():
    for d in (DATA_DIR, TASKS_DIR):
        os.makedirs(d, exist_ok=True)


def skill_version():
    """skill 自己的版本号，用于产出文件夹命名。读不到就返回 "0" ——
    命名只是标签，绝不能成为一次运行拒绝开始的理由。"""
    try:
        with open(VERSION_FILE, encoding="utf-8") as fh:
            return fh.read().strip() or "0"
    except OSError:
        return "0"


def cache_dir():
    """和 run_pipeline.default_cache_dir() 完全一致。

    预填和正式流水线必须共用同一份解析缓存，否则同一本 265 页的年报会被解析
    两遍，而解析占了整次运行 90% 的时间。

    路径必须跨平台：Windows 有 %LOCALAPPDATA%，Linux 没有，而服务端要能
    部署到 Linux。缺省落到项目自己的 data/cache，容器里挂一个卷就够。
    """
    local = os.environ.get("LOCALAPPDATA")
    if local and os.path.isdir(local):
        return os.path.join(local, "a-share-valuation-fill", "cache")
    base = os.environ.get("XDG_CACHE_HOME")
    if base:
        return os.path.join(base, "a-share-valuation-fill", "cache")
    return os.path.join(DATA_DIR, "cache")


def _sha256(path):
    import hashlib
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def pristine_sha():
    """mapping.json 校准过的那份模板的 sha256。"""
    m = json.load(io.open(MAPPING, encoding="utf-8"))
    return m["template"]["sha256"]


def find_template():
    """按 sha256 找模板，不按文件名。

    原件旁边常放着「已解锁」「副本」之类的改写版，按名字挑会挑错，而挑错模板
    会让整份产出建立在一个被动过手脚的引擎上 —— 字典以行号为键，会静默地把
    每个值写到错行。
    """
    want = pristine_sha()
    roots = [
        TEMPLATE_DIR,                       # 项目自带的那份
        os.path.join(PROJECT_ROOT, "原件"),
        os.path.join(WEB_DIR, "模板"),
        os.path.join(SKILL_DIR, "模板"),
        SKILL_DIR,
        PROJECT_ROOT,
    ]
    seen = set()
    for root in roots:
        if not os.path.isdir(root):
            continue
        for f in sorted(os.listdir(root)):
            if not f.lower().endswith(".xlsm"):
                continue
            full = os.path.join(root, f)
            if full in seen:
                continue
            seen.add(full)
            try:
                if _sha256(full) == want:
                    return full
            except OSError:
                continue
    return None


def analyst_fields():
    """模板里六个「只能由人定」的判断格，作为表单字段。

    从 analyst_inputs.json 读，不写死：标签里的年份跟着报告期走
    （今年是 2026E，明年是 2027E），写死会指向一个不存在的格。
    """
    c = json.load(io.open(ANALYST_INPUTS, encoding="utf-8"))
    out = []
    for t in c.get("true_inputs", []):
        label, year = t.get("label"), t.get("year")
        key = "%s(%s)" % (label, year) if year else label
        out.append({"key": key, "label": label, "year": year,
                    "cell": t.get("cell", ""), "if_blank": t.get("if_blank", "")})
    return out
