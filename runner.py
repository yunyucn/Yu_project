"""把一次 Web 任务翻译成 run_pipeline.py 的一次调用。

这一层不碰算法：它只负责
  1. 用 skill 自己的 detect.py 认清这批 PDF 是什么；
  2. 组装 params.json（缺的字段留 null，绝不变成 0）；
  3. 起子进程跑流水线，把 stdout 里的「第 N/11 步」翻成进度；
  4. 跑完后把产物（xlsm、结论卡、判定）找出来。

流水线的中文输出在 Windows 控制台会是乱码，所以子进程一律以 UTF-8 通信
（PYTHONIOENCODING=utf-8），这里按 utf-8 解码。
"""
import io
import html
import json
import os
import re
import subprocess
import sys
import time
import zipfile

import runtime

import config

STEP_RE = re.compile(r"第\s*(\d+)\s*/\s*(\d+)\s*步[：:]\s*(.*)")
PAREN_TAIL = re.compile(r"[（(][^）)]*[）)]\s*$")
STOP_RE = re.compile(r"^[^\n]*?(?:停下|停止|拒绝|中止)[^\n]*", re.M)

PERIOD_LABEL = {"03-31": "一季报", "06-30": "半年报",
                "09-30": "三季报", "12-31": "年报"}

CORP_FORM = re.compile(r"(股份有限公司|有限责任公司|有限公司|公司)$")
CORP_DESCRIPTOR = re.compile(r"(集团|控股|科技|股份)$")


def short_company(company):
    """兴源环境科技股份有限公司 -> 兴源环境。"""
    name = (company or "").strip()
    while True:
        stripped = CORP_FORM.sub("", CORP_DESCRIPTOR.sub("", CORP_FORM.sub("", name)))
        stripped = stripped.strip()
        if stripped == name or not stripped:
            break
        name = stripped
    return name


def period_label(period):
    """2025-12-31 -> 2025年报。"""
    if not period:
        return None
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", str(period).strip())
    if not m:
        return str(period).strip()
    return m.group(1) + PERIOD_LABEL.get(m.group(2) + "-" + m.group(3),
                                         m.group(2) + "-" + m.group(3))


def result_dir(workdir, stem, label):
    """按 skill 的约定落产出：<workdir>/估值结果/<公司>_<期间>_v<版本>_<时间戳>/。

    同名同分钟会撞车，撞了就加 -2、-3，绝不复用文件夹 —— 复用会把两次运行的
    产物混在一起，而打开文件夹的人分不出哪个数是哪一次的。
    """
    run = safe_component("_".join(x for x in (stem, label, "v" + config.skill_version()) if x))
    stamp = time.strftime("%Y-%m-%d_%H%M")
    root = os.path.join(workdir, config.OUT_DIRNAME)
    path = os.path.join(root, "%s_%s" % (run, stamp))
    n = 2
    while os.path.exists(path):
        path = os.path.join(root, "%s_%s-%d" % (run, stamp, n))
        n += 1
    return path


def safe_component(value):
    value = re.sub(r'[\\/:*?"<>|\x00-\x1f]', '_', str(value or '')).strip(' .')
    return value[:120] or '估值模型'


def _py():
    return sys.executable


def _env():
    return dict(os.environ, PYTHONIOENCODING="utf-8")


def detect_files(pdf_paths, workdir):
    """认清每份 PDF：公司、报告类型、报告期。

    这是 skill 的 detect.py，不是我们的复刻：报告期必须由封面决定，
    不能让调用方猜 —— 一份一季报被当成年度报告，三个月的数字会写进年度列，
    整个估值建立在三个月的数据上，而工作簿不报任何错。
    """
    os.makedirs(workdir, exist_ok=True)
    params_path = os.path.join(workdir, "_detect_params.json")
    out_path = os.path.join(workdir, "_detect_out.json")
    report_path = os.path.join(workdir, "_detect_report.md")
    params = {"reports": [{"pdf": p} for p in pdf_paths]}
    with io.open(params_path, "w", encoding="utf-8") as fh:
        json.dump(params, fh, ensure_ascii=False, indent=1)
    cmd = [_py(), config.DETECT, "--params", params_path,
           "--out", out_path, "--out-report", report_path]
    returncode, tail = runtime.run_command(
        cmd, os.path.join(workdir, '_detect.log'), config.DETECT_TIMEOUT_SECONDS,
        context={'task_id': os.path.basename(workdir)})
    if returncode != 0 or not os.path.exists(out_path):
        raise RuntimeError("认料失败：%s" % tail[-800:])
    with io.open(out_path, encoding="utf-8") as fh:
        data = json.load(fh)
    blocked = {b["pdf"]: b["why"] for b in data.get("blocking", [])
               if b.get("pdf")}
    global_blocking = [b.get("why") for b in data.get("blocking", [])
                       if not b.get("pdf") and b.get("why")]
    out = []
    for r in data.get("reports", []):
        rec = dict(r)
        rec["blocking"] = bool(global_blocking) or r["pdf"] in blocked
        reasons = ([blocked[r["pdf"]]] if r["pdf"] in blocked else [])
        reasons.extend(global_blocking)
        rec["why"] = "；".join(reasons) if reasons else None
        out.append(rec)
    return out


def build_params(company, market, analyst, pdf_paths, options):
    """组装 params.json。

    空白保持 null —— 模板把空白读成 0 且不报错，所以「没填」和「填了 0」
    必须在整条链路上保持可区分。
    """
    market = dict(market or {})
    market.setdefault("history", {})
    market.setdefault("allow_missing_history", True)
    return {
        "company": {
            "name": company.get("name") or None,
            "code": company.get("code") or None,
            "report_period": company.get("report_period") or None,
            "valuation_date": company.get("report_period") or None,
        },
        "market": market,
        "reports": [{"pdf": p} for p in pdf_paths],
        "analyst": {k: v for k, v in (analyst or {}).items() if v is not None},
        "allow_incomplete_history": bool(options.get("allow_incomplete_history", True)),
        "_comment": "由 web项目 生成；改参数请回到页面，不要手改这个文件。",
    }


def parse_step(line):
    """第 3/11 步：读取年报 PDF：这一步最慢，请等待（1-3 分钟） -> (3, 11, 名称)"""
    m = STEP_RE.search(line or "")
    if not m:
        return None
    name = PAREN_TAIL.sub("", m.group(3).strip()).strip()
    return int(m.group(1)), int(m.group(2)), name


def run_pipeline(params_path, template, workdir, name, on_progress=None,
                 skip_accept=False, log_path=None, cache_dir=None,
                 allow_no_annual=False):
    """跑流水线，逐行解析进度。返回 (returncode, log 尾部)。"""
    os.makedirs(workdir, exist_ok=True)
    cmd = [_py(), config.PIPELINE,
           "--params", params_path,
           "--template", template,
           "--workdir", workdir,
           "--name", name]
    if cache_dir:
        # 显式给缓存目录：预填解析过的 PDF 在这里被复用，不重复解析。
        cmd += ["--cache-dir", cache_dir]
    if allow_no_annual:
        # 只有季报/半年报时也继续：年度数据留空，季度表照填。
        cmd.append("--allow-no-annual")
    if skip_accept:
        cmd.append("--no-accept")
    def handle_line(line):
        step = parse_step(line)
        if step and on_progress:
            on_progress(step[0], step[1], step[2])

    params = _load_json(params_path) or {}
    return runtime.run_command(cmd, log_path or os.path.join(workdir, '_web_run.log'),
                               config.PIPELINE_TIMEOUT_SECONDS, handle_line,
                               {'task_id': params.get('_web_task_id'),
                                'run_id': params.get('_web_run_id')})


def collect_artifacts(workdir):
    """查找本次文件；校验只提供提示，不作为下载门槛。"""
    out = {"xlsm_path": None, "xlsm_name": None, "card_html": None,
           "verdict": None, "passed": None, "raw_dir": None,
           "accepted": None, "needs_excel_recalc": True,
           "fingerprint_matches": False, "package_valid": False, "intermediate": False}
    if not os.path.isdir(workdir):
        return out
    card = os.path.join(workdir, "请先看这个.html")
    if os.path.exists(card):
        out["card_html"] = card
    raw = os.path.join(workdir, "_AI参考_可删")
    readback = {}
    acceptance = {}
    if os.path.isdir(raw):
        out["raw_dir"] = raw
        compat = _load_json(os.path.join(raw, 'compat.json')) or {}
        readback = _load_json(os.path.join(raw, 'readback.json')) or {}
        acceptance = _load_json(os.path.join(raw, 'acceptance.json')) or {}
        out['verdict'] = compat.get('verdict')
        out['passed'] = readback.get('passed')
        out['accepted'] = acceptance.get('passed')
        out['needs_excel_recalc'] = readback.get('needs_excel_recalc', True)
    root = os.path.realpath(workdir)
    recorded = readback.get('workbook')
    candidates = [recorded] if isinstance(recorded, str) and recorded else []
    candidates.extend(os.path.join(root, name) for name in sorted(
        os.listdir(root), key=lambda name: (name.lower() == 'clean_template.xlsm', name)))
    for filename in candidates:
        candidate = os.path.realpath(filename)
        try:
            contained = os.path.commonpath([candidate, root]) == root
        except ValueError:
            contained = False
        if (not contained or not candidate.lower().endswith('.xlsm')
                or not os.path.isfile(candidate)):
            continue
        out['xlsm_path'] = candidate
        out['xlsm_name'] = os.path.basename(candidate)
        out['intermediate'] = out['xlsm_name'].lower() == 'clean_template.xlsm'
        try:
            out['package_valid'] = zipfile.is_zipfile(candidate)
            out['fingerprint_matches'] = config._sha256(candidate)[:16] == readback.get('sha256_after')
        except OSError:
            pass
        if os.path.realpath(str(acceptance.get('output') or '')) != candidate:
            out['accepted'] = False
        break
    return out


def artifact_warning(artifacts):
    if artifacts.get('intermediate'):
        return '仅留下中间模板，可下载查看，不是完成的模型。'
    if any(artifacts.get(key) is not True for key in
           ('passed', 'accepted', 'fingerprint_matches', 'package_valid')):
        return '文件未通过完整检查，仍可下载查看。'
    return None


def stop_reason(log_tail, fallback=None):
    """从流水线输出里把「为什么停下」那一行捞出来。"""
    m = STOP_RE.search(log_tail or "")
    if m:
        return m.group(0).strip()
    return fallback or "流水线在写出估值模型之前停下了。"


def task_title(task, options):
    """结论卡标题：公司名 · 报告期，取得到哪个用哪个。"""
    name = (options.get("company_name") or task.get("company_name")
            or "未识别公司")
    period = options.get("report_period") or task.get("report_period") or ""
    return (name + " · " + period) if period else name


def stop_advice(options):
    """停下之后该做什么。"""
    if options.get("interim_only"):
        return [
            "这次没有年报，年度数据整片留空，绝对估值和相对估值都算不出结果。"
            "要让它们出结果，需要年报（一份年报含相邻两年数据）。",
        ]
    return ["按上面的原因处理，然后重新上传 PDF 或回到任务页重跑。"]


def _load_json(path):
    try:
        with io.open(path, encoding="utf-8") as fh:
            value = json.load(fh)
            return value if isinstance(value, dict) else None
    except (OSError, ValueError):
        return None


def _human_card():
    """skill 里那张结论卡的模型层，直接复用。

    页面上的结论和「请先看这个.html」必须是同一套说法，否则同一次运行会给出
    两个版本。与其抄一遍规则，不如引同一个函数 —— runner 本来就已经在调
    skill 的 scripts/ 目录，这不是新引入的依赖。
    """
    if config.SCRIPTS_DIR not in sys.path:
        sys.path.insert(0, config.SCRIPTS_DIR)
    import human_card
    return human_card


def summarize(run_dir):
    """给办公人员看的结论：这批 PDF 的年报数据填进了哪几年。

    不写「这份模型能不能做估值」——那是读数字的人自己的判断，写在交付页上只会
    让一份正常降级的产出看起来像失败。办公人员要的是交付确认：哪几年填上了、
    哪几年这批文件里没有、还有哪些格子要人工填。
    """
    run_dir = run_dir or ""
    raw = os.path.join(run_dir, "_AI参考_可删")
    # 逐年归属要报告期，只有 resolved 那份才有；原始 params.json 里的 reports
    # 只有路径，用它会把「哪份年报填了哪一年」全部说成未知。
    params = (_load_json(os.path.join(raw, "params.resolved.json"))
              or _load_json(os.path.join(run_dir, "params.json")) or {})
    compat = _load_json(os.path.join(raw, "compat.json")) or {}
    gaps = _load_json(os.path.join(raw, "gaps.json")) or {}
    rb = _load_json(os.path.join(raw, "readback.json")) or {}

    facts = _human_card().fill_summary(params, compat, gaps,
                                       rb.get("downstream") or {})
    years = facts["years"]
    have = [r["year"] for r in years if r["ok"]]
    miss = [r["year"] for r in years if not r["ok"]]

    if not years:
        # 还没跑、或流水线在写出年度数据之前就停了。这时候说「填了哪几年」是
        # 空话，标题直接落到「没有生成」上，理由由 message 给。
        return {"level": "bad", "title": "没有生成估值模型",
                "verdict": compat.get("verdict") or gaps.get("verdict"),
                "years": [], "manual": [], "manual_short": [], "message": None}

    title = "已生成估值模型"
    if miss:
        level = "warn"
    else:
        level = "ok"
    return {"level": level, "title": title,
            "verdict": compat.get("verdict") or gaps.get("verdict"),
            "years": years, "manual": facts["manual"],
            "manual_short": facts["manual_short"],
            "have": have, "miss": miss, "message": None}


def write_stop_card(run_dir, title, reason, files, log_tail=None, advice=None):
    """提前停下时自己写一张结论卡。

    skill 的 run_pipeline.py 只在少数分支写卡：第 1 步「这批文件里没有一份年报」
    和第 2-7 步的任何一次失败都是直接 return 1，磁盘上一个字都没有。用户在页面上
    看到「原因见结论卡」，点开却是空的 —— 这正是要在这里补上的那张卡。
    """
    rows = "".join(
        "<tr><td>%s</td><td>%s</td><td>%s</td><td>%s</td></tr>"
        % (html.escape(f.get("orig_name") or "—"),
           html.escape(f.get("kind_cn") or "认不出"),
           html.escape(f.get("period") or "—"),
           html.escape(f.get("company") or "—"))
        for f in (files or []))
    steps = "".join("<li>%s</li>" % html.escape(x) for x in (advice or []))
    html_text = """<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8">
<title>结论 · %(title)s</title>
<style>
body{font-family:"Microsoft YaHei",system-ui,sans-serif;line-height:1.7;
margin:0;padding:24px;color:#1f2933;background:#f7f9fb}
h1{font-size:20px;margin:0 0 4px}
.verdict{background:#fdecea;border-left:4px solid #d64545;padding:12px 16px;
border-radius:4px;margin:16px 0}
.why{background:#fff;border:1px solid #e3e8ee;border-radius:6px;padding:12px 16px}
table{border-collapse:collapse;width:100%%;margin-top:12px;font-size:14px}
th,td{border:1px solid #e3e8ee;padding:6px 10px;text-align:left}
th{background:#eef2f6}
code{background:#eef2f6;padding:1px 5px;border-radius:3px}
</style></head><body>
<h1>%(title)s</h1>
<p style="color:#7b8794;margin:0">由本机 Web 版生成 · 流水线在写出估值模型之前停下</p>
<div class="verdict"><strong>没有生成估值模型。</strong><br>%(reason)s</div>
<h2 style="font-size:16px">这批文件是什么</h2>
<table><thead><tr><th>文件</th><th>类型</th><th>报告期</th><th>公司</th></tr></thead>
<tbody>%(rows)s</tbody></table>
%(advice)s
</body></html>
""" % {
        "title": html.escape(title),
        "reason": html.escape(reason),
        "rows": rows,
        "advice": ("<h2 style=\"font-size:16px\">下一步</h2><ul>%s</ul>" % steps)
                  if steps else "",
    }
    md = "# %s\n\n**没有生成估值模型。** %s\n\n| 文件 | 类型 | 报告期 | 公司 |\n|---|---|---|---|\n%s\n" % (
        title, reason,
        "\n".join("| %s | %s | %s | %s |"
                  % (f.get("orig_name") or "—", f.get("kind_cn") or "认不出",
                     f.get("period") or "—", f.get("company") or "—")
                  for f in (files or [])))
    if steps:
        md += "\n## 下一步\n" + "".join("- %s\n" % x for x in advice)
    with io.open(os.path.join(run_dir, "请先看这个.html"), "w", encoding="utf-8") as fh:
        fh.write(html_text)
    with io.open(os.path.join(run_dir, "请先看这个.md"), "w", encoding="utf-8") as fh:
        fh.write(md)
    return os.path.join(run_dir, "请先看这个.html")
