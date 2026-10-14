"""[card] The one document an office user opens, rendered twice.

A .md file double-clicked on Windows opens in Notepad and shows `#`, `|` and `**`
as literal characters -- a table becomes a row of pipes. So the human copy is
.html, which every machine already has a program for.

The same content is also written as .md, because the reader of the other half of
this pipeline is an agent, and Markdown is what it reads cheaply. Both are
rendered from one model, so neither is a lossy conversion of the other and the two
can never drift.

This is a **delivery confirmation**, not a verdict. The office user's question is
"which of the numbers I handed you actually landed in the workbook", so the card
leads with the per-year fill list. It deliberately does not argue about whether
the model is usable: that judgement belongs to whoever reads the numbers, and
putting it here painted a red warning on correctly-degraded deliverables.

Two rules the content follows, both learned from what went wrong before:

* **Self-contained.** Nothing here points at `_AI参考_可删/`. The user is free to
  delete that folder, and a document that says "see gaps.md" becomes a dead end the
  moment they do.
* **Plain words first.** The workbook raises no error when data is missing, so if
  this document does not say "this column is empty", nothing else will.
"""
import html
import os
import re

ESC = html.escape

# 一年填了多少行才算「这一年填上了」。低于它的一年是要点名说出来的空格，
# 但它是陈述，不是判定：模板自己不会因为一列不满而报错。
FILL_OK_PCT = 80.0


def _fmt(v, nd=4):
    """A cell value as a reader sees it. Excel returns errors as negative ints."""
    if v is None:
        return "-"
    if isinstance(v, bool):
        return "是" if v else "否"
    if isinstance(v, int) and v < -2146826000:
        return "计算错误"
    if isinstance(v, float):
        return ("%%.%df" % nd) % v
    return str(v)


def _year_of(value):
    """2025-12-31 / "2025" / 2025 -> 2025; anything else -> None."""
    m = re.match(r"^\s*(\d{4})", str(value if value is not None else ""))
    return int(m.group(1)) if m else None


def _span(years):
    """[2022,2023,2024,2025] -> "2022-2025"; a hole breaks the run: "2020、2022"."""
    ys = sorted({int(y) for y in years})
    if not ys:
        return ""
    runs, start, prev = [], ys[0], ys[0]
    for y in ys[1:]:
        if y == prev + 1:
            prev = y
            continue
        runs.append((start, prev))
        start = prev = y
    runs.append((start, prev))
    return "、".join("%d" % a if a == b else "%d-%d" % (a, b) for a, b in runs)


def _source_by_year(params):
    """Which annual report supplies which 年度数据 column.

    An annual report writes "本期 + 上年同期", so the report dated 2025-12-31
    fills both 2025 and 2024. That is not a guess: gaps.py says「一份年报含相邻
    两年数据」and compat.py derives missing_reports as [y-12-31, (y+1)-12-31],
    which is the same statement read backwards.
    """
    out = {}
    for r in params.get("reports") or []:
        y = _year_of(r.get("period"))
        if not y:
            continue
        name = os.path.basename(r.get("pdf") or "-")
        out.setdefault(y, name)
        out.setdefault(y - 1, name)
    return out


def _fill_rows(params, compat):
    """The per-year 年度数据 list: filled, by which file, or empty."""
    src = _source_by_year(params)
    rows = []
    for x in compat.get("coverage") or []:
        y = _year_of(x.get("year"))
        if y is None:
            continue
        ok = (x.get("pct") or 0) >= FILL_OK_PCT
        filled, total = x.get("filled") or 0, x.get("total") or 0
        if ok:
            note = "已填 %d/%d 项" % (filled, total)
        elif src.get(y):
            note = "只填上 %d/%d 项，这一列不完整" % (filled, total)
        else:
            note = "空 —— 这批文件里没有覆盖 %d 年的年报" % y
        rows.append({"year": y, "ok": ok, "source": src.get(y), "note": note})
    return rows


def _company_rows(params, gaps):
    """What the tool read out of the PDFs and wrote into the workbook."""
    c = params.get("company") or {}
    m = params.get("market") or {}
    # 「财报封面」不是「年报封面」：只传了一季报时公司名和代码也是从它读的，
    # 写成年报会让办公人员以为系统认错了文件。
    rows = [
        ("公司名称", c.get("name"), "财报封面"),
        ("股票代码", c.get("code"), "财报封面"),
        ("报告期", c.get("report_period"), "财报封面"),
        ("总股本（百万股）", m.get("total_shares_mn"), "财报「股份变动情况」"),
        ("流通股本（百万股）", m.get("float_shares_mn"), "财报「股份变动情况」"),
    ]
    # 分红那几格写的是 0，但那是年报原文写明「不派发现金红利」的结果，不是漏填。
    # 不写这一句，办公人员看到 0 会当成 bug 去查。
    ev_all = ((gaps or {}).get("dividend") or {}).get("evidence") or {}
    for field in sorted(ev_all):
        ev = ev_all[field] or {}
        page = ev.get("page", "-")
        if ev.get("basis") == "explicit_absence":
            rows.append((field, 0, "年报第 %s 页写明本期不分配" % page))
        elif ev.get("basis") == "numeric" and ev.get("cross_checked"):
            rows.append((field, "见「年度数据」表",
                         "取自年报第 %s 页，与同页现金分红金额对账一致" % page))
    out = []
    for label, value, source in rows:
        if value is None:
            text = "未填 —— 自动识别失败，请按行情软件手填"
        elif isinstance(value, str):
            text = value
        else:
            text = _fmt(value)
        out.append((label, text, source))
    return out


def _manual_rows(params, gaps, downstream):
    """The numbers no PDF contains -- the workbook is waiting for a human.

    Two different jobs, so they read differently: 收盘价 is a lookup (行情软件里
    抄一个数), 预测假设 is a judgement (没有人能替你定). Both say where to get
    it, because「未填」on its own tells an office user nothing actionable.
    """
    rows = []
    m = params.get("market") or {}
    hist = m.get("history") or {}
    if hist:
        years = _span([_year_of(k) for k in hist if _year_of(k)])
        rows.append(("历年年末收盘价", "已填（%s 年）" % years,
                     "行情软件里查每年 12 月最后一个交易日；PE/PS/PB/PCF "
                     "由系统按它和年报财务数算出"))
    else:
        rows.append(("历年年末收盘价", "未填",
                     "行情软件里查每年 12 月最后一个交易日；PE/PS/PB/PCF "
                     "由系统按它和年报财务数算出"))
    for it in analyst_items(gaps):
        value = (downstream or {}).get(it.get("where"))
        if value in (None, 0):
            hint = (it.get("if_blank") or "").split("。")[0]
            state = "未填" + ("（%s）" % hint if hint else "")
        else:
            state = _fmt(value)
        rows.append((it.get("what") or "-", state,
                     "对未来的判断，年报里没有，系统不替你猜"))
    return rows


def manual_short(params, gaps):
    """The manual work, as two words instead of seven rows.

    The page has one line to spend on this, so it gets「历年年末收盘价、6 个预测
    假设」rather than the full list. The card prints the full table.
    """
    out = []
    if not ((params or {}).get("market") or {}).get("history"):
        out.append("历年年末收盘价")
    n = len(analyst_items(gaps))
    if n:
        out.append("%d 个预测假设" % n)
    return out


def analyst_items(gaps):
    groups = {g.get("id"): g for g in ((gaps or {}).get("groups") or [])}
    return (groups.get("analyst") or {}).get("items") or []


def fill_summary(params, compat, gaps, downstream=None):
    """The facts the card and the web page both report.

    One function, two renderers: the .html/.md card and the box on the task page.
    Before this they disagreed -- the card listed the years that landed while the
    page announced whether the model was usable -- so the same run could be read
    two ways depending on which one you opened.
    """
    params, gaps = params or {}, gaps or {}
    return {"years": _fill_rows(params, compat or {}),
            "manual": _manual_rows(params, gaps, downstream or {}),
            "manual_short": manual_short(params, gaps)}


def build(params, rb, pre, post, compat, gaps, cleanup=None, version=None,
          verified=None):
    """The content model. Rendering decides layout; this decides what is true."""
    params = params or {}
    rb = rb or {}
    compat = compat or {}
    gaps = gaps or {}
    c = params.get("company") or {}
    d = rb.get("downstream") or {}
    verdict = compat.get("verdict", rb.get("verdict", "full"))

    name = c.get("name", "-")
    code = c.get("code", "-")
    period = c.get("report_period", "-")

    # ---- (1) which years landed in the workbook --------------------------
    facts = fill_summary(params, compat, gaps, d)
    years, company, manual = facts["years"], _company_rows(params, gaps), facts["manual"]
    have = [r["year"] for r in years if r["ok"]]
    miss = [r["year"] for r in years if not r["ok"]]

    if verdict == "blocked":
        headline = "没有生成估值模型 —— 有几份文件对不上，这批 PDF 没有被使用。"
    elif not years:
        headline = "已按模板生成估值模型，但没有留下逐年填入记录，请打开文件核对。"
    elif have and not miss:
        headline = ("已按模板生成估值模型，%s 年的年报数据已填入「年度数据」表。"
                    % _span(have))
    elif have:
        headline = ("已按模板生成估值模型。%s 年的年报数据已填入「年度数据」表；"
                    "%s 年这批文件里没有，那几列是空的。"
                    % (_span(have), _span(miss)))
    else:
        headline = "已按模板生成估值模型。这批文件里没有年报，年度数据整片留空。"

    # 转述块是给「把结论说给同事听」用的，所以只留一句结论 + 人工项，逐年明细
    # 在下面那张表里，不在这里重复一遍。
    transcribe = [headline]
    if facts["manual_short"]:
        transcribe.append("年报里没有、要人工填的还有：%s。"
                          % "、".join(facts["manual_short"]))
    transcribe.append("逐年填入情况和每个数的来源见下表。")

    reports = [(os.path.basename(r.get("pdf") or "-"), r.get("period") or "-")
               for r in params.get("reports") or []]
    interim = [(os.path.basename(r.get("pdf") or "-"), r.get("period") or "-")
               for r in params.get("reports_interim") or []]

    if verified is None:
        # 这句话专指保真度（工作表/公式/图表/宏/控件有没有被改坏），所以取
        # fidelity 那一条断言，而不是拿总判定顶 —— 总判定里还混着别的断言，
        # 它们失败并不代表文件被改坏了。
        for a in (post or {}).get("post_assertions") or []:
            if a.get("id") == "fidelity":
                verified = bool(a.get("passed"))
                break
        else:
            verified = bool((post or {}).get("passed"))

    return {
        "title": "%s（%s）估值模型" % (name, code),
        "notes": [],
        "headline": headline,
        "period": period,
        "valuation_date": c.get("valuation_date", "-"),
        "verdict": verdict,
        "transcribe": transcribe,
        "years": years,
        "problems": [],
        "company": company,
        "manual": manual,
        "reports": reports,
        "interim": interim,
        "verified": verified,
        "blocked": verdict == "blocked",
        "version": version,
    }


# ---------------------------------------------------------------- markdown --

def to_markdown(m):
    """The same content as to_html, in the markup an agent reads for free."""
    L = ["# %s" % m["title"], ""]
    L.append("**%s**" % m["headline"])
    L.append("")
    if m.get("transcribe"):
        L.append("> **可以直接转述给同事的话：**")
        for t in m["transcribe"]:
            L.append("> %s" % t)
        L.append("")
    for n in m.get("notes") or []:
        L.append(n)
        L.append("")

    if m.get("problems"):
        L.append("## 需要你处理的问题")
        L.append("")
        for x in m["problems"]:
            L.append("- %s" % x)
        L.append("")

    if m.get("years"):
        L.append("## 年报数据填入情况")
        L.append("")
        L.append("| 年份 | 来源年报 | 「年度数据」表 |")
        L.append("|---|---|---|")
        for r in m["years"]:
            L.append("| %d | %s | %s |"
                     % (r["year"], r.get("source") or "—", r["note"]))
        L.append("")

    if m.get("company"):
        L.append("## 系统自动填入的数据")
        L.append("")
        L.append("| 项目 | 填入值 | 来源 |")
        L.append("|---|---|---|")
        for label, value, source in m["company"]:
            L.append("| %s | %s | %s |" % (label, value, source))
        L.append("")

    if m.get("manual"):
        L.append("## 年报里没有、要人工填的数据")
        L.append("")
        L.append("| 项目 | 现在 | 去哪儿找 |")
        L.append("|---|---|---|")
        for label, state, where in m["manual"]:
            L.append("| %s | %s | %s |" % (label, state, where))
        L.append("")

    if m["reports"] or m["interim"]:
        L.append("## 这份文件是怎么来的")
        L.append("")
        L.append("| 来源文件 | 报告期 |")
        L.append("|---|---|")
        for f, p in m["reports"]:
            L.append("| %s | %s |" % (f, p))
        for f, p in m["interim"]:
            L.append("| %s（季报/半年报，只进「报告期历史数据」） | %s |" % (f, p))
        L.append("")

    if m["verified"] is not None:
        L.append("模板完整性：%s —— 工作表、公式、图表、宏、控件与原件逐项比对%s。"
                 % ("已核对" if m["verified"] else "**有问题**",
                    "一致" if m["verified"] else "**不一致**"))
        L.append("")

    L.append("报告期 %s ｜ 估值日期 %s ｜ 工具版本 %s"
             % (m["period"], m["valuation_date"], m["version"] or "-"))
    L.append("")
    return "\n".join(L).strip() + "\n"


# ------------------------------------------------------------------- html --

_CSS = """
body{font-family:"Microsoft YaHei","Segoe UI",sans-serif;max-width:820px;
     margin:32px auto;padding:0 20px;color:#1a1a1a;line-height:1.7}
h1{font-size:22px;border-bottom:2px solid #d0d0d0;padding-bottom:8px}
h2{font-size:17px;margin-top:28px;color:#0b4f8a}
.headline{font-size:18px;font-weight:bold;padding:14px 16px;border-radius:6px;
          background:#f1f9f2;border-left:5px solid #2e8b57;margin:18px 0}
.headline.partial{background:#fff9e8;border-left-color:#c9a227}
table{border-collapse:collapse;width:100%;margin:10px 0}
th,td{border:1px solid #d5d5d5;padding:7px 10px;text-align:left;font-size:14px}
th{background:#f2f5f8}
td.empty{color:#b03030}
.transcribe{background:#f4f8fc;border:1px solid #cfe0f0;border-radius:6px;
            padding:12px 16px;margin:18px 0;font-size:14px}
.transcribe b{color:#0b4f8a}
.meta{color:#666;font-size:13px;margin-top:30px;border-top:1px solid #e2e2e2;
      padding-top:10px}
"""


def _cls(m):
    return "partial" if any(not r["ok"] for r in m.get("years") or []) else ""


def to_html(m):
    e = ESC
    P = []
    P.append("<!DOCTYPE html>")
    P.append('<html lang="zh-CN">')
    P.append("<head>")
    P.append('<meta charset="utf-8">')
    P.append("<title>%s</title>" % e(m["title"]))
    P.append("<style>%s</style>" % _CSS)
    P.append("</head>")
    P.append("<body>")
    P.append("<h1>%s</h1>" % e(m["title"]))

    P.append('<div class="transcribe"><b>可以直接转述给同事的话：</b><br>%s</div>'
             % "<br>".join(e(t) for t in m["transcribe"]))

    P.append('<div class="headline %s">%s</div>' % (_cls(m), e(m["headline"])))
    for n in m.get("notes") or []:
        P.append("<p>%s</p>" % e(n))

    if m.get("problems"):
        P.append("<h2>需要你处理的问题</h2><ul>")
        for x in m["problems"]:
            P.append("<li>%s</li>" % e(x))
        P.append("</ul>")

    if m.get("years"):
        P.append("<h2>年报数据填入情况</h2>")
        P.append("<table><tr><th>年份</th><th>来源年报</th>"
                 "<th>「年度数据」表</th></tr>")
        for r in m["years"]:
            P.append('<tr><td>%d</td><td>%s</td><td class="%s">%s</td></tr>'
                     % (r["year"], e(r.get("source") or "—"),
                        "" if r["ok"] else "empty", e(r["note"])))
        P.append("</table>")

    if m.get("company"):
        P.append("<h2>系统自动填入的数据</h2>")
        P.append("<table><tr><th>项目</th><th>填入值</th><th>来源</th></tr>")
        for label, value, source in m["company"]:
            P.append("<tr><td>%s</td><td>%s</td><td>%s</td></tr>"
                     % (e(label), e(value), e(source)))
        P.append("</table>")

    if m.get("manual"):
        P.append("<h2>年报里没有、要人工填的数据</h2>")
        P.append("<table><tr><th>项目</th><th>现在</th><th>去哪儿找</th></tr>")
        for label, state, where in m["manual"]:
            P.append("<tr><td>%s</td><td>%s</td><td>%s</td></tr>"
                     % (e(label), e(state), e(where)))
        P.append("</table>")

    if m["reports"] or m["interim"]:
        P.append("<h2>这份文件是怎么来的</h2>")
        P.append("<table><tr><th>来源文件</th><th>报告期</th></tr>")
        for f, p in m["reports"]:
            P.append("<tr><td>%s</td><td>%s</td></tr>" % (e(f), e(p)))
        for f, p in m["interim"]:
            P.append("<tr><td>%s（季报/半年报，只进「报告期历史数据」）</td>"
                     "<td>%s</td></tr>" % (e(f), e(p)))
        P.append("</table>")

    if m["verified"] is not None:
        P.append("<p>模板完整性：%s —— 工作表、公式、图表、宏、控件与原件逐项"
                 "比对%s。</p>"
                 % ("已核对" if m["verified"] else "<b>有问题</b>",
                    "一致" if m["verified"] else "<b>不一致</b>"))
    P.append('<p class="meta">报告期 %s ｜ 估值日期 %s ｜ 工具版本 %s</p>'
             % (e(str(m["period"])), e(str(m["valuation_date"])),
                e(str(m["version"] or "-"))))
    P.append("</body></html>")
    return "\n".join(P) + "\n"