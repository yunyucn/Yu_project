"""[5] Reporter: one-page company card plus a traceable run manifest."""
import argparse
import datetime
import io
import json
import os

import com_excel as cx

# Excel hands errors back as negative CVErr integers, so printing a "value"
# straight from a cell renders #DIV/0! as -2146826281. Name it instead.
# full/degraded/facts_only is the pipeline's vocabulary. The reader of the
# delivery note is not the pipeline author.
VERDICT_CN = {"full": "\u5b8c\u6574\u53ef\u7528", "degraded": "\u964d\u7ea7\u53ef\u7528",
              "facts_only": "\u53ea\u51fa\u4e8b\u5b9e\u5c42\uff0c\u4e0d\u53ef\u7528\u4e8e\u4f30\u503c"}


ERR_NAME = {-2146826281: "#DIV/0!", -2146826265: "#REF!", -2146826273: "#VALUE!",
            -2146826288: "#NULL!", -2146826246: "#N/A", -2146826259: "#NAME?",
            -2146826274: "#NUM!"}


def fmt(v, nd=4):
    if v is None:
        return "-"
    if cx.is_error(v):
        return "**%s**（不可用）" % ERR_NAME.get(int(v), "计算错误")
    if isinstance(v, (int, float)):
        return ("%%.%df" % nd) % v
    return str(v)


VERDICT_CN = {"full": "完整", "degraded": "降级", "facts_only": "只出事实层"}


def build_summary(params, rb, compat, gaps, cleanup):
    """Plain-language verdict, first thing on the card.

    The card used to open with a completeness banner written in pipeline
    vocabulary. The person who has to act on this file is a finance analyst, not
    the pipeline author, and the template raises no error -- so if the summary
    does not say "do not use this number", nothing else will. Plain words first,
    evidence after.
    """
    d = rb.get("downstream", {})
    verdict = compat.get("verdict", rb.get("verdict", "full"))
    dcf_ok = compat.get("dcf_reliable", True)
    rel_ok = compat.get("rel_reliable", True)
    L = []

    if verdict == "full" and dcf_ok and rel_ok:
        L.append("**这份文件可以用。** 六年历史数据齐全，估值结果可信。")
    elif verdict == "degraded":
        L.append("**这份文件部分可用。** 有些数据没凑齐，下面的表逐项说明哪些数字能信、"
                 "哪些不能信。")
    else:
        L.append("**这份文件目前不能用来做估值。** 历史数据或市场数据不齐，"
                 "模型照样会算出一串数字，但那些数字没有意义。")
        if rb.get("incomplete_history"):
            L.append("")
            L.append("文件名里的 `待补数据` 就是这个意思：仅供人工核对事实。")
    L.append("")

    L.append("## 能不能用")
    L.append("")
    L.append("| 你要做的事 | 能不能用 | 原因 |")
    L.append("|---|---|---|")
    cov = compat.get("coverage", [])
    have_years = [str(x["year"]) for x in cov if x.get("pct", 0) >= 80]
    miss_years = compat.get("coverage_gap_years", [])
    if miss_years:
        L.append("| 看历史财务数据 | 可以看，但有缺 | 已填 %s；缺 %s，这几列是空的 |"
                 % ("、".join(have_years) or "-", "、".join(str(y) for y in miss_years)))
    else:
        L.append("| 看历史财务数据 | **可以** | 六年齐全，从年报 PDF 自动提取 |")
    L.append("| DCF 绝对估值 | %s | %s |"
             % ("**可以**" if dcf_ok else "**不可以**",
                "假设齐全" if dcf_ok
                else "%d 个关键假设因缺数据被算成 0；有效税率变 0 会让结果连正负号都不可靠"
                     % len(compat.get("dcf_critical_collapsed", []))))
    L.append("| 相对估值（PE/PS/PB） | %s | %s |"
             % ("**可以**" if rel_ok else "**不可以**",
                "市场倍数齐全" if rel_ok
                else "缺 %d 个历史估值倍数；这些是股市行情，任何年报里都没有，必须另外提供"
                     % compat.get("market_missing_count", 0)))
    L.append("| 看敏感性分析和图表 | 能打开，数字别用 | 它们建立在上面不可信的结果之上 |")
    L.append("")

    L.append("## 下一步做什么")
    L.append("")
    blocks = [g for g in gaps.get("groups", []) if g.get("blocking")]
    if not blocks:
        L.append("不用补数据，这份文件已经可用。")
        L.append("")
    else:
        for i, g in enumerate(blocks, 1):
            L.append("%d. **%s**（%d 项）" % (i, g.get("title", ""), len(g.get("items", []))))
            if g.get("id") == "reports":
                if g.get("minimal_pdfs"):
                    L.append("   最小补法：补 %s 各一份就够了 —— 一份年报含相邻两年的数据。"
                             % "、".join("FY%d" % y for y in g["minimal_pdfs"]))
                for it in g.get("items", []):
                    L.append("   - 缺 %s：%s" % (it.get("what"), it.get("how", "")))
            elif g.get("id") == "market":
                L.append("   这些是历史估值倍数，要从行情数据源取，填进运行参数的 "
                         "`market.history`：")
                for it in g.get("items", []):
                    L.append("   - %s（%s，缺 %d 格）"
                             % (it.get("what"), it.get("where"), it.get("count", 0)))
            elif g.get("id") == "analyst":
                L.append("   这几格要靠人的判断，AI 不能替你拍板。"
                         "**留空不等于安全：Excel 把空当 0**，空着会算出一个假的好结果：")
                for it in g.get("items", []):
                    L.append("   - %s（`%s`）：%s"
                             % (it.get("what"), it.get("where"), it.get("if_blank", "")))
            L.append("")
        L.append("补完之后重跑一遍，判定会从「只出事实层」往上走，直到「完整」。")
        L.append("")

    PURPOSE = {
        "数据输入": "以「元」为单位的原始数据。它不影响估值，但「主营业务」等表读它的年份表头",
        "报告期历史数据": "按季度排列的历史数据，年报里没有，需要季报",
        "数据输入季度数据": "季度原始数据，需要季报",
        "收入成本预测": "分产品的收入成本表，需要年报里的「主营业务分产品」明细",
        "主营业务": "分产品/分行业的收入拆分，需要年报里的「主营业务分产品、分地区」明细",
        "情景表": "乐观/一般情景快照，属于分析师判断，不是年报数据",
    }
    rows = []
    for act in (cleanup or {}).get("actions", []):
        if act.get("tier") != "P1" or act.get("action") != "blank_literals":
            continue
        name = act.get("sheet")
        # 年度数据 is only cleared in O:T, an off-print mirror of the same rows;
        # the sheet itself is filled. Naming it here would read as "this sheet is
        # empty", which is the opposite of true.
        if not name or name == "年度数据":
            continue
        # A sheet the run now fills is not "empty": naming it here would make the
        # card contradict the 季度数据 section three lines above it.
        if act.get("filled_by"):
            continue
        if name not in [r[0] for r in rows]:
            rows.append((name, PURPOSE.get(name, act.get("note", ""))))
    dv = (gaps or {}).get("dividend") or {}
    if dv.get("evidence") or dv.get("blank_rows"):
        L.append("## 分红口径（DDM 那一块为什么是 0）")
        L.append("")
        L.append("DDM（股利折现）只认分红：公司不分红，这一块就是 0。下面说清这个 0 "
                 "是**事实**还是**没取到数**——两者在表里长得一模一样。")
        L.append("")
        for field in sorted(dv["evidence"]):
            ev = dv["evidence"][field] or {}
            page = ev.get("page", "-")
            basis = ev.get("basis")
            if basis == "explicit_absence":
                L.append("- **%s**：年报第 %s 页原文写明本期不分红，所以 0 是**对的**。"
                         % (field, page))
            elif basis == "numeric":
                if ev.get("cross_checked"):
                    L.append("- **%s**：取自年报第 %s 页，并与同页「现金分红金额」对账一致。"
                             % (field, page))
                else:
                    L.append("- **%s**：取自年报第 %s 页（年报没印现金分红金额，无法二次对账）。"
                             % (field, page))
            elif basis == "cross_check_failed":
                L.append("- **%s**：年报第 %s 页的数与同页「现金分红金额」对不上，"
                         "**已拒绝写入**，请人工核对后再填。" % (field, page))
        L.append("")
        if dv.get("blank_rows"):
            rows_ = sorted({b["row"] for b in dv["blank_rows"]})
            L.append("中期分红（`年度数据` 第 %s 行）留空：年报里没有这一项，要半年报才有。"
                     "模板出厂时写死 0，会把有中期分红的公司静默抹平；现在留空，"
                     "该期确实没有中期分红就填 0。" % "、".join(str(r) for r in rows_))
            L.append("")
    excluded = (params or {}).get("_excluded_reports") or []
    if excluded:
        L.append("## 收到但没有写进「年度数据」的报告")
        L.append("")
        L.append("这几份不是年报，**数字没有写进 `年度数据`**——模型按报告期写数："
                 "把一季报或半年报当年报用，会把三个月的数字写进年度那一列，"
                 "工作簿照样算、不报错，结果却是错的。它们改填 `报告期历史数据`，"
                 "见上面「季度数据」一节。")
        L.append("")
        for r in excluded:
            L.append("- %s %s（%s）" % (r.get("kind_cn") or "?", r.get("period") or "-",
                                        os.path.basename(r.get("pdf") or "-")))
        L.append("")
    L.append("## 表里哪些地方是空的")
    L.append("")
    L.append("模板出厂时带着一家样板公司的数据，这些已经被清掉——否则新公司的表里会"
             "静默混进上一家公司的数字。下面几张表**不参与估值计算**，但翻表时会是空的：")
    L.append("")
    for name, why in sorted(rows):
        L.append("- `%s`：%s" % (name, why))
    L.append("")
    q0 = (rb or {}).get("quarterly") or {}
    if q0.get("columns"):
        L.append("`报告期利润变化` 的 240 个公式读的就是「报告期历史数据」；本批已填 %d 列，"
                 "所以那张表有数、折线画得出来。" % len(q0["columns"]))
    else:
        L.append("另外 `报告期利润变化` 有 240 个公式读「报告期历史数据」，后者是空的，"
                 "它会显示 `#DIV/0!`；拿到季报或半年报后重跑一次即恢复。")
    L.append("")
    return L


def _quarterly_shortfalls(q):
    """Quarterly columns written from a report that came out short.

    A column can exist and be short at the same time: the balance sheet parsed,
    the income statement did not. The workbook raises nothing, and 报告期利润变化
    draws its line from the income-statement rows, so the newest point is simply
    missing. Which statements count as missing is decided in the quarterly stage
    -- only what that report kind is meant to carry -- and this only renders it.
    """
    cols = {c["period"]: c for c in (q.get("columns") or [])}
    out = []
    for item in q.get("incomplete") or []:
        c = cols.get(item.get("period"))
        if not c:
            continue          # the template has no column for that period
        out.append((item["period"], c.get("col"), item.get("missing") or []))
    return out


def _quarterly_shortfall_lines(q):
    """The shortfall written the way the reader can act on it.

    The cause is a statement title this tool did not recognise -- a defect here,
    not a gap in the company's disclosure, and not something the person holding
    the workbook can fix. Say what is missing and what not to do with it; never
    hand over a config file to edit.
    """
    L = []
    for per, col, missing in _quarterly_shortfalls(q):
        L.append("- %s（%s 列）：缺 %s" % (per, col, "、".join(missing)))
    if not L:
        return L
    L.append("")
    L.append("这不是公司没披露，是解析工具没认出那几张表的标题。缺利润表的季度，"
             "`报告期利润变化` 的折线上就没有这个点；那一列只能用来核对资产表，"
             "不要拿它算利润类指标。这一条请交回给运行它的人处理。")
    if q.get("parser_warnings"):
        L.append("")
        L.append("解析器同时报了 %d 处「报表标题可能没认出来」，可作为排查线索。"
                 % len(q["parser_warnings"]))
    L.append("")
    return L


def build_delivery_note(params, rb, compat, gaps, cleanup):
    """Provenance, in the order a reader asks for it.

    The card answers "can I use this?". This answers "where did it come from, what
    did a human type in, and what is still missing?" -- the questions that arrive
    the moment the file is forwarded to someone who was not in the room.
    """
    company = params.get("company", {})
    market = params.get("market", {}) or {}
    L = []
    L.append("# 交付说明")
    L.append("")
    L.append("**%s（%s）** 的估值模型。" % (company.get("name", "-"), company.get("code", "-")))
    L.append("")
    L.append("| 项 | 值 |")
    L.append("|---|---|")
    L.append("| 报告期 | %s |" % company.get("report_period", "-"))
    L.append("| 估值日期 | %s |" % company.get("valuation_date", "-"))
    _v = compat.get("verdict", rb.get("verdict", "-"))
    L.append("| 数据判定 | %s |" % VERDICT_CN.get(_v, _v))
    L.append("| 映射表版本 | %s |" % params.get("_mapping_version", "-"))
    L.append("| 工具版本 | %s |" % params.get("_skill_version", "-"))
    L.append("")

    fp = params.get("_fingerprints") or {}
    if fp:
        L.append("## 代码指纹")
        L.append("")
        L.append("版本号是人写的，会忘；指纹是算出来的。将来查出某个缺陷时，")
        L.append("用这组指纹能反推哪些产出是那段代码跑的。")
        L.append("")
        L.append("| 文件 | sha256 前 8 位 |")
        L.append("|---|---|")
        for k in sorted(fp):
            L.append("| `%s` | `%s` |" % (k, fp[k]))
        L.append("")

    L.append("## 数据来源")
    L.append("")
    L.append("公司财务数据全部来自下列年报，逐项与 PDF 原值对过账（容差 0.005 百万元）：")
    L.append("")
    L.append("| 年报 PDF | 报告期 | 文件指纹 |")
    L.append("|---|---|---|")
    for r in params.get("reports", []):
        sha = (params.get("_pdf_sha") or {}).get(r.get("pdf"), "-")
        L.append("| %s | %s | `%s` |" % (os.path.basename(str(r.get("pdf"))), r.get("period"), sha))
    L.append("")

    inter = params.get("reports_interim") or []
    q = rb.get("quarterly") or {}
    L.append("## 季度数据（报告期历史数据）")
    L.append("")
    if inter:
        L.append("这批文件里还有 %d 份季报/半年报。它们**不进 `年度数据`**：把一季报或半年报"
                 "当年报用，模型会照常算出一串数、而且不报任何错。它们填进 "
                 "`报告期历史数据`——模板里唯一按季度排列的表，也是 `报告期利润变化` 那张图的"
                 "唯一数据来源。" % len(inter))
        L.append("")
        L.append("| 季报/半年报 PDF | 报告期 |")
        L.append("|---|---|")
        for r in inter:
            L.append("| %s | %s |" % (os.path.basename(str(r.get("pdf"))), r.get("period")))
        L.append("")
    if q.get("columns"):
        L.append("已填入 %d 个季度列、共 %d 格。表头是日期（由 `TODAY()` 生成），"
                 "所以是按**日期对号入座**，不是按第几列。" % (len(q["columns"]), q.get("cells", 0)))
        L.append("")
        L.append("| 报告期 | 列 | 格数 |")
        L.append("|---|---|---|")
        for c in q["columns"]:
            L.append("| %s | %s | %d |" % (c["period"], c["col"], c["cells"]))
        L.append("")
        short = _quarterly_shortfall_lines(q)
        if short:
            L.append("**下面这几个季度列只填了一部分：**")
            L.append("")
            L.extend(short)
        have = {c["period"] for c in q["columns"]}
        want = q.get("window") or []
        miss = [w for w in want if w not in have]
        if miss:
            L.append("模板的季度窗口是 %s 起共 %d 列，**还缺 %d 个季度**：%s。"
                     "缺的列留空，`报告期利润变化` 图上那几段就没有点；"
                     "补对应期间的季报/半年报（或年报）即可。"
                     % (want[0], len(want), len(miss), "、".join(miss)))
            L.append("")
        if q.get("skipped"):
            L.append("另有 %s 不在模板的季度窗口内，没有对应列，未写入。"
                     % "、".join(q["skipped"]))
            L.append("")
    else:
        L.append("没有季报或半年报落进这张表，`报告期历史数据` 是空的，"
                 "`报告期利润变化` 不会画出折线。补一份季报或半年报即可解锁。")
        L.append("")

    L.append("## 人工填入的数据")
    L.append("")
    hist = market.get("history") or {}
    typed = []
    if market.get("total_shares_mn") is not None:
        typed.append(("目录!F8 总股本（百万股）", market["total_shares_mn"]))
    if market.get("float_shares_mn") is not None:
        typed.append(("目录!F9 流通股本（百万股）", market["float_shares_mn"]))
    if market.get("close_price") is not None:
        typed.append(("目录!F12 收盘价（元）", market["close_price"]))
    if hist:
        typed.append(("年度数据!144-148 历史 PE/PS/PB/PCF",
                      "%d 个年度，来自运行参数 market.history" % len(hist)))
    for key, val in (params.get("analyst") or {}).items():
        typed.append(("分析师假设 %s" % key, val))
    if typed:
        L.append("下面这些不是年报里的数，是**人提供的**（行情数据、分析师判断）：")
        L.append("")
        L.append("| 项目 | 值 |")
        L.append("|---|---|")
        for k, v in typed:
            L.append("| %s | %s |" % (k, v))
        L.append("")
        L.append("> 如果这些是测试值或估算值，**不要对外当成真实市场数据使用**。")
    else:
        L.append("没有任何人工输入：市值参数与历史估值倍数都是空的，"
                 "六个预测假设也是空的（Excel 把空白当 0）。")
    L.append("")

    blocks = [g for g in gaps.get("groups", []) if g.get("blocking")]
    L.append("## 还缺什么")
    L.append("")
    if not blocks:
        L.append("没有阻塞项，这份文件可以直接用。")
    else:
        for g in blocks:
            L.append("- **%s**（%d 项）：%s"
                     % (g.get("title", ""), len(g.get("items", [])),
                        "；".join(str(it.get("what", "")) for it in g.get("items", [])[:4])))
    L.append("")
    L.append("细节和最小补法见 `%s`。" % "请先看这个.md")
    L.append("")

    L.append("## 怎么验收")
    L.append("")
    L.append("```bash")
    L.append("python a-share-valuation-fill/scripts/acceptance.py \\")
    L.append("  --pristine \"<原件模板.xlsm>\" \\")
    L.append("  --output   \"<%s>\" --equivalence" % os.path.basename(rb.get("workbook", "模型.xlsm")))
    L.append("```")
    L.append("")
    L.append("它比对两份文件的工作表、公式、图表、控件、宏与外部链接，"
             "再把同一份输入喂进去逐格比结果。`%s` 是它的输出。" % "验收报告.md")
    L.append("")
    L.append("## 这份说明的位置")
    L.append("")
    L.append("这份说明和同目录的解析结果、字段映射、兼容性报告、待补清单一样，"
             "都在 `_AI参考_可删/` 里。"
             "排查问题时才需要，日常使用不用打开。")
    L.append("")
    return "\n".join(L)


def _summary_model(params, compat, gaps, cleanup, version):
    """A minimal card model for a run that stopped before a workbook existed."""
    c = (params or {}).get("company") or {}
    return {
        "title": "%s（%s）估值模型" % (c.get("name", "-"), c.get("code", "-")),
        "headline": "这次没有生成估值模型 —— 流水线在写出文件之前停下了。",
        "notes": [], "period": c.get("report_period", "-"),
        "valuation_date": c.get("valuation_date", "-"),
        "verdict": compat.get("verdict", "facts_only"),
        "transcribe": ["这次没有产出估值文件，下面说明是为什么。",
                       "按原因处理完，重新上传 PDF 或回到任务页重跑一次。"],
        "years": [], "problems": [], "company": [], "manual": [],
        "reports": [], "interim": [],
        "verified": None, "version": version, "blocked": True,
    }


def build_card(params, rb, pre, post, compat, gaps, cleanup=None, version=None):
    """The human card. Delegates to human_card, which owns the layout.

    Kept as a function so callers written against the old signature keep working;
    the markup itself now lives in one place and is rendered twice from a model.
    """
    import human_card
    return human_card.to_markdown(
        human_card.build(params, rb, pre, post, compat, gaps, cleanup, version))


def build_card_legacy(params, rb, pre, post, compat, gaps, cleanup=None):
    c = params["company"]
    d = rb.get("downstream", {})
    verdict = compat.get("verdict", rb.get("verdict", "full"))
    L = []
    L.append("# %s (%s) 估值卡片" % (c.get("name", "-"), c.get("code", "-")))
    L.append("")
    # Conclusion first, evidence after: this is the only part a non-technical
    # reader has to understand.
    L += build_summary(params, rb, compat, gaps, cleanup)
    L.append("## 数据完整性详情")
    L.append("")
    if verdict != "full":
        # The template swallows missing data into zeros, so the card has to say
        # out loud which conclusions are not trustworthy.
        L.append("> **数据完整性：%s**" % VERDICT_CN.get(verdict, verdict))
        L.append(">")
        if not compat.get("dcf_reliable", True):
            L.append("> - 绝对估值 DCF **不可信**：%d 个驱动 DCF 的派生假设归零，"
                     "有效税率会变 0，符号可能翻转。"
                     % len(compat.get("dcf_critical_collapsed", [])))
        if not compat.get("rel_reliable", True):
            L.append("> - 相对估值 **不可信**。")
        if compat.get("missing_reports"):
            need = "、".join("%d 年" % m["year"] for m in compat["missing_reports"])
            L.append("> - 缺年份：%s，需补年报后才能出正式结论。" % need)
        if gaps.get("blocking_count"):
            L.append("> - 还差 %d 类输入，最小补法见 `gaps.md`。" % gaps["blocking_count"])
        L.append(">")
        if rb.get("incomplete_history"):
            L.append("> **本工作簿仅供人工核对事实，不可用于估值。** 文件名带 "
                     "`待补数据`。")
            L.append(">")
        L.append("> 详见 `compat.md`。")
        L.append("")
    L.append("| 项 | 值 |")
    L.append("|---|---|")
    L.append("| 报告期 | %s |" % c.get("report_period"))
    L.append("| 估值日期 | %s |" % c.get("valuation_date"))
    L.append("| 历史年度 | %s |" % ", ".join(str(v) for v in sorted(rb.get("years", {}).values())))
    L.append("")
    L.append("## 估值结果")
    L.append("")
    L.append("| 指标 | 值 |")
    L.append("|---|---|")
    L.append("| 绝对估值 DCF (百万元) | %s |" % fmt(d.get("绝对估值!D40")))
    L.append("| 相对估值 PE 法 每股价值 (元) | %s |" % fmt(d.get("相对估值!D6")))
    L.append("| 相对估值 PS 法 每股价值 (元) | %s |" % fmt(d.get("相对估值!D25")))
    L.append("")
    L.append("## 关键输入（历史，百万元）")
    L.append("")
    L.append("| 指标 | 最新年度 |")
    L.append("|---|---|")
    L.append("| 营业收入 | %s |" % fmt(d.get("利润表!H2"), 2))
    L.append("| 营业成本 | %s |" % fmt(d.get("利润表!H3"), 2))
    L.append("")
    q = rb.get("quarterly") or {}
    L.append("## 报告期历史数据（季度）")
    L.append("")
    if q.get("columns"):
        L.append("模板里按季度排列的那张表已填 %d 列：%s。"
                 % (len(q["columns"]),
                    "、".join("%s（%s 列）" % (c["period"], c["col"])
                              for c in q["columns"])))
        L.append("")
        L.append("`报告期利润变化` 的折线取自这张表，现在有数据可画。")
        short = _quarterly_shortfall_lines(q)
        if short:
            L.append("")
            L.append("**下面这几个季度列只填了一部分：**")
            L.append("")
            L.extend(short)
        have = {c["period"] for c in q["columns"]}
        miss = [w for w in (q.get("window") or []) if w not in have]
        if miss:
            L.append("")
            L.append("还缺 %d 个季度（%s），那几段没有点。" % (len(miss), "、".join(miss)))
    else:
        L.append("**这张表是空的。** `报告期利润变化` 的折线取自它，所以图上没有线。"
                 "这批 PDF 里没有季报/半年报，或报告期不在模板的季度窗口内。")
    L.append("")
    groups = {g["id"]: g for g in gaps.get("groups", [])}
    L.append("## 待人工确认的预测假设")
    L.append("")
    _an = groups.get("analyst", {})
    _items = _an.get("items", [])
    _open = _an.get("open", len(_items))
    _done = _an.get("filled", 0)
    if _done and not _open:
        L.append("这 %d 格预测假设**已经填上（来自运行参数的 `analyst` 一节）**。"
                 "如果这些是你自己定的，下面的结果就有意义；如果是别人代填的，先核对这六个数。" % _done)
    else:
        L.append("模板把这 %d 格留空等分析师判断（已填 %d，还空 %d）。"
                 "**留空不等于安全**：Excel 把空白当 0，"
                 "营业成本率空着就是毛利率 100%%，模型不报错。"
                 % (len(_items), _done, _open))
    L.append("")
    L.append("**怎么填（二选一）**：")
    L.append("")
    L.append("1. 在 Excel 里直接填：打开 `假设` 表，B 列是项目名，"
             "I 列是 2026E、J 列是 2027E。")
    L.append("2. 不想打开 Excel：在运行参数 JSON 里加一节 `analyst`，按**项目名**填，"
             "例如 `\"analyst\": {\"营业收入增长率(2026E)\": 0.03}`，重跑一次程序会替你填好。")
    L.append("")
    L.append("| 假设 | 位置 | 当前值 | 状态 | 留空会怎样 |")
    L.append("|---|---|---|---|---|")
    for it in groups.get("analyst", {}).get("items", []):
        cell = it["where"]
        v = d.get(cell)
        filled = v is not None and v != ""
        L.append("| %s | `%s` | %s | %s | %s |"
                 % (it["what"], cell, fmt(v), "已填" if filled else "**空白待填**",
                    it["if_blank"]))
    tax = d.get("假设!I19")
    L.append("| 有效税率（历史派生，非人工输入） | `假设!I19` | %s | %s | 历史不足时归零，"
             "DCF 符号可能翻转 |" % (fmt(tax), "有值" if tax else "**为 0**"))
    L.append("")
    ph = gaps.get("placeholders", {})
    # Only show cells the readback actually read; printing "-" for a cell the
    # template filled would read as blank, which is the opposite of the point.
    # I19 is the effective tax rate; it already has its own row above, and
    # listing it twice would read as two separate findings.
    shown = [c for c in ph.get("cells", [])
             if ("假设!" + c) in d and c != "I19"]
    if ph.get("count") and shown:
        L.append("另有 %d 格是**模板自动占位值**（`IF(ISNUMBER(AVERAGE(历史)), ..., 0)`），"
                 "看着像做好的预测，其实是公式，不是分析师的判断：" % ph["count"])
        L.append("")
        L.append("| 位置 | 当前值 | 性质 |")
        L.append("|---|---|---|")
        for cell in shown[:6]:
            L.append("| `假设!%s` | %s | 模板公式（非人工输入） |" % (cell, fmt(d.get("假设!" + cell))))
        L.append("")
        L.append("全部 %d 格见 `gaps.md`。" % ph["count"])
        L.append("")
    L.append("## A′ 类（硬编码事实，已随公司更新）")
    L.append("")
    L.append("| 科目 | 值（百万元） |")
    L.append("|---|---|")
    for cell, name in (("假设!I55", "商誉"), ("假设!I68", "合同负债"),
                       ("假设!I79", "递延所得税负债"), ("假设!I81", "预计负债")):
        L.append("| %s | %s |" % (name, fmt(d.get(cell), 2)))
    L.append("")
    L.append("## 数据完整性（L2）")
    L.append("")
    L.append("| 列 | 年度 | 覆盖率 |")
    L.append("|---|---|---|")
    for cv in compat.get("coverage", []):
        L.append("| %s | %s | %.1f%% |" % (cv["col"], cv["year"], cv["pct"]))
    L.append("")
    if compat.get("missing_reports"):
        L.append("需补年报：%s"
                 % "；".join("%d 年（%s）" % (m["year"], " 或 ".join(m["report_periods"]))
                             for m in compat["missing_reports"]))
        L.append("")
    L.append("## 待补输入清单")
    L.append("")
    L.append("| 缺口 | 是否阻塞 | 项数 |")
    L.append("|---|---|---|")
    for g in gaps.get("groups", []):
        # "open", not len(items): a group whose inputs have since been supplied is
        # not a gap, and a count that ignores that makes the card contradict the
        # table three lines above it.
        n = g.get("open", len(g["items"]))
        L.append("| %s | %s | %d |" % (g["title"], "**是**" if g["blocking"] else "否", n))
    L.append("")
    rep_ = groups.get("reports", {})
    if rep_.get("minimal_note"):
        L.append(rep_["minimal_note"])
        L.append("")
    L.append("补什么、解锁什么：")
    L.append("")
    for u in gaps.get("unlocks", []):
        L.append("- %s → %s" % (u["action"], u["effect"]))
    L.append("")
    L.append("详见 `gaps.md`。")
    L.append("")
    L.append("## 校验")
    L.append("")
    L.append("- 写入前（数据侧）：%s" % ("通过" if pre.get("passed") else "**未通过**"))
    L.append("- 写入后（结果侧）：%s" % ("通过" if post.get("passed") else "**未通过**"))
    L.append("- 模板保真：%s" % ("通过" if rb.get("fidelity_ok") else "**变更**"))
    return "\n".join(L)


def build_manifest(params, rb, pre, post, extra, compat, gaps):
    L = ["# 运行清单", ""]
    L.append("| 项 | 值 |")
    L.append("|---|---|")
    L.append("| 运行时间 | %s |" % datetime.datetime.now().isoformat(timespec="seconds"))
    L.append("| L2 判定 | %s |" % VERDICT_CN.get(compat.get("verdict"), compat.get("verdict")))
    L.append("| 派生假设归零 | %d / %d |"
             % (compat.get("assumptions_collapsed", 0), compat.get("assumptions_total", 0)))
    L.append("| 派生假设仅被稀释 | %d |" % compat.get("assumptions_diluted", 0))
    L.append("| 模板 sha256（写入前） | `%s` |" % rb.get("sha256_before"))
    L.append("| 模板 sha256（写入后） | `%s` |" % rb.get("sha256_after"))
    L.append("| 映射表版本 | %s |" % extra.get("mapping_version"))
    L.append("")
    if compat.get("must_have_rows"):
        L.append("## 必须齐的关键行")
        L.append("")
        L.append("D:H 任一列缺了这些行，对应假设就会归零（逐格实测）。")
        L.append("")
        L.append("| 行 | 科目 | 缺失打崩的假设数 |")
        L.append("|---|---|---|")
        for m in compat["must_have_rows"]:
            L.append("| %d | %s | %d |" % (m["row"], m["label"] or "-", m["breaks"]))
        L.append("")

    L.append("## 源文件")
    L.append("")
    L.append("| PDF | 报告期 | sha256 |")
    L.append("|---|---|---|")
    for r in params.get("reports", []):
        L.append("| %s | %s | `%s` |" % (os.path.basename(r["pdf"]), r["period"], extra.get("pdf_sha", {}).get(r["pdf"], "-")))
    L.append("")
    L.append("## 写入摘要")
    L.append("")
    for line in rb.get("log", []):
        L.append("- %s" % line)
    L.append("")
    L.append("## 需补年报（L2）")
    L.append("")
    if compat.get("missing_reports"):
        for m in compat["missing_reports"]:
            L.append("- %d 年：报告期 %s" % (m["year"], " 或 ".join(m["report_periods"])))
    else:
        L.append("- 无")
    L.append("")
    L.append("## 待补输入清单")
    L.append("")
    for g in gaps.get("groups", []):
        L.append("- %s：%s，%d 项"
                 % (g["title"], "阻塞" if g["blocking"] else "不阻塞", len(g["items"])))
    if gaps.get("placeholders", {}).get("count"):
        L.append("- 模板自动占位值：%d 格（不是判断，不是缺口）"
                 % gaps["placeholders"]["count"])
    for u in gaps.get("unlocks", []):
        L.append("- %s → %s" % (u["action"], u["effect"]))
    L.append("")
    L.append("## 写入前断言")
    L.append("")
    for r in pre.get("results", []):
        L.append("- [%s] %s" % ("PASS" if r.get("passed") else "FAIL", r.get("id")))
    L.append("")
    L.append("## 写入后断言")
    L.append("")
    for p in post.get("post_assertions", []):
        L.append("- [%s] %s" % ("PASS" if p.get("passed") else "FAIL", p.get("id")))
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--params", required=True)
    ap.add_argument("--readback")
    ap.add_argument("--validate-pre")
    ap.add_argument("--out-card")
    ap.add_argument("--out-card-html", help="the same card, for a browser")
    ap.add_argument("--version", help="skill version, recorded on the card")
    ap.add_argument("--out-manifest")
    ap.add_argument("--out-note", help="write the delivery note (provenance) here")
    ap.add_argument("--summary-only", action="store_true",
                    help="print the plain-language verdict and exit; used when the run "
                         "stops before a workbook exists")
    ap.add_argument("--compat", help="L2 compatibility report")
    ap.add_argument("--gaps", help="to-do list of missing inputs")
    ap.add_argument("--cleanup", help="cleanup list, so the card can name the sheets left blank")
    a = ap.parse_args()
    compat = json.load(open(a.compat, encoding="utf-8")) if a.compat else {}
    gaps = json.load(open(a.gaps, encoding="utf-8")) if a.gaps else {}
    cleanup = json.load(open(a.cleanup, encoding="utf-8")) if a.cleanup else {}
    params = json.load(open(a.params, encoding="utf-8"))

    if a.summary_only:
        text = "\n".join(build_summary(params, {}, compat, gaps, cleanup)).strip()
        print("========== 交给使用者的结论 ==========")
        print(text)
        # Written as well as printed: a run that stops before a workbook exists is
        # exactly when the user needs the one file they were told to open.
        if a.out_card:
            import human_card
            model = _summary_model(params, compat, gaps, cleanup, a.version)
            io.open(a.out_card + ".md", "w", encoding="utf-8",
                    newline="\n").write(human_card.to_markdown(model))
            if a.out_card_html:
                io.open(a.out_card_html, "w", encoding="utf-8",
                        newline="\n").write(human_card.to_html(model))
            print("\ncard    ->", a.out_card + ".md")
        return

    for flag, value in (("--readback", a.readback), ("--validate-pre", a.validate_pre),
                        ("--out-card", a.out_card), ("--out-manifest", a.out_manifest)):
        if not value:
            ap.error("%s is required unless --summary-only is given" % flag)

    rb = json.load(open(a.readback, encoding="utf-8"))
    pre = json.load(open(a.validate_pre, encoding="utf-8"))
    rb["years"] = rb.get("years", {})
    rb["fidelity_ok"] = all(v["same"] for v in rb.get("fidelity", {}).values())
    post = {"passed": rb.get("passed"), "post_assertions": rb.get("post_assertions", [])}
    extra = {"mapping_version": params.get("_mapping_version", "1.0.0"),
             "pdf_sha": params.get("_pdf_sha", {})}
    import human_card
    # "Verified" means the workbook is still the original engine. It is NOT
    # rb["passed"]: a facts_only run is *expected* to fail derived_nonzero, and
    # reading that as "verification failed" would put a red warning on every
    # correctly-degraded deliverable.
    post_ok = {x.get("id"): x.get("passed") for x in post.get("post_assertions", [])}
    verified = bool(rb.get("fidelity_ok")) and post_ok.get("formula_set_unchanged", True)
    model = human_card.build(params, rb, pre, post, compat, gaps, cleanup, a.version,
                             verified=verified)
    io.open(a.out_card + ".md", "w", encoding="utf-8", newline="\n").write(
        human_card.to_markdown(model))
    if a.out_card_html:
        io.open(a.out_card_html, "w", encoding="utf-8", newline="\n").write(
            human_card.to_html(model))
    open(a.out_manifest, "w", encoding="utf-8").write(
        build_manifest(params, rb, pre, post, extra, compat, gaps))
    if a.out_note:
        io.open(a.out_note, "w", encoding="utf-8", newline="\n").write(
            build_delivery_note(params, rb, compat, gaps, cleanup))
        print("note    ->", a.out_note)
    print("card    ->", a.out_card + ".md")
    if a.out_card_html:
        print("card    ->", a.out_card_html)
    print("manifest->", a.out_manifest)
    print("")
    print("========== 交给使用者的结论 ==========")
    print("\n".join(build_summary(params, rb, compat, gaps, cleanup)).strip())


if __name__ == "__main__":
    main()

