"""[3c] Gap list: what is still missing, and what filling it unlocks.

The compatibility layer says whether the data is enough. This one says what to
do about it, and the audience is different: compat.md is a verdict for whoever
called the pipeline, gaps.md is a to-do list for the analyst who has to go find
another annual report or a market multiple.

Three kinds of gap, and they are not interchangeable:

  reports  A year of history no PDF covers. Excel averages the missing column
           into 0, the derived assumption collapses, and the DCF can flip sign
           without raising an error anywhere.
  market   Historical PE/PS/PB/EV-EBITDA/PCF. These live in a market data feed,
           not in an annual report. Missing ones make 相对估值 return #DIV/0!.
  analyst  The six cells a human has to judge. Leaving them blank is not an
           error: Excel reads blank as 0, so a blank cost ratio silently becomes
           a 100% gross margin. This is the gap that looks filled.

A fourth section is not a gap at all but has to be said out loud: the template
fills 137 cells in 假设!I:L with its own historical averages. They look like
finished forecasts and are not. Reporting them as "to fill" would be wrong, and
saying nothing would be worse.
"""
import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import analyst

VERDICT_CN = {"full": "完整", "degraded": "降级", "facts_only": "只出事实层"}
DRIVE_CN = {"dcf": "绝对估值 DCF", "pe": "相对估值 PE", "ps": "相对估值 PS"}
FIELD_CN = {"PE": "历史 PE", "PS": "历史 PS", "PB": "历史 PB",
            "EV_EBITDA": "历史 EV/EBITDA", "PCF": "历史 PCF"}
FIELD_ORDER = ["PE", "PS", "PB", "EV_EBITDA", "PCF"]
# Why the dividend block says what it says. DDM is the one place where a
# correct answer and a hole look identical in the workbook: both are 0.
BASIS_CN = {"numeric": "年报表格值",
            "explicit_absence": "年报原文写明本期不分红",
            "cross_check_failed": "**交叉校验不通过，已拒绝写入**"}


def load(path):
    return json.load(open(path, encoding="utf-8"))


def drives_cn(drives):
    return "、".join(DRIVE_CN.get(d, d) for d in sorted(drives)) or "（不影响估值结论）"


def minimal_reports(years):
    """Smallest set of annual reports that covers the missing years.

    A report for FY(N) carries the closing balance of N and the opening balance
    of N-1, so it fills two adjacent columns. Taking the newest missing year
    first therefore covers it and its predecessor at once: a gap of 2020+2021
    needs one FY2021 report, not two.
    """
    need, picks = set(years), []
    while need:
        y = max(need)
        picks.append(y)
        need -= {y, y - 1}
    return sorted(picks)


def build(params, fieldmap, compat, cleanup, inputs):
    coverage = compat.get("coverage", [])
    col_of_year = {cv["year"]: cv["col"] for cv in coverage}
    gap_cells = {cv["year"]: cv["total"] - cv["filled"] for cv in coverage}

    # One assumption can be reported as broken for several drives; merge it so a
    # single missing year does not get counted three times in the impact column.
    broken = {}
    for key in ("dcf_critical_collapsed", "pe_critical_collapsed",
                "ps_critical_collapsed", "other_collapsed"):
        for c in compat.get(key, []):
            e = broken.setdefault(c["cell"], {"cell": c["cell"], "label": c["label"],
                                              "drives": set(), "cols": set()})
            e["drives"] |= set(c.get("drives") or [])
            e["cols"] |= {b["col"] for b in c["broken_cols"]}

    groups = []

    # ---- 1. missing annual reports -------------------------------------
    items = []
    for m in compat.get("missing_reports", []):
        col = col_of_year.get(m["year"])
        hit = [e for e in broken.values() if col in e["cols"]]
        items.append({
            "what": "%d 年年度报告" % m["year"],
            "year": m["year"],
            "how": "报告期 %s 任意一份即可（每份年报提供期末+期初两年）"
                   % " 或 ".join(m["report_periods"]),
            "cells": gap_cells.get(m["year"], 0),
            "column": col,
            # A missing column only matters if some derived assumption averages
            # it. The oldest column of the window is not read by any of them,
            # so it shows up as a coverage gap without breaking anything.
            "blocking": bool(hit),
            "breaks": [{"cell": e["cell"], "label": e["label"],
                        "drives": sorted(e["drives"])} for e in
                       sorted(hit, key=lambda x: x["cell"])],
            "unlocks": ("恢复 %d 个派生假设，有效税率与成本费用结构才有值"
                        % len(hit)) if hit else "该列不驱动任何假设，只影响同比展示",
        })
    minimal = minimal_reports([i["year"] for i in items])
    groups.append({
        "id": "reports", "title": "待补年报", "blocking": bool(items),
        "minimal_pdfs": minimal,
        "minimal_note": ("补 %s 各一份即可覆盖全部缺口列（一份年报 = 相邻两年）。"
                         % "、".join("FY%d" % y for y in minimal)) if minimal else "无缺口。",
        "why": ("整列缺失时 Excel 的 IF(ISNUMBER(AVERAGE(...))) 走 else 分支，"
                "派生假设直接变 0。有效税率变 0 会让 DCF 符号翻转，而工作簿不报任何错。"),
        "items": items,
    })

    # ---- 2. missing market data ----------------------------------------
    mk = [u for u in fieldmap.get("unresolved", [])
          if "market" in str(u.get("reason", ""))]
    by_field = {}
    for u in mk:
        by_field.setdefault(u.get("field"), []).append(u)
    items = []
    for f in sorted(by_field, key=lambda x: FIELD_ORDER.index(x) if x in FIELD_ORDER else 99):
        rows = sorted({u["row"] for u in by_field[f]})
        cols = sorted({u["col"] for u in by_field[f]})
        years = sorted({u["year"] for u in by_field[f] if u.get("year")})
        where = "年度数据!%s" % ",".join(str(r) for r in rows)
        if cols and rows:
            where = "年度数据!%s%d:%s%d" % (cols[0], rows[0], cols[-1], rows[-1])
        items.append({
            "what": FIELD_CN.get(f, f),
            "where": where,
            "years": years,
            "count": len(by_field[f]),
            "how": "由 params.market.history 提供，键是 %s"
                   % "、".join(str(y) + "-12-31" for y in years[:2]),
            "unlocks": "相对估值 PE/PS 才有倍数可乘，否则整张表是 #DIV/0!",
        })
    groups.append({
        "id": "market", "title": "待补市场数据（B 类，年报里没有）",
        "blocking": bool(items),
        "why": ("历史 PE/PS/PB/EV-EBITDA/PCF 是二级市场行情数据，任何年报 PDF 里都"
                "没有。年度数据!I127 的 AVERAGE 两个分支同时为空时返回 #DIV/0!，"
                "错误会顺着 相对估值!D5 扩散到整张表。"),
        "items": items,
    })

    # ---- 3. analyst judgement ------------------------------------------
    # params.analyst may already answer some of these. Saying "6 to fill" while
    # the card three lines above shows all six filled is the kind of output that
    # teaches a reader to stop trusting the report, so the count follows reality.
    answered = set()
    if params.get("analyst"):
        answered, _ = analyst.resolve(inputs, params["analyst"])
    items = [{
        "what": "%s（%s）" % (t["label"], t["year"]),
        "where": t["cell"],
        "formula": t["used_by"],
        "if_blank": t["if_blank"],
        "affects": drives_cn(t["drives"]),
        "filled": t["cell"] in answered,
    } for t in inputs.get("true_inputs", [])]
    n_open = sum(1 for i in items if not i["filled"])
    groups.append({
        "id": "analyst", "title": "待分析师填写（C 类判断，非事实）",
        "filled": len(items) - n_open, "open": n_open,
        "blocking": False,
        "why": ("这几格是预测假设，AI 不能替分析师拍板，所以出厂模板留空、清理脚本也保持"
                "空白。危险的是空白不报错：Excel 把空当 0，营业成本率空着等于毛利率 100%。"),
        "items": items,
    })

    # ---- 4. not gaps, but must be said ---------------------------------
    ph = inputs.get("template_placeholders", {})
    notes = [{
        "id": "placeholders",
        "title": "模板自动占位值（不是缺口，也不是判断）",
        "cells": ph.get("cells", []),
        "text": ("假设!I:L 有 %d 格由模板的 IF(ISNUMBER(AVERAGE(历史)), ..., 0) 自动填成"
                 "历史均值，例如 %s。它们看着像已经做好的预测，其实是公式；K/L 列更是从"
                 "I/J 列滚出来的。分析师要改就手动覆盖，不改就是「历史均值外推」。"
                 % (ph.get("count", 0), "、".join(ph.get("cells", [])[:6]))),
    }, {
        "id": "non_inputs",
        "title": "被误当成待填的标题格（已剔除）",
        "text": ("假设!I:L 共 %d 个空白格里，%d 个是章节标题和文字标签（如「II﹑资产负债表假设」"
                 "「流动资产：」「（1）IPO募集资金:」），不是输入。清单已按 "
                 "config/analyst_inputs.json 的白名单剔除，不列入待办。"
                 % (len(inputs.get("non_input_cells", {}).get("cells", [])) + 6 + 16,
                    inputs.get("non_input_cells", {}).get("count", 0))),
    }]

    residue = []
    for a in cleanup.get("actions", []):
        if a.get("action") in ("replace", "blank", "blank_literals"):
            residue.append({
                "tier": a["tier"], "sheet": a["sheet"],
                "range": a.get("range") or ",".join(a.get("cells", []))
                          or "C:H 全部数据行",
                "note": a.get("note", ""),
            })
    notes.append({
        "id": "residue",
        "title": "已清理的出厂残留（%s / %s）"
                 % (cleanup.get("origin_residue", {}).get("company", "-"),
                    cleanup.get("origin_residue", {}).get("code", "-")),
        "items": residue,
        "text": ("样本公司数据必须清干净，否则新公司的工作簿里会静默混进上一家的数字。"
                 "P0 是会影响结论的，P1 是隔离区与装饰性残留。"),
    })

    # ---- 5. the dividend basis: an explanation, not a gap ---------------
    dividend = {"evidence": fieldmap.get("dividend_evidence") or {},
                "blank_rows": fieldmap.get("optional_blank") or []}

    # ---- 4b. reports received and deliberately not used ------------------
    excluded = params.get("_excluded_reports") or []
    if excluded:
        notes.append({
            "id": "excluded",
            "title": "收到但本次未使用的报告（不是年报）",
            "lines": ["%s %s —— %s" % (r.get("kind_cn") or "?", r.get("period") or "-",
                                       os.path.basename(r.get("pdf") or "-"))
                      for r in excluded],
            "text": ("这些文件的数字没有写进 `年度数据`。模型按报告期写数：把一季报或"
                     "半年报当年报用，会把三个月的数写进年度列，而工作簿照样计算、"
                     "不报任何错。它们填的是 `报告期历史数据` 那张季度表——"
                     "模板里唯一按季度排列的表，也是 `报告期利润变化` 那张图的"
                     "唯一数据来源。每列填了多少、缺哪张表，见 `请先看这个.md`。"),
        })

    verdict = compat.get("verdict", "full")
    blocking = [g for g in groups if g["blocking"]]
    unlocks = []
    if any(g["id"] == "reports" for g in blocking):
        n = sum(len(i["breaks"]) for g in groups if g["id"] == "reports" for i in g["items"])
        unlocks.append({"action": "补年报",
                        "effect": "恢复 %d 个被归零的派生假设，DCF 与 PE 才可能可信" % n})
    if any(g["id"] == "market" for g in blocking):
        unlocks.append({"action": "补市场倍数",
                        "effect": "相对估值从 #DIV/0! 恢复"})
    # Only a to-do if something is still unanswered: params.analyst may have
    # supplied all six, and a leftover "go fill these" line contradicts the table
    # directly above it.
    n_analyst_open = next((g.get("open", 0) for g in groups if g["id"] == "analyst"), 0)
    if n_analyst_open:
        unlocks.append({"action": "填 %d 个 C 类假设" % n_analyst_open,
                        "effect": "预测期不再是「零增长 + 零成本」，卡片上的 DCF 才有意义"})
    unlocks.append({"action": "两项都补齐后重跑",
                    "effect": "verdict 有望从 facts_only 升到 full"})

    return {
        "version": "1.0.0",
        "phase": "gaps",
        "verdict": verdict,
        "verdict_cn": VERDICT_CN.get(verdict, verdict),
        "blocking_count": len(blocking),
        "placeholders": {"count": ph.get("count", 0), "cells": ph.get("cells", [])},
        "groups": groups,
        "notes": notes,
        "dividend": dividend,
        "unlocks": unlocks,
        "usable_now": ["extract.json（年报事实）", "fieldmap.json（字段映射）",
                       "compat.md（够不够用）", "gaps.md（还差什么、怎么补）"],
    }


def render(g, params):
    c = params.get("company", {})
    L = []
    L.append("# 待补输入清单 —— %s (%s)" % (c.get("name", "-"), c.get("code", "-")))
    L.append("")
    L.append("> 这份清单回答一件事：**要把这张表变成能出估值的模型，还差什么。**")
    L.append("> 判定在 `compat.md`，能不能交付由 L3 决定；这里只列缺口和补齐后的效果。")
    L.append("")
    L.append("## 结论")
    L.append("")
    L.append("| 项 | 值 |")
    L.append("|---|---|")
    L.append("| 当前判定 | **%s** |" % g["verdict_cn"])
    L.append("| 阻塞项 | %d 类 |" % g["blocking_count"])
    L.append("| 现在就能用 | %s |" % "、".join(g["usable_now"]))
    L.append("")
    for grp in g["groups"]:
        L.append("## %s%s" % (grp["title"], "（阻塞）" if grp["blocking"] else "（不阻塞）"))
        L.append("")
        L.append("> %s" % grp["why"])
        L.append("")
        if grp["id"] == "reports":
            if not grp["items"]:
                L.append("无。")
            else:
                L.append("**最小补法：%s**" % grp.get("minimal_note", ""))
                L.append("")
                L.append("| 缺的年份 | 怎么补 | 影响格子 | 是否阻塞 | 会打崩的假设 | 其中驱动 |")
                L.append("|---|---|---|---|---|")
                for it in grp["items"]:
                    L.append("| %s | %s | %d | %s | %s | %s |"
                             % (it["what"], it["how"], it["cells"],
                                "是" if it["blocking"] else "否（不影响估值）",
                                "、".join(b["label"] or b["cell"] for b in it["breaks"]) or "-",
                                drives_cn({d for b in it["breaks"] for d in b["drives"]})))
        elif grp["id"] == "market":
            if not grp["items"]:
                L.append("无。")
            else:
                L.append("| 科目 | 位置 | 缺哪些年 | 缺几格 | 怎么补 |")
                L.append("|---|---|---|---|---|")
                for it in grp["items"]:
                    L.append("| %s | %s | %s | %d | %s |"
                             % (it["what"], it["where"],
                                "、".join(str(y) for y in it["years"]), it["count"], it["how"]))
        elif grp["id"] == "analyst":
            L.append("| 假设 | 位置 | 状态 | 影响 | 留空会怎样 |")
            L.append("|---|---|---|---|---|")
            for it in grp["items"]:
                L.append("| %s | `%s` | %s | %s | %s |"
                         % (it["what"], it["where"],
                            "**已填**" if it.get("filled") else "**空白待填**",
                            it["affects"], it["if_blank"]))
            L.append("")
            L.append("情景类输入（只在做 IPO／增发情景时才需要）：%s。"
                     % "；".join("%s %s" % (s["label"], s["cell"])
                                 for s in grp.get("scenario_inputs", [])))
        L.append("")
    for n in g["notes"]:
        L.append("## %s" % n["title"])
        L.append("")
        L.append(n["text"])
        L.append("")
        for line in n.get("lines", []):
            L.append("- %s" % line)
        for it in n.get("items", []):
            L.append("- `%s` %s!%s —— %s" % (it["tier"], it["sheet"], it["range"], it["note"]))
        if n.get("items"):
            L.append("")
    dv = g.get("dividend") or {}
    if dv.get("evidence") or dv.get("blank_rows"):
        L.append("## 分红口径：DDM 那一块为什么是 0")
        L.append("")
        L.append("> DDM（股利折现）只认分红。公司不分红，这一块就是 0——这可能是事实，"
                 "也可能是没取到数。下面写清是哪一种，并附年报原文作为依据。")
        L.append("")
        L.append("| 项目 | 判定 | 交叉校验 | 出处（原文片段） |")
        L.append("|---|---|---|---|")
        for field in sorted(dv["evidence"]):
            ev = dv["evidence"][field] or {}
            quote = (ev.get("quote") or "").replace("|", "/")
            if len(quote) > 80:
                quote = quote[:80] + "…"
            chk = "与同页现金分红金额一致" if ev.get("cross_checked") else "—"
            L.append("| %s | %s | %s | 第 %s 页：「%s」 |"
                     % (field, BASIS_CN.get(ev.get("basis"), ev.get("basis", "-")),
                        chk, ev.get("page", "-"), quote))
        L.append("")
        if dv.get("blank_rows"):
            rows_ = sorted({b["row"] for b in dv["blank_rows"]})
            L.append("另外 `年度数据` 第 %s 行（中期分红）**留空**：中期分红只在半年报里披露，"
                     "年报没有。模板出厂时这几行写死 0，会把有中期分红的公司静默抹平，"
                     "所以改成留空——如果该期确实没有中期分红，填 0 即可。"
                     % "、".join(str(r) for r in rows_))
            L.append("")
    L.append("## 补什么，解锁什么")
    L.append("")
    L.append("| 动作 | 效果 |")
    L.append("|---|---|")
    for u in g["unlocks"]:
        L.append("| %s | %s |" % (u["action"], u["effect"]))
    L.append("")
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--params", required=True)
    ap.add_argument("--fieldmap", required=True)
    ap.add_argument("--compat", required=True)
    ap.add_argument("--cleanup", required=True)
    ap.add_argument("--inputs", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--out-report", required=True)
    a = ap.parse_args()
    params = load(a.params)
    g = build(params, load(a.fieldmap), load(a.compat), load(a.cleanup), load(a.inputs))
    # The scenario cells belong to the analyst group in the JSON too, so the
    # markdown and the machine-readable copy say the same thing.
    scen = load(a.inputs).get("scenario_inputs", [])
    for grp in g["groups"]:
        if grp["id"] == "analyst":
            grp["scenario_inputs"] = scen
    json.dump(g, open(a.out, "w", encoding="utf-8"), ensure_ascii=False,
              indent=1, default=str)
    open(a.out_report, "w", encoding="utf-8").write(render(g, params))
    print("gaps    ->", a.out)
    print("gaps.md ->", a.out_report)


if __name__ == "__main__":
    main()
