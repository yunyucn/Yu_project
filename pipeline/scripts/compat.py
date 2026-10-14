"""[3b] Compatibility layer (L2): facts vs. what the template needs.

Sits between the fact layer (parse + map) and the instance layer (write). It
answers "is this enough, and what is still missing", and it never aborts.

The refusal to produce a valuation belongs to the instance layer, which reads
this report. Aborting here would also throw away the fact layer, which is a
usable deliverable on its own: one annual report yields two years of facts and
supports year-on-year work whether or not five years exist.

Why this exists: 年度数据 D:H feeds 30 derived assumptions through
IF(ISNUMBER(AVERAGE(...)), ..., 0). A missing year makes that average error out,
the assumption collapses to 0, the effective tax rate goes to zero, and the DCF
flips sign with no error raised anywhere. Each assumption needs a different
number of years, so "not enough data" must be reported per assumption rather
than as one blunt threshold.

A missing year and a missing cell are different failures. A whole column gone is
a coverage gap; one cell gone inside an otherwise complete column is a mapping
failure, and whether it is fatal depends on whether that cell is a divisor.
Measured per cell: 24 of the 214 referenced cells collapse an assumption, 55 do
nothing at all, the rest only dilute. The fatal ones are the denominators -
revenue, pretax profit, net profit, attributable net profit, investment property
and paid-in capital - which is why 营业收入 is the one row that must never be
missing.
"""
import argparse
import json

COLS = ["C", "D", "E", "F", "G", "H"]
VERDICT_CN = {"full": "完整", "degraded": "降级", "facts_only": "只出事实层"}
DRIVE_CN = {"dcf": "绝对估值 DCF", "pe": "相对估值 PE", "ps": "相对估值 PS"}
MARKET_REASON = "market data not supplied (B 类：不在年报 PDF 中)"
ACTION = {
    "no PDF line matched":
        "把 PDF 里的实际写法补进 `config/mapping.json` 对应行的 `match` 列表",
    "const value not supplied":
        "分红数据缺失：确认该期报告是否披露利润分配方案，或改用上一期数值",
    MARKET_REASON:
        "在 `params.market.history` 里补齐该年度的估值倍数",
}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--fieldmap", required=True)
    ap.add_argument("--mapping", required=True)
    ap.add_argument("--deps", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--out-report", required=True)
    a = ap.parse_args()
    fm = json.load(open(a.fieldmap, encoding="utf-8"))
    mapping = json.load(open(a.mapping, encoding="utf-8"))
    deps = json.load(open(a.deps, encoding="utf-8"))

    # Formula rows are computed by the workbook itself. Counting them as cells
    # we failed to fill would understate every column's coverage by a constant.
    rows = sorted(r["row"] for r in mapping["rows"] if r.get("mode") != "formula")
    years = fm["years"]
    filled = {(c["row"], c["col"]) for c in fm["cells"].values()}

    coverage = []
    for col in COLS:
        have = sum(1 for r in rows if (r, col) in filled)
        coverage.append({"col": col, "year": years[col], "filled": have,
                         "total": len(rows),
                         "pct": round(100.0 * have / len(rows), 1)})

    gaps = fm.get("coverage_gaps", [])
    gap_years = sorted({g["year"] for g in gaps})
    missing_reports = [{"year": y,
                        "report_periods": ["%d-12-31" % y, "%d-12-31" % (y + 1)]}
                       for y in gap_years]

    unresolved = fm.get("unresolved", [])
    absent_zeroed = fm.get("absent_zeroed", [])
    # Market multiples are missing input, not a dictionary defect. They get their
    # own section, so keep them out of the mapping-failure count and table.
    failures = [u for u in unresolved
                if u.get("reason") != MARKET_REASON]
    by_reason = {}
    for u in failures:
        by_reason.setdefault(u.get("reason", "?"), []).append(u)

    # Both thresholds are measured, not inferred, and they answer different
    # questions. collapse_on_missing comes from clearing one column at a time
    # (a coverage gap); collapse_cells comes from clearing one cell at a time
    # (a mapping failure). The wrapper IF(ISNUMBER(AVERAGE(...)), ..., 0) hides
    # the difference: an unguarded divisor turns one blank into #DIV/0! and the
    # assumption becomes 0, while a blank addend only drags the average down.
    # Of 214 referenced cells, 24 are fatal and 55 have no effect at all.
    collapsed = []
    diluted = []
    for spec in deps["assumptions"]:
        hard, soft = [], []
        for col in spec.get("collapse_on_missing", []):
            need = spec["col_rows"].get(col, [])
            miss = [r for r in need if (r, col) not in filled]
            if not miss:
                continue
            entry = {"col": col, "year": years.get(col), "missing_rows": miss[:8],
                     "count": len(miss)}
            if len(miss) == len(need):
                # Whole column absent (a coverage gap). The measured column-level
                # result applies, including the one assumption that averages a
                # range and therefore only dies when every cell of it is gone.
                hard.append(dict(entry, reason="整列缺失"))
                continue
            # Partial gap (a mapping failure). Only the cells measured to act as
            # divisors collapse the assumption; the rest just add a 0 and dilute
            # the average. Clearing a whole column cannot reveal this: it hides
            # which single cell was load-bearing.
            fatal = sorted(set(miss) & set(spec.get("collapse_cells", {}).get(col, [])))
            if fatal:
                hard.append(dict(entry, reason="缺关键行",
                                 missing_rows=fatal[:8], count=len(fatal)))
            else:
                soft.append(dict(entry, reason="仅稀释"))
        base = {"cell": spec["cell"], "label": spec["label"], "drives": spec["drives"],
                "required": len(spec["required_cols"])}
        if hard:
            collapsed.append(dict(base, broken_cols=hard))
        elif soft:
            diluted.append(dict(base, broken_cols=soft))

    row_label = {r["row"]: (r.get("label") or "") for r in mapping["rows"]}
    must_have_rows = [
        {"row": int(row), "label": row_label.get(int(row), ""),
         "breaks": len(victims), "assumptions": victims}
        for row, victims in deps.get("fatal_rows", {}).items()
        if int(row) in row_label]
    must_have_rows.sort(key=lambda x: (-x["breaks"], x["row"]))

    def driven(kind):
        return [c for c in collapsed if kind in c["drives"]]

    dcf_broken = driven("dcf")
    pe_broken = driven("pe")
    ps_broken = driven("ps")
    other_broken = [c for c in collapsed if not c["drives"]]
    market_missing = [u for u in unresolved if u.get("reason") == MARKET_REASON]

    dcf_reliable = not dcf_broken
    rel_reliable = not pe_broken and not ps_broken and not market_missing
    if not dcf_reliable and not rel_reliable:
        verdict = "facts_only"
    elif dcf_reliable and rel_reliable and not gaps and not failures:
        verdict = "full"
    else:
        verdict = "degraded"

    out = {"phase": "compat", "verdict": verdict,
           "dcf_reliable": dcf_reliable, "rel_reliable": rel_reliable,
           "coverage": coverage, "coverage_gap_years": gap_years,
           "missing_reports": missing_reports,
           "unresolved_count": len(failures),
           "absent_zeroed_count": len(absent_zeroed),
           "unresolved_by_reason": {k: len(v) for k, v in by_reason.items()},
           "assumptions_total": len(deps["assumptions"]),
           "assumptions_collapsed": len(collapsed),
           "assumptions_diluted": len(diluted),
           "diluted": diluted,
           "must_have_rows": must_have_rows,
           "dcf_critical_collapsed": dcf_broken,
           "pe_critical_collapsed": pe_broken,
           "ps_critical_collapsed": ps_broken,
           "other_collapsed": other_broken,
           "market_missing_count": len(market_missing)}
    json.dump(out, open(a.out, "w", encoding="utf-8"), ensure_ascii=False,
              indent=1, default=str)

    L = ["# L2 兼容性报告", ""]
    L.append("> 本报告只描述「够不够用」，不做拦截。是否交付由 L3 实例层决定。")
    L.append("")
    L.append("| 判定 | 值 |")
    L.append("|---|---|")
    L.append("| 结论 | **%s** |" % VERDICT_CN[verdict])
    L.append("| 绝对估值 DCF 可信 | %s |" % ("是" if dcf_reliable else "**否**"))
    L.append("| 相对估值可信 | %s |" % ("是" if rel_reliable else "**否**"))
    L.append("| 派生假设确定归零 | %d / %d |"
             % (len(collapsed), len(deps["assumptions"])))
    L.append("| 派生假设仅被稀释（部分行缺失，不致命） | %d |" % len(diluted))
    L.append("")

    L.append("## 年度覆盖")
    L.append("")
    L.append("| 列 | 年度 | 已填 | 缺口 | 覆盖率 |")
    L.append("|---|---|---|---|---|")
    for cv in coverage:
        L.append("| %s | %s | %d | %d | %.1f%% |"
                 % (cv["col"], cv["year"], cv["filled"],
                    cv["total"] - cv["filled"], cv["pct"]))
    L.append("")

    L.append("## 覆盖缺口（该年没有年报，属取数问题）")
    L.append("")
    if not gaps:
        L.append("无。")
    else:
        L.append("缺 %d 格，涉及年份 %s。" % (len(gaps), gap_years))
        L.append("")
        L.append("**需补以下年报（每份提供 2 年）：**")
        L.append("")
        L.append("| 缺的年份 | 可用报告期 |")
        L.append("|---|---|")
        for m in missing_reports:
            L.append("| %d | %s |" % (m["year"], " 或 ".join(m["report_periods"])))
    L.append("")

    L.append("## 完整报表内未列示的科目（记 0，非错误）")
    L.append("")
    if not absent_zeroed:
        L.append("无。")
    else:
        rows_ = sorted({(a["row"], a["label"]) for a in absent_zeroed})
        L.append("报表抓取完整，但公司未列示这些科目 —— A 股报表按标准格式编制，"
                 "余额为零的科目直接不列示，故记 0。共 %d 格、涉及 %d 个科目。"
                 % (len(absent_zeroed), len(rows_)))
        L.append("")
        L.append("| 行 | 科目 |")
        L.append("|---|---|")
        for r, lab in rows_:
            L.append("| %d | %s |" % (r, lab))
    L.append("")

    L.append("## 映射失败（PDF 在手但标签没对上，属字典问题）")
    L.append("")
    if not failures:
        L.append("无。")
    else:
        L.append("共 %d 条：" % len(failures))
        L.append("")
        L.append("| 原因 | 条数 |")
        L.append("|---|---|")
        for k, v in sorted(by_reason.items(), key=lambda kv: -len(kv[1])):
            L.append("| %s | %d |" % (k, len(v)))
        L.append("")
        L.append("处理方式（按原因）：")
        L.append("")
        L.append("| 原因 | 处理 |")
        L.append("|---|---|")
        for k in sorted(by_reason, key=lambda x: -len(by_reason[x])):
            L.append("| %s | %s |" % (k, ACTION.get(k, "见 references/field-mapping.md")))
    L.append("")

    L.append("## 派生假设归零风险")
    L.append("")
    if not collapsed:
        L.append("无。全部 %d 个派生假设所需年份齐全。" % len(deps["assumptions"]))
    else:
        for title, group in (("驱动 DCF 的假设（归零则 DCF 不可信）", dcf_broken),
                             ("驱动相对估值 PE 的假设", pe_broken),
                             ("驱动相对估值 PS 的假设", ps_broken),
                             ("其他假设（不影响估值结论）", other_broken)):
            if not group:
                continue
            L.append("### %s —— %d 个" % (title, len(group)))
            L.append("")
            L.append("| 假设 | 名称 | 需年数 | 断在哪列 |")
            L.append("|---|---|---|---|")
            for c in group:
                cols = "、".join("%s(%s)" % (b["col"], b["year"]) for b in c["broken_cols"])
                L.append("| %s | %s | %d | %s |"
                         % (c["cell"], c["label"] or "-", c["required"], cols))
            L.append("")

    if must_have_rows:
        L.append("### 必须齐的关键行（D:H 任一列缺了就打崩假设，逐格实测）")
        L.append("")
        L.append("| 行 | 科目 | 缺失会打崩的假设数 | 其中驱动 DCF/PE |")
        L.append("|---|---|---|---|")
        by_cell = {c["cell"]: c for c in deps["assumptions"]}
        for m in must_have_rows:
            crit = sum(1 for a in m["assumptions"]
                       if by_cell.get(a, {}).get("drives"))
            L.append("| %d | %s | %d | %d |"
                     % (m["row"], m["label"] or "-", m["breaks"], crit))
        L.append("")

    if diluted:
        L.append("### 仅被稀释（部分行缺失，实测不归零）—— %d 个" % len(diluted))
        L.append("")
        L.append("缺失格只是作为加数计入平均，值会偏低但不会归零。可以交付，但要知道偏低。")
        L.append("")
        L.append("| 假设 | 名称 | 缺在哪列 |")
        L.append("|---|---|---|")
        for c in diluted:
            cols = "、".join("%s(%s)" % (b["col"], b["year"]) for b in c["broken_cols"])
            L.append("| %s | %s | %s |" % (c["cell"], c["label"] or "-", cols))
        L.append("")

    L.append("## 市场数据（B 类，年报里没有）")
    L.append("")
    if not market_missing:
        L.append("已提供完整。")
    else:
        L.append("缺 %d 格。历史 PE/PS/PB/EV-EBITDA/PCF 必须由 `params.market.history` 提供；"
                 "缺失会让 `相对估值` 出 `#DIV/0!` 而工作簿不报错。" % len(market_missing))
    L.append("")

    L.append("## 结论")
    L.append("")
    if verdict == "full":
        L.append("数据完整，可出正式 xlsm。")
    elif verdict == "degraded":
        L.append("**降级交付**：")
        if not dcf_reliable:
            L.append("- 绝对估值 DCF **不可信**（%d 个关键假设归零，有效税率会变 0，符号可能翻转）。"
                     % len(dcf_broken))
        if not rel_reliable:
            L.append("- 相对估值**不可信**。")
        L.append("- 交付物文件名与卡片首行会带降级标记。")
    else:
        L.append("**只出事实层，不出估值**：DCF 与相对估值均不可信。")
        L.append("事实表本身仍可用（同比、可比公司集）。")
    open(a.out_report, "w", encoding="utf-8").write("\n".join(L) + "\n")

    print("verdict: %s (dcf_reliable=%s rel_reliable=%s)"
          % (verdict, dcf_reliable, rel_reliable))
    print("coverage gaps: %d cells over years %s" % (len(gaps), gap_years))
    print("mapping failures: %d" % len(failures))
    print("absent->0: %d" % len(absent_zeroed))
    print("assumptions collapsed: %d / %d (dcf-critical %d), diluted: %d"
          % (len(collapsed), len(deps["assumptions"]), len(dcf_broken), len(diluted)))
    print("report ->", a.out_report)


if __name__ == "__main__":
    main()

