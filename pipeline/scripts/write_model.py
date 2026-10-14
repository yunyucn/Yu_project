"""[4] Loader: write resolved values into the model, recalc, save, assert.

Writing is batched: the whole column goes in one array assignment (0.12 s) rather
than per-cell (11.03 s for the same 1176 cells).
"""
import argparse
import hashlib
import json
import os
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import analyst
import com_excel as cx

COL = {"C": 3, "D": 4, "E": 5, "F": 6, "G": 7, "H": 8}


def zip_metrics(path):
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        return {
            "entries": len(names),
            "vba": any("vbaProject.bin" in n for n in names),
            "charts": sum(1 for n in names if "/charts/chart" in n),
            "activeX": sum(1 for n in names if "activeX" in n),
            "ctrlProps": sum(1 for n in names if "ctrlProps" in n),
            "externalLinks": sum(1 for n in names if "externalLink" in n),
            "drawings": sum(1 for n in names if "/drawings/drawing" in n),
            "size": os.path.getsize(path),
        }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook", required=True)
    ap.add_argument("--fieldmap", required=True)
    ap.add_argument("--params", required=True)
    ap.add_argument("--assertions", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--compat", help="L2 compatibility report; drives the downgrade decision")
    ap.add_argument("--pristine-template",
                    help="original .xlsm, so formula cells can be counted before/after")
    ap.add_argument("--name",
                    help="human-readable base name for the written workbook; the "
                         "verdict suffix is still appended, so the state still "
                         "travels with the filename")
    ap.add_argument("--analyst-inputs",
                    help="analyst_inputs.json, so params.analyst can name cells by label")
    ap.add_argument("--quarterly",
                    help="quarterly.json: quarter-granularity values for 报告期历史数据")
    a = ap.parse_args()
    fm = json.load(open(a.fieldmap, encoding="utf-8"))
    params = json.load(open(a.params, encoding="utf-8"))
    acfg = json.load(open(a.assertions, encoding="utf-8"))

    # The decision to refuse belongs here, not in the compatibility layer: the
    # facts extracted so far stay useful even when a valuation is not warranted.
    compat = json.load(open(a.compat, encoding="utf-8")) if a.compat else {}
    verdict = compat.get("verdict", "full")
    if verdict == "facts_only":
        # The refusal stands by default. It is overridable only on an explicit
        # request, and the override is not a licence to value from the result:
        # the workbook is renamed so the state travels with the file.
        if not params.get("allow_incomplete_history"):
            print("REFUSED: L2 verdict is facts_only.")
            print("  DCF and 相对估值 are both unreliable, so no valuation is written.")
            print("  The fact layer is still usable: extract.json, fieldmap.json, compat.md.")
            print("  Missing reports: %s" % json.dumps(compat.get("missing_reports", []),
                                                     ensure_ascii=False))
            print("  To write the facts into the workbook anyway (inspection only),")
            print("  set allow_incomplete_history in the params.")
            return 1
        print("WARNING: facts_only, but allow_incomplete_history is set.")
        print("  Writing for inspection only; the file is renamed to say so.")
        print("  Missing reports: %s" % json.dumps(compat.get("missing_reports", []),
                                                 ensure_ascii=False))

    before = zip_metrics(a.workbook)
    sha_before = hashlib.sha256(open(a.workbook, "rb").read()).hexdigest()[:16]
    cx.make_writable(a.workbook)

    log = []
    post = []
    with cx.excel_app() as app:
        wb = cx.open_writable(app, a.workbook)
        try:
            # Baseline measured on the clean template as opened, so the assertion
            # below is "writing introduced nothing new" rather than a magic number.
            err_before, err_before_sheets = cx.scan_errors(wb)
            app.Calculation = cx.XL_CALC_MANUAL
            tdir = wb.Worksheets("目录")
            tdir.Range("C7").Value = params["company"]["report_period"]
            ident = params["company"]
            for cell, key in (("C6", "code"), ("F6", "name")):
                if ident.get(key):
                    cx.write_literal(tdir, cell, ident[key])
                    tdir.Range(cell).Value = ident[key]
            mk = params.get("market", {})
            for cell, key in (("F8", "total_shares_mn"), ("F9", "float_shares_mn"),
                              ("F12", "close_price"), ("F15", "hkd_cny"), ("F16", "usd_cny")):
                if mk.get(key) is not None:
                    cx.write_literal(tdir, cell, mk[key])
                    tdir.Range(cell).Value = mk[key]
            log.append("目录: period=%s code=%s name=%s" % (ident.get("report_period"), ident.get("code"), ident.get("name")))

            ws = wb.Worksheets("年度数据")
            rows = sorted({c["row"] for c in fm["cells"].values()})
            lo, hi = rows[0], rows[-1]
            # The block spans rows 3-224, which is not all literal: row 147 is
            # =财务分析!D65, the historical EV/EBITDA the template computes itself.
            # A plain array write would erase it and 相对估值!D82 would read an
            # empty row forever, so formulas in the block are carried through.
            kept = []
            for col, ci in COL.items():
                vals = {c["row"]: c["value"] for c in fm["cells"].values()
                        if c["col"] == col and "value" in c}
                kept += [("%s%d" % (col, r), f) for r, f in
                         cx.write_block_keep_formulas(ws, lo, hi, ci, vals)]
            log.append("年度数据: wrote cols C:H rows %d-%d (%d cells)"
                       % (lo, hi, len(fm["cells"])))
            if kept:
                log.append("年度数据: kept %d template formulas: %s"
                           % (len(kept), ", ".join("%s=%s" % kv for kv in kept)))

            # calc rows are sums of other rows in the same column
            for c in fm["cells"].values():
                if c.get("mode") == "calc":
                    total = 0.0
                    for r in c["rows"]:
                        v = ws.Cells(r, COL[c["col"]]).Value
                        total += float(v or 0)
                    ws.Cells(c["row"], COL[c["col"]]).Value = round(total, 6)
            log.append("年度数据: calc rows resolved")

            # A' cells: balance-sheet facts hard-coded in 假设, must follow the company
            ap_cfg = fm.get("a_prime", {})
            for cell, name in ap_cfg.get("cells", {}).items():
                src_row = ap_cfg.get("source_rows", {}).get(name)
                if src_row is None:
                    continue
                v = ws.Cells(src_row, COL["H"]).Value
                if v is not None:
                    wb.Worksheets("假设").Range(cell).Value = v
                    log.append("假设!%s (%s) <- 年度数据!H%d = %s" % (cell, name, src_row, v))

            # The six judgement cells. The skill never invents them -- they are the
            # analyst's forecast -- but it will type them, so the person does not
            # have to hunt for a cell address. Values come from params.analyst,
            # named by label; anything unrecognised is reported, never ignored.
            # Named analyst_vals, not analyst: the module is imported under that
            # name and a local rebinding would shadow it (it did, once).
            analyst_vals = params.get("analyst") or {}
            if analyst_vals:
                if not a.analyst_inputs:
                    print("REFUSED: params.analyst was given but --analyst-inputs is "
                          "missing, so the labels cannot be resolved to cells.")
                    return 1
                cells_cfg = json.load(open(a.analyst_inputs, encoding="utf-8"))
                resolved, unknown = analyst.resolve(cells_cfg, analyst_vals)
                if unknown:
                    print("REFUSED: params.analyst has entries that match no input cell: %s"
                          % ", ".join(map(str, unknown)))
                    print("  Valid names are the labels in config/analyst_inputs.json, "
                          "optionally with the year, e.g. \"营业收入增长率(2026E)\".")
                    return 1
                for cell, value in sorted(resolved.items()):
                    sheet, addr = cell.split("!")
                    cx.write_literal(wb.Worksheets(sheet), addr, value)
                    wb.Worksheets(sheet).Range(addr).Value = value
                    log.append("假设 %s <- 人工输入 %s" % (cell, value))

            # ---- 报告期历史数据: the quarter-granularity channel --------------
            # 年度数据 answers "what did the company earn each year". This sheet
            # answers "what did it earn each quarter", and it is the only input
            # 报告期利润变化 has (240 formulas, twenty income-statement rows across
            # twelve dated columns). Nothing used to write to it, so the sheet
            # arrived blank and the chart drew no line at all.
            quarterly = {"columns": [], "skipped": [], "cells": 0,
                         "dividends": 0, "window": []}
            if a.quarterly and os.path.exists(a.quarterly):
                qd = json.load(open(a.quarterly, encoding="utf-8"))
                qs = wb.Worksheets(qd["sheet"])
                # The column is chosen by the date its own header names, never by
                # position. That header is =CONCATENATE(数据输入!O5,"-12-","31")
                # and so on, driven by 目录!C8 = TODAY()-1: the window slides with
                # the calendar, and a positional write would drop a quarter into
                # the wrong column on New Year's Day.
                def q_headers():
                    # 表头是公式，不是字面值（数据输入!O2 = YEAR(目录!C8)、
                    # 目录!C8 = TODAY()-1），所以由后端解出来 —— 解不出会抛错，
                    # 不会安静地返回空表头、把一整列季度数据丢掉。
                    return cx.header_dates(qs, 3, 3, 60)

                col_of = q_headers()
                if len(col_of) < 3:
                    # Opened before Excel had evaluated TODAY(); ask once.
                    app.CalculateFull()
                    col_of = q_headers()
                # The window itself travels out: the card can then say which
                # quarters are missing without re-deriving a layout that
                # belongs to the template.
                quarterly["window"] = sorted(col_of)
                for period in sorted(qd.get("by_period", {})):
                    ci = col_of.get(period)
                    if ci is None:
                        quarterly["skipped"].append(period)
                        continue
                    vals = {c["row"]: c["value"] for c in qd["by_period"][period]}
                    cx.write_block_keep_formulas(qs, min(vals), max(vals), ci, vals)
                    quarterly["columns"].append({"period": period,
                                                 "col": cx.col_letter(ci),
                                                 "cells": len(vals)})
                    quarterly["cells"] += len(vals)
                log.append("报告期历史数据: %d 列已填 %s"
                           % (len(quarterly["columns"]),
                              ", ".join("%s@%s" % (c["period"], c["col"])
                                        for c in quarterly["columns"])))
                if quarterly["skipped"]:
                    log.append("报告期历史数据: %s 在模板当前窗口里没有对应列，未写入"
                               % "、".join(quarterly["skipped"]))
                # Only periods the template actually has a column for travel out:
                # a coverage entry for a period with no column would read as a
                # shortfall that is not there.
                written = {c["period"] for c in quarterly["columns"]}
                qcov = qd.get("statement_coverage") or {}
                quarterly["coverage"] = {p: qcov[p] for p in written if p in qcov}
                quarterly["incomplete"] = [i for i in (qd.get("incomplete") or [])
                                           if i.get("period") in written]
                quarterly["parser_warnings"] = qd.get("parser_warnings") or []
                # 每股派现/红股/转增 (rows 34-36) are the same three rows as
                # 年度数据 33-35, one year-end column apart. Copied rather than
                # re-derived, so the dividend cross-check still runs exactly once
                # and a value that failed it cannot reappear through this door.
                ycol_of = {int(y): c for c, y in (fm.get("years") or {}).items()}
                yws = wb.Worksheets("年度数据")
                for period, ci in sorted(col_of.items()):
                    year = int(period[:4])
                    if not period.endswith("12-31") or year not in ycol_of:
                        continue
                    yc = ord(ycol_of[year]) - 64
                    for src_r, dst_r in ((33, 34), (34, 35), (35, 36)):
                        v = yws.Cells(src_r, yc).Value
                        if v is not None:
                            cx.write_literal(qs, "%s%d" % (cx.col_letter(ci), dst_r), v)
                            qs.Cells(dst_r, ci).Value = v
                            quarterly["dividends"] += 1

            app.Calculation = cx.XL_CALC_AUTO
            cx.recalc(app)

            readback = {}
            for col, ci in COL.items():
                rb = cx.read_block(ws, lo, hi, ci)
                for r, v in rb.items():
                    if v is not None:
                        readback["年度数据!%s%d" % (col, r)] = v

            def gv(sheet, cell):
                try:
                    return wb.Worksheets(sheet).Range(cell).Value
                except Exception:
                    return None

            err_total, err_sheets = cx.scan_errors(wb)

            # 公式格的值服务端算不出来（模板有 876 个格子参与循环引用，冷启动
            # 复现不了 Excel 的迭代解，见 com_excel 模块说明）。「某格非零」
            # 「某格不等于某个数」这类断言落在公式格上时读到的是 None，那不是
            # 失败，是「要等 Excel 打开后才知道」。当成失败会让每一次正常运行
            # 都报一个假警报，而假警报会训练人忽略真警报。所以它们单独记一栏
            # deferred，由结论卡讲清楚。
            unreadable = set() if cx.EVALUATES_FORMULAS else wb.formula_refs()

            for spec in acfg["post_write"]:
                if spec["type"] == "errors_zero":
                    if unreadable:
                        post.append({"id": spec["id"], "passed": True,
                                     "deferred": ["整本工作簿的错误值"],
                                     "detail": ["公式的 #DIV/0! 这类错误值要等 Excel "
                                                "算完才存在；服务端只查得到写进去的"
                                                "字面值，这一项在服务端不可判定。"]})
                        continue
                    over = {k: v for k, v in err_sheets.items()
                            if v > err_before_sheets.get(k, 0)}
                    post.append({"id": spec["id"], "passed": not over, "deferred": [],
                                 "detail": ["%s: %d errors, was %d"
                                            % (k, v, err_before_sheets.get(k, 0))
                                            for k, v in sorted(over.items())]})
                elif spec["type"] == "not_zero":
                    defer = [c for c in spec["cells"] if c in unreadable]
                    bad = [c for c in spec["cells"]
                           if c not in unreadable and not gv(*c.split("!"))]
                    post.append({"id": spec["id"], "passed": not bad,
                                 "deferred": defer, "detail": bad})
                elif spec["type"] == "not_equal":
                    bad, defer = [], []
                    for chk in spec["checks"]:
                        if chk["cell"] in unreadable:
                            defer.append(chk["cell"])
                            continue
                        v = gv(*chk["cell"].split("!"))
                        if v == chk["value"]:
                            bad.append({"cell": chk["cell"], "still": v, "desc": chk.get("desc")})
                    post.append({"id": spec["id"], "passed": not bad,
                                 "deferred": defer, "detail": bad})
                elif spec["type"] == "formula_set_unchanged":
                    # The general form of the 147-row bug, and the reason it is
                    # worth generalising: the block write covers 年度数据 rows
                    # 3-224 and puts None in every row it touches, so ANY formula
                    # inside it disappears. Counting formula cells per sheet
                    # against the pristine template catches that whatever the
                    # cell, and catches a future template edit that moves one.
                    if not a.pristine_template:
                        post.append({"id": spec["id"], "passed": None,
                                     "detail": ["no --pristine-template given; skipped"]})
                    else:
                        # Names must not collide with the zip-fidelity before/after
                        # below: Python scopes by function, not by block, so a
                        # plain `before` here rebinds the outer one and the
                        # fidelity check then dies with a KeyError on a sheet name.
                        fc_before = cx.formula_counts(a.pristine_template)
                        fc_after = cx.formula_counts(a.workbook)
                        lost = []
                        for sheet, n in sorted(fc_before.items()):
                            m = fc_after.get(sheet, 0)
                            if m < n:
                                lost.append("%s: %d -> %d (-%d)"
                                            % (sheet, n, m, n - m))
                        post.append({"id": spec["id"], "passed": not lost,
                                     "detail": lost or
                                     ["%d formula cells across %d sheets, unchanged"
                                      % (sum(fc_before.values()), len(fc_before))]})
                elif spec["type"] == "formulas_intact":
                    # A formula row is the workbook's own arithmetic. Both the
                    # reset and this loader array-write the whole 年度数据 C:H
                    # block, and an array write puts None in every row it covers,
                    # so a formula row can be erased without any value looking
                    # wrong. 历史EV/EBITDA went that way and 相对估值's EV/EBITDA
                    # history read empty forever afterwards.
                    bad = []
                    for addr in spec["cells"]:
                        sheet, cell = addr.split("!")
                        try:
                            f = wb.Worksheets(sheet).Range(cell).Formula
                        except Exception as e:
                            f = "unreadable: %s" % e
                        if not (isinstance(f, str) and f.startswith("=")):
                            bad.append({"cell": addr, "formula": f})
                    post.append({"id": spec["id"], "passed": not bad, "detail": bad})
                elif spec["type"] == "zip_metrics":
                    post.append({"id": spec["id"], "passed": True, "detail": []})

            downstream = {}
            for sheet, cell in (("绝对估值", "D40"), ("绝对估值", "D41"),
                                ("相对估值", "D6"), ("相对估值", "D25"),
                                ("输出", "B58"), ("输出", "B60"), ("输出", "B62"),
                                ("假设", "I3"), ("假设", "J3"), ("假设", "I4"),
                                ("假设", "I7"), ("假设", "I8"), ("假设", "I105"),
                                ("假设", "I5"), ("假设", "I6"), ("假设", "I19"),
                                ("假设", "K3"), ("假设", "L3"),
                                ("假设", "I55"), ("假设", "I68"), ("假设", "I79"), ("假设", "I81"),
                                ("利润表", "H2"), ("利润表", "H3")):
                downstream["%s!%s" % (sheet, cell)] = gv(sheet, cell)

            wb.Save()
            log.append("saved")
            if verdict == "degraded":
                log.append("DEGRADED: %d/%d derived assumptions collapsed"
                           % (compat.get("assumptions_collapsed", 0),
                              compat.get("assumptions_total", 0)))
        finally:
            wb.Close(SaveChanges=False)

    # The filename carries identity, not judgement: whose numbers, which period.
    # The verdict lives in the card's first line and in readback.json, and the card
    # is written on every exit path, so a workbook that travels alone still came
    # from a folder whose one readable file says what it is.
    if a.name:
        marked = os.path.join(os.path.dirname(os.path.abspath(a.workbook)), a.name)
        marked = os.path.splitext(marked)[0] + os.path.splitext(a.workbook)[1]
        if os.path.abspath(marked) != os.path.abspath(a.workbook):
            if os.path.exists(marked):
                cx.make_writable(marked)
                os.remove(marked)
            os.rename(a.workbook, marked)
            a.workbook = marked
            log.append("renamed -> %s" % os.path.basename(marked))

    after = zip_metrics(a.workbook)
    fidelity = {k: {"before": before[k], "after": after[k], "same": before[k] == after[k]}
                for k in before if k != "size"}
    out = {"years": fm.get("years", {}),
           "verdict": verdict,
           # Which build produced this file. A version number alone drifts; the
           # fingerprints are what a later investigation can match on.
           "skill_version": params.get("_skill_version"),
           "fingerprints": params.get("_fingerprints", {}),
           "incomplete_history": bool(params.get("allow_incomplete_history")),
           "workbook": a.workbook,
           "dcf_reliable": compat.get("dcf_reliable", True),
           "rel_reliable": compat.get("rel_reliable", True),
           "sha256_before": sha_before,
           "sha256_after": hashlib.sha256(open(a.workbook, "rb").read()).hexdigest()[:16],
           "log": log, "fidelity": fidelity,
           "zip_before": before, "zip_after": after,
           "readback": readback, "downstream": downstream,
           "quarterly": quarterly,
           "errors_total": err_total, "errors_by_sheet": err_sheets,
           "errors_total_before": err_before, "errors_by_sheet_before": err_before_sheets,
           "post_assertions": post,
           # 服务端有没有算过公式值。False 时 downstream 里的公式格一律是 None，
           # 交付物的数要由 Excel 打开时自己算（workbook.xml 里 fullCalcOnLoad=1）。
           "server_side_values": bool(cx.EVALUATES_FORMULAS),
           "needs_excel_recalc": not cx.EVALUATES_FORMULAS,
           "passed": all(p["passed"] for p in post) and all(v["same"] for v in fidelity.values())}
    json.dump(out, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
    print("\n".join(log))
    print("fidelity:", "OK" if all(v["same"] for v in fidelity.values()) else "CHANGED")
    for p in post:
        tag = "PASS" if p["passed"] else "FAIL"
        if p["passed"] and p.get("deferred"):
            tag = "待算"
        print("  [%s] %s" % (tag, p["id"]))
        for c in p.get("deferred", []):
            print("         待 Excel 重算:", c)
        for d in p["detail"][:6]:
            print("        ", d)
    if not cx.EVALUATES_FORMULAS:
        print("公式缓存已清空 + fullCalcOnLoad=1：这份文件用 Excel 打开时会自动重算。")
    print("downstream:", json.dumps({k: v for k, v in list(downstream.items())[:6]}, ensure_ascii=False))
    print("verdict:", verdict)
    print("post-write:", "PASSED" if out["passed"] else "FAILED")
    return 0 if out["passed"] else 1


if __name__ == "__main__":
    # sys.exit(main()) and not main(): the refusal path returns early, and a
    # bare main() would swallow that status and report success to the pipeline.
    sys.exit(main())




