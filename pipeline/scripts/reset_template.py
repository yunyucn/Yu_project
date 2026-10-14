"""[0] Resetter: strip the sample company's data from a template copy.

The template ships with a full sample company (科顺股份 / 300737.SZ, FY2017-2022).
Only replacing 年度数据 leaves that company's hard-coded assumptions and balance
sheet items in place, and the model then returns a non-zero but meaningless
valuation without raising any error.
"""
import argparse
import json
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import com_excel as cx

COL = {"C": 3, "D": 4, "E": 5, "F": 6, "G": 7, "H": 8}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("--cleanup", required=True)
    ap.add_argument("--mapping", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--allow-template-mismatch", action="store_true")
    a = ap.parse_args()
    cleanup = json.load(open(a.cleanup, encoding="utf-8"))
    mapping = json.load(open(a.mapping, encoding="utf-8"))
    # Before anything is copied or opened: is this the build the dictionary was
    # calibrated against? Row numbers are the dictionary key, so a different
    # build means every value lands on the wrong row and nothing says so.
    expect = (mapping.get("template") or {}).get("sha256")
    try:
        actual = cx.verify_template(a.template, expect, a.allow_template_mismatch)
        print("template sha256 OK  %s" % actual[:16])
    except RuntimeError as e:
        print("ABORTED: %s" % e)
        return 1
    rows = sorted(r["row"] for r in mapping["rows"])
    lo, hi = rows[0], rows[-1]

    if not a.dry_run:
        if os.path.exists(a.out):
            cx.make_writable(a.out)
            os.remove(a.out)
        shutil.copy2(a.template, a.out)
        cx.make_writable(a.out)

    log = []
    with cx.excel_app() as app:
        wb = cx.open_writable(app, a.out)
        try:
            for act in cleanup["actions"]:
                if act["action"] == "keep":
                    log.append("keep    %s %s" % (act["sheet"], act.get("cells", act.get("range", ""))))
                    continue
                ws = wb.Worksheets(act["sheet"])
                if act["action"] == "replace" and act.get("rows_from"):
                    # Blank the sample company's literals but keep the template's
                    # own formulas in the block: row 147 (历史EV/EBITDA) is
                    # =财务分析!D65, not sample-company data, and clearing it
                    # leaves 相对估值's EV/EBITDA history empty for every issuer.
                    kept = []
                    for c in COL:
                        kept += [("%s%d" % (c, r), f) for r, f in
                                 cx.write_block_keep_formulas(ws, lo, hi, COL[c], {})]
                    log.append("blank   %s rows %d-%d cols C:H (197 data rows), "
                               "kept %d formulas" % (act["sheet"], lo, hi, len(kept)))
                elif act["action"] == "replace":
                    for cell in act["cells"]:
                        ws.Range(cell).ClearContents()
                    log.append("blank   %s %s (will be replaced)" % (act["sheet"], ",".join(act["cells"])))
                elif act["action"] == "blank":
                    for cell in act["cells"]:
                        # Scalar destinations on this template are literals today,
                        # but a formula here would vanish silently under a plain
                        # write. write_literal refuses instead.
                        cx.write_literal(ws, cell, None)
                        ws.Range(cell).ClearContents()
                    log.append("blank   %s %s" % (act["sheet"], ",".join(act["cells"])))
                elif act["action"] == "blank_literals":
                    cx.blank_literals(ws, act["range"])
                    log.append("blankL  %s %s" % (act["sheet"], act["range"]))
            app.CalculateFull()
            wb.Save()
        finally:
            wb.Close(SaveChanges=False)
    print("\n".join(log))
    print("clean template ->", a.out)


if __name__ == "__main__":
    main()
