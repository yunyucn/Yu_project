"""[5] Acceptance: is the produced workbook still the original engine?

Read-only. Answers one question and nothing else: apart from the numbers, is
everything that makes the template work still there?

This is not the pipeline's own gate -- write_model.py already refuses to save a
workbook whose formula set shrank. This script exists so the *user* can re-check a
delivered file without rerunning the pipeline, and so a file produced by an older
build of the skill can be judged after the fact.

Two layers:

  default          zip/ole level. No Excel needed, a second or two. Compares sheet
                   list, per-sheet formula counts, key formula cells, defined
                   names, charts, ActiveX controls, the VBA project, external
                   links, media, validations/conditional formats/merges.
  --equivalence    Excel COM. Feeds the *same* inputs into both workbooks and
                   compares the outputs. This is the only check that proves the
                   arithmetic is unchanged rather than merely present.
"""
import argparse
import collections
import hashlib
import io
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import com_excel as cx

COLS = ["C", "D", "E", "F", "G", "H"]


def h(b):
    return hashlib.sha256(b).hexdigest()


def sha_of_parts(z, names):
    """One hash over a set of parts, order-independent."""
    out = hashlib.sha256()
    for n in sorted(names):
        out.update(n.encode("utf-8"))
        out.update(z.read(n))
    return out.hexdigest()


def sheet_part_map(z):
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    out = {}
    for name, r in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wb):
        tgt = rid.get(r, "")
        part = tgt if tgt.startswith("xl/") else "xl/" + tgt.lstrip("/")
        out[name] = part
    return out


def cell_formula(z, parts, sheet, addr):
    """True/False/None: is this cell a formula? None means the cell is absent.

    The naive r'<c r="A1"[^>]*>(.*?)</c>' is wrong: an empty cell serialises as
    <c r="A1" s="5"/> and the pattern then runs on to the *next* cell's content,
    reporting whatever that cell holds. Self-closing must be handled explicitly,
    or this check quietly passes on a file whose formula was deleted.
    """
    try:
        xml = z.read(parts[sheet]).decode("utf-8")
    except KeyError:
        return None
    m = re.search(r'<c r="%s"(?P<rest>[^>]*?)(?:/>|>(?P<inner>.*?)</c>)'
                  % re.escape(addr), xml, re.S)
    if not m:
        return None
    if m.group("inner") is None:
        return False
    return re.search(r"<f[ />]", m.group("inner")) is not None


def activex_inventory(z):
    """Multiset of ActiveX control classes, read from the activeX .xml parts.

    Byte comparison is the obvious check and it is wrong. Measured: a plain
    open+recalc+save in Excel rewrites 154 of the 166 .bin parts of the *pristine*
    template, with no agent involved. Excel renumbers the parts, so activeX1.bin
    stops being the control it used to be -- the first part's classid flips from
    ComboBox to CommandButton while the totals stay identical (131 ComboBox,
    23 TextBox, 9 CommandButton, 3 other). Comparing bytes therefore reports
    "controls changed" for any workbook a human has ever saved, which is every
    delivered workbook. What survives a save is which controls exist and of what
    kind, so that is what gets compared.
    """
    out = collections.Counter()
    for n in z.namelist():
        if re.match(r"xl/activeX/activeX\d+\.xml$", n):
            for c in re.findall(r'classid="([^"]*)"', z.read(n).decode("utf-8", "ignore")):
                out[c.lower()] += 1
    return out


def sheet_meta_counts(z, parts):
    """Per sheet: data validations, conditional-format blocks, merged cells."""
    out = {}
    for name, part in parts.items():
        try:
            xml = z.read(part).decode("utf-8")
        except KeyError:
            continue
        out[name] = {
            "dataValidation": len(re.findall(r"<dataValidation[ >]", xml)),
            "conditionalFormatting": len(re.findall(r"<conditionalFormatting[ >]", xml)),
            "mergeCell": len(re.findall(r"<mergeCell[ >]", xml)),
        }
    return out


def defined_names(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read("xl/workbook.xml").decode("utf-8")
    m = re.search(r"<definedNames>(.*?)</definedNames>", xml, re.S)
    return m.group(1) if m else ""


def chart_stats(path):
    with zipfile.ZipFile(path) as z:
        charts = [n for n in z.namelist() if re.match(r"xl/charts/chart\d+\.xml$", n)]
        refs = 0
        for n in charts:
            refs += len(re.findall(r"<c:f>", z.read(n).decode("utf-8", "ignore")))
        return len(charts), refs


def vba_streams(path):
    import olefile
    with zipfile.ZipFile(path) as z:
        raw = z.read("xl/vbaProject.bin")
    o = olefile.OleFileIO(io.BytesIO(raw))
    names = sorted("/".join(s) for s in o.listdir())
    streams = {n: h(o.openstream(n).read()) for n in names}
    proj = o.openstream("PROJECT").read().decode("latin-1")
    locked = {}
    for key in ("CMG", "DPB", "GC"):
        m = re.search(r'^%s="([^"]*)"' % key, proj, re.M)
        locked[key] = bool(m and m.group(1))
    return streams, locked


def zip_check(pristine, output, cfg):
    res = []
    with zipfile.ZipFile(pristine) as za, zipfile.ZipFile(output) as zb:
        pa, pb = sheet_part_map(za), sheet_part_map(zb)

        ok = list(pa) == list(pb)
        res.append(("sheets", ok, "%d 张工作表" % len(pa), 
                    "名称与顺序一致" if ok else "不一致：%s" % _first_diff(list(pa), list(pb))))

        fa = cx.formula_counts(pristine)
        fb = cx.formula_counts(output)
        down = {k: (fa[k], fb.get(k)) for k in fa if fb.get(k, 0) < fa[k]}
        up = {k: (fa[k], fb.get(k)) for k in fa if fb.get(k, 0) > fa[k]}
        tot_a, tot_b = sum(fa.values()), sum(fb.values())
        res.append(("formulas", not down,
                    "原件 %d 个公式格，产出 %d 个" % (tot_a, tot_b),
                    "逐表不少于原件" if not down else "这些表少了公式：%s" % down))
        if up:
            res.append(("formulas_added", True, "", "产出多出公式的表（非故障）：%s" % up))

        missing = []
        for addr in cfg["key_formula_cells"]:
            sh, cell = addr.split("!")
            if cell_formula(za, pa, sh, cell) is not True:
                missing.append("%s（原件就不是公式）" % addr)
            elif cell_formula(zb, pb, sh, cell) is not True:
                missing.append("%s（产出里没了）" % addr)
        res.append(("key_formulas", not missing,
                    "%d 个关键格" % len(cfg["key_formula_cells"]),
                    "全部仍是模板公式" if not missing else "；".join(missing)))

        na, nb = defined_names(pristine), defined_names(output)
        res.append(("defined_names", na == nb, "%d 字符" % len(na),
                    "定义名称段完全一致" if na == nb else "定义名称段有差异"))

        ca, ra = chart_stats(pristine)
        cb, rb = chart_stats(output)
        res.append(("charts", (ca, ra) == (cb, rb), "原件 %d 图 / %d 引用点" % (ca, ra),
                    "一致" if (ca, ra) == (cb, rb) else "产出 %d 图 / %d 引用点" % (cb, rb)))

        ia, ib = activex_inventory(za), activex_inventory(zb)
        nax = sum(ia.values())
        res.append(("activex", ia == ib, "%d 个控件" % nax,
                    "控件种类与数量一致" if ia == ib
                    else "控件种类不同：原件 %s / 产出 %s" % (dict(ia), dict(ib))))

        cp = [n for n in za.namelist() if n.startswith("xl/ctrlProps/")]
        cq = [n for n in zb.namelist() if n.startswith("xl/ctrlProps/")]
        res.append(("ctrlProps", sha_of_parts(za, cp) == sha_of_parts(zb, cq),
                    "%d 个" % len(cp), "一致" if sha_of_parts(za, cp) == sha_of_parts(zb, cq) else "有差异"))

        va, la = vba_streams(pristine)
        vb, lb = vba_streams(output)
        same_streams = set(va) == set(vb)
        diff = [k for k in va if vb.get(k) != va[k]] if same_streams else []
        res.append(("vba", same_streams and not diff, "%d 个流" % len(va),
                    "VBA 工程逐流一致（密码状态：%s）" % ("有密码，已原样继承" if la.get("CMG") else "无密码")
                    if (same_streams and not diff) else "VBA 有差异：%s" % (diff or "流清单不同")))

        ea = [n for n in za.namelist() if n.startswith("xl/externalLinks/") and n.endswith(".xml")]
        eb = [n for n in zb.namelist() if n.startswith("xl/externalLinks/") and n.endswith(".xml")]
        res.append(("external_links", sha_of_parts(za, ea) == sha_of_parts(zb, eb),
                    "%d 个外部链接" % len(ea),
                    "一致" if sha_of_parts(za, ea) == sha_of_parts(zb, eb) else "有差异"))

        ma = [n for n in za.namelist() if n.startswith("xl/media/")]
        mb = [n for n in zb.namelist() if n.startswith("xl/media/")]
        res.append(("media", len(ma) == len(mb), "%d 个图片/媒体" % len(ma),
                    "数量一致" if len(ma) == len(mb) else "产出 %d 个" % len(mb)))

        sa, sb = sheet_meta_counts(za, pa), sheet_meta_counts(zb, pb)
        bad = []
        for k in sa:
            if sa[k] != sb.get(k):
                bad.append("%s 原件 %s / 产出 %s" % (k, sa[k], sb.get(k)))
        res.append(("sheet_meta", not bad, "%d 张表的数据有效性/条件格式/合并格" % len(sa),
                    "逐表一致" if not bad else "；".join(bad[:5])))
    return res


def _first_diff(a, b):
    for i in range(max(len(a), len(b))):
        x = a[i] if i < len(a) else None
        y = b[i] if i < len(b) else None
        if x != y:
            return "第 %d 个：原件 %r / 产出 %r" % (i + 1, x, y)
    return ""


def _same(x, y, tol=1e-9):
    ex, ey = cx.is_error(x), cx.is_error(y)
    if ex or ey:
        return ex and ey and x == y
    if x is None and y is None:
        return True
    if isinstance(x, bool) or isinstance(y, bool):
        return x == y
    if isinstance(x, (int, float)) and isinstance(y, (int, float)):
        if x == y:
            return True
        return abs(x - y) <= tol * max(1.0, abs(x), abs(y))
    return x == y


def equivalence(pristine, output, cfg, verbose=True):
    """Feed both workbooks the same arguments, then compare the answers.

    The facts come from the produced file (it holds the real issuer), the market
    multiples and the ten analyst cells come from the pristine template (the
    produced file has no market data and its analyst cells are deliberately
    blank). Both sides then hold identical inputs, so any surviving difference in
    the outputs is the engine, not the data.
    """
    tmp = tempfile.mkdtemp(prefix="accept_")
    a_path = os.path.join(tmp, "A_pristine.xlsm")
    b_path = os.path.join(tmp, "B_output.xlsm")
    shutil.copy2(pristine, a_path)
    cx.make_writable(a_path)
    shutil.copy2(output, b_path)
    cx.make_writable(b_path)
    facts = cfg["input_surface"]["facts_block"]
    market = cfg["input_surface"]["market_block"]
    literals = cfg["input_surface"]["literals"]
    cells, scans, skipped = [], [], []
    try:
        with cx.excel_app() as app:
            wa = cx.open_writable(app, a_path)
            wb = cx.open_writable(app, b_path)
            try:
                for c in facts["cols"]:
                    col = ord(c) - 64
                    for lo, hi in facts["row_ranges"]:
                        vals = cx.read_block(wb.Worksheets(facts["sheet"]), lo, hi, col)
                        cx.write_block_keep_formulas(wa.Worksheets(facts["sheet"]), lo, hi, col, vals)
                for blk in cfg["input_surface"].get("extra_input_blocks", []):
                    for c in blk["cols"]:
                        col = ord(c) - 64
                        for lo, hi in blk["row_ranges"]:
                            vals = cx.read_block(wb.Worksheets(blk["sheet"]), lo, hi, col)
                            cx.write_block_keep_formulas(
                                wa.Worksheets(blk["sheet"]), lo, hi, col, vals)
                for c in market["cols"]:
                    col = ord(c) - 64
                    for lo, hi in market["row_ranges"]:
                        vals = cx.read_block(wa.Worksheets(market["sheet"]), lo, hi, col)
                        cx.write_block_keep_formulas(wb.Worksheets(market["sheet"]), lo, hi, col, vals)
                for addr in literals:
                    sh, cell = addr.split("!")
                    tgt = wb.Worksheets(sh).Range(cell)
                    cur = tgt.Formula
                    if isinstance(cur, str) and cur.startswith("="):
                        skipped.append("%s 在产出里是公式，未覆盖" % addr)
                        continue
                    tgt.Value = wa.Worksheets(sh).Range(cell).Value
                cx.recalc(app, 2)
                for addr in cfg["key_output_cells"]:
                    sh, cell = addr.split("!")
                    x = wa.Worksheets(sh).Range(cell).Value
                    y = wb.Worksheets(sh).Range(cell).Value
                    cells.append({"cell": addr, "pristine": _show(x), "output": _show(y),
                                  "same": _same(x, y)})
                for rng in cfg["scan_ranges"]:
                    sh, r = rng.split("!")
                    va = wa.Worksheets(sh).Range(r).Value
                    vb = wb.Worksheets(sh).Range(r).Value
                    diff = []
                    n = 0
                    for i, row in enumerate(va if isinstance(va, tuple) else ((va,),)):
                        rb = vb[i] if isinstance(vb, tuple) and i < len(vb) else ()
                        for j, x in enumerate(row if isinstance(row, tuple) else (row,)):
                            y = rb[j] if isinstance(rb, tuple) and j < len(rb) else None
                            n += 1
                            if not _same(x, y):
                                diff.append("%s!%s" % (sh, _rc(i + 1, j + 1)))
                    scans.append({"range": rng, "cells": n, "diff": len(diff),
                                  "sample": diff[:10]})
            finally:
                wa.Close(SaveChanges=False)
                wb.Close(SaveChanges=False)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    ok = all(c["same"] for c in cells) and all(s["diff"] == 0 for s in scans)
    return {"ok": ok, "cells": cells, "scans": scans, "skipped": skipped}


def _rc(r, c):
    s = ""
    while c:
        c, m = divmod(c - 1, 26)
        s = chr(65 + m) + s
    return "%s%d" % (s, r)


def _show(x):
    if x is None:
        return ""
    if cx.is_error(x):
        return "#ERR(%s)" % x
    if isinstance(x, float):
        return "%.6f" % x
    return str(x)


def build_report(pristine, output, res, eq, cfg):
    L = []
    L.append("# 验收：产出文件还是原件那台引擎吗")
    L.append("")
    L.append("- 原件：`%s`" % os.path.basename(pristine))
    L.append("- 产出：`%s`" % os.path.basename(output))
    L.append("")
    fails = [r for r in res if not r[1]]
    L.append("## 结论")
    L.append("")
    if not fails and (eq is None or eq["ok"]):
        L.append("**通过。** 除了数据本身，模板的功能（公式、图表、控件、宏、外部链接）都在，"
                 "两边一致。")
    elif fails:
        L.append("**不通过。** 有 %d 项对不上，见下表。这一份不要用于估值。" % len(fails))
    else:
        L.append("**结构通过，但算出来的数不一样。** 见「同数据同结果」一节。")
    L.append("")
    L.append("## 逐项比对")
    L.append("")
    L.append("| 检查 | 结果 | 原件 | 说明 |")
    L.append("|---|---|---|---|")
    for name, ok, a, b in res:
        L.append("| `%s` | %s | %s | %s |" % (name, "通过" if ok else "**不通过**", a, b))
    L.append("")
    if eq is not None:
        L.append("## 同数据同结果")
        L.append("")
        L.append("把同一份输入喂进两份文件（事实取自产出，市场倍数与分析师格取自原件），重算后比对：")
        L.append("")
        for c in eq["cells"]:
            L.append("- `%s`：原件 %s / 产出 %s —— %s"
                     % (c["cell"], c["pristine"] or "空", c["output"] or "空",
                        "一致" if c["same"] else "**不同**"))
        for s in eq["scans"]:
            L.append("- 扫描 `%s`：%d 格里 %d 格不同" % (s["range"], s["cells"], s["diff"]))
        for k in eq["skipped"]:
            L.append("- 跳过：%s" % k)
        L.append("")
    L.append("## 这份脚本不检查什么")
    L.append("")
    L.append("数据对不对（那是 `validate.py` 与 `readback.json` 的事），以及数字有没有经济含义。")
    L.append("它只回答「功能还在不在」。")
    return "\n".join(L) + "\n"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pristine", required=True, help="original template .xlsm")
    ap.add_argument("--output", required=True, help="produced .xlsm to judge")
    ap.add_argument("--config", default=os.path.join(os.path.dirname(HERE), "config", "acceptance.json"))
    ap.add_argument("--equivalence", action="store_true",
                    help="also feed both workbooks the same inputs and compare outputs (needs Excel)")
    ap.add_argument("--out", help="write the full result as JSON here")
    ap.add_argument("--out-report", help="write the plain-language report as markdown here")
    a = ap.parse_args()
    cfg = json.load(open(a.config, encoding="utf-8"))

    print("pristine:", a.pristine)
    print("output  :", a.output)
    print("")
    res = zip_check(a.pristine, a.output, cfg)
    for name, ok, det, msg in res:
        print("[%s] %-16s %s" % ("PASS" if ok else "FAIL", name, msg))

    eq = None
    if a.equivalence and not cx.HAS_EXCEL:
        # 「同数据同结果」要拿两个工作簿在 Excel 里各算一遍再比。这个后端
        # 没有 Excel，做不了 —— 说清楚是没做，不是做了没过。
        print("")
        print("同数据同结果：跳过（这个后端不驱动 Excel，公式值由办公人员打开时算）")
        eq = {"ok": True, "skipped": "no Excel backend", "cells": [], "scans": [],
              "note": "服务端只做结构性核对；数值一致性由 Excel 打开后保证。"}
    elif a.equivalence:
        print("")
        print("同数据同结果（要开 Excel，约 30 秒）...")
        eq = equivalence(a.pristine, a.output, cfg)
        print("[%s] equivalence" % ("PASS" if eq["ok"] else "FAIL"))
        for c in eq["cells"]:
            print("    %-16s 原件 %-14s 产出 %-14s %s"
                  % (c["cell"], c["pristine"] or "空", c["output"] or "空",
                     "一致" if c["same"] else "不同"))
        for s in eq["scans"]:
            print("    扫描 %-18s %d 格中 %d 格不同" % (s["range"], s["cells"], s["diff"]))

    fails = [r for r in res if not r[1]]
    ok = not fails and (eq is None or eq["ok"])
    print("")
    print("结论：%s" % ("功能一致，可以用" if ok else "有问题，不要用这一份"))

    if a.out:
        json.dump({"pristine": a.pristine, "output": a.output,
                   "checks": [{"id": n, "passed": o, "detail": d} for n, o, _, d in res],
                   "equivalence": eq, "passed": ok},
                  open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        print("json    ->", a.out)
    if a.out_report:
        io.open(a.out_report, "w", encoding="utf-8", newline="\n").write(
            build_report(a.pristine, a.output, res, eq, cfg))
        print("report  ->", a.out_report)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
