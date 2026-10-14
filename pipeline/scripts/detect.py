"""[0] Preflight: what is this document, and which columns can it fill?

The run used to trust the `period` typed into params.json. Nothing checked it and
nothing opened the document first, so a first-quarter report handed in as
2025-03-31 was read as FY2025: three months of figures went into the annual column
of 年度数据, the whole valuation chain was built on them, and no error appeared
anywhere. A half-year report does the same and is worse, because its numbers look
like a plausible annual figure.

So the document is asked first. A cover page states what it is --
"2025年年度报告", "2025年半年度报告", "2025年第一季度报告" -- and the period end
follows from the kind, not from the caller's memory. The declared `period` is kept
as a cross-check: when the two disagree the run stops instead of picking one.
"""
import argparse
import io
import json
import os
import re
import sys

import pdfplumber

KIND_CN = {"annual": "年度报告", "half": "半年度报告",
           "q1": "第一季度报告", "q3": "第三季度报告"}
KIND_END = {"annual": "12-31", "half": "06-30", "q1": "03-31", "q3": "09-30"}

# 半年度报告 contains 年度报告 as a substring, so the more specific pattern has to
# be tried as well; the earliest match position in the title block wins.
KIND_PATTERNS = [
    ("half", re.compile(r"(20\d{2})\s*年\s*半年度报告")),
    ("q1", re.compile(r"(20\d{2})\s*年\s*第[一1]\s*季度报告")),
    ("q3", re.compile(r"(20\d{2})\s*年\s*第[三3]\s*季度报告")),
    ("annual", re.compile(r"(20\d{2})\s*年\s*年度报告")),
]

COMPANY = re.compile(r"([\u4e00-\u9fff]{2,20}(?:股份有限公司|有限责任公司|有限公司))")
CODE = re.compile(r"(?:证券代码|股票代码)\s*[：:]\s*(\d{6})")

# The title block is the first thing on the cover. A "半年度报告" mentioned 300
# characters further down is body text about some other filing, not this
# document's identity, so the search is bounded.
HEAD_CHARS = 500


def _page_texts(path, limit=2):
    out = []
    with pdfplumber.open(path) as pdf:
        for i in range(min(limit, len(pdf.pages))):
            out.append(pdf.pages[i].extract_text() or "")
    return out


def detect(path):
    """What this PDF is, according to its own cover page."""
    rec = {"pdf": path, "kind": None, "kind_cn": None, "period": None, "year": None,
           "company": None, "code": None, "evidence": None}
    try:
        pages = _page_texts(path)
    except Exception as e:  # unreadable file is a finding, not a crash
        rec["error"] = str(e)
        return rec

    scopes = []
    if pages:
        scopes.append((1, pages[0][:HEAD_CHARS]))
        scopes.append((1, pages[0]))
    if len(pages) > 1:
        scopes.append((2, pages[1][:HEAD_CHARS]))

    for page_no, text in scopes:
        cand = []
        for kind, pat in KIND_PATTERNS:
            for m in pat.finditer(text):
                cand.append((m.start(), kind, int(m.group(1)), m.group(0)))
        if cand:
            cand.sort(key=lambda x: x[0])
            _, kind, year, quote = cand[0]
            rec.update({"kind": kind, "kind_cn": KIND_CN[kind], "year": year,
                        "period": "%d-%s" % (year, KIND_END[kind]),
                        "evidence": {"page": page_no, "quote": quote}})
            break

    head = pages[0] if pages else ""
    m = COMPANY.search(head)
    if m:
        rec["company"] = m.group(1)
    m = CODE.search(head)
    if m:
        rec["code"] = m.group(1)
    return rec


def classify(reports):
    """Split the declared reports into usable / excluded / blocking."""
    recs, blocking = [], []
    for rep in reports:
        path = rep.get("pdf")
        declared = rep.get("period")
        if not path or not os.path.exists(path):
            rec = {"pdf": path, "kind": None, "error": "file not found"}
        else:
            rec = detect(path)
        rec["declared_period"] = declared

        # A document the detector cannot name is not silently treated as an
        # annual report: that is the exact failure this module exists to stop.
        if rec.get("error") or not rec.get("kind"):
            blocking.append({"pdf": path, "why": rec.get("error") or "封面认不出报告类型",
                             "declared_period": declared})
        elif declared and str(declared) != rec["period"]:
            blocking.append({"pdf": path, "why": "参数写的报告期与封面不符：参数 %s，封面 %s"
                             % (declared, rec["period"]), "declared_period": declared})
        recs.append(rec)

    usable = [r for r in recs if r.get("kind") == "annual" and not r.get("error")]
    excluded = [r for r in recs if r.get("kind") and r.get("kind") != "annual"]

    # One workbook holds one company, and nothing used to check this. The annual
    # reports are indexed by (statement, label, year), so two issuers' filings in
    # one run put both companies' figures under the same key and the lookup
    # silently takes whichever was read first. Naming the reports one by one made
    # that the caller's problem; discovering them from a folder makes it a
    # certainty, so the check belongs here.
    groups = {}
    for r in recs:
        if r.get("company"):
            groups.setdefault(r["company"], []).append(r)
    if len(groups) > 1:
        blocking.append({
            "pdf": None,
            "why": "这批文件里有不止一家公司：%s。一份工作簿只装一家公司，"
                   "混在一起会静默算错，请只保留本公司的 PDF。"
                   % "；".join("%s（%d 份）" % (k, len(v))
                               for k, v in sorted(groups.items())),
            "declared_period": None})
    return recs, usable, excluded, blocking


def render(recs, usable, excluded, blocking):
    L = ["# 认料与排期 —— 这批 PDF 是什么，能填哪几列", ""]
    L.append("> 这一步只看每份报告的封面，几秒钟。它决定后面往哪写数：")
    L.append("> **非年报一律不写进 `年度数据`**，因为把一季报或半年报的数当年报用，")
    L.append("> 模型会照常算出一串数字，却没有任何地方报错。")
    L.append("")
    L.append("## 这批文件是什么")
    L.append("")
    L.append("| 文件 | 公司 | 报告类型 | 报告期 | 依据 |")
    L.append("|---|---|---|---|---|")
    for r in recs:
        ev = r.get("evidence") or {}
        src = "第 %s 页「%s」" % (ev.get("page", "-"), ev.get("quote", "")) if ev else (r.get("error") or "-")
        L.append("| %s | %s | %s | %s | %s |"
                 % (os.path.basename(r.get("pdf") or "-"),
                    r.get("company") or "-", r.get("kind_cn") or "**认不出**",
                    r.get("period") or "-", src))
    L.append("")
    L.append("## 能不能用")
    L.append("")
    if blocking:
        L.append("**这份运行被拦住了**，以下问题必须先解决：")
        L.append("")
        for b in blocking:
            L.append("- %s：%s" % (os.path.basename(b["pdf"] or "-"), b["why"]))
        L.append("")
    L.append("| 用途 | 用了哪几份 | 说明 |")
    L.append("|---|---|---|")
    years = sorted({r["year"] for r in usable})
    L.append("| 年度数据（估值用） | %s | 每份年报提供期末 + 期初两年 |"
             % ("、".join("FY%d" % y for y in years) if years else "**无**"))
    if excluded:
        L.append("| 非年报（只进季度表） | %s | 季报/半年报不进 `年度数据`；"
                 "它们填 `报告期历史数据`，`报告期利润变化` 的折线取自那里 |"
                 % "、".join("%s %s" % (r["kind_cn"], r["period"]) for r in excluded))
    else:
        L.append("| 非年报 | 无 | 本次没有收到季报或半年报 |")
    L.append("")
    if usable:
        L.append("## 排期：年报能填哪几列")
        L.append("")
        L.append("| 年报 | 填哪两年 |")
        L.append("|---|---|")
        for r in sorted(usable, key=lambda x: -x["year"]):
            L.append("| FY%d | %d 与 %d |" % (r["year"], r["year"], r["year"] - 1))
        L.append("")
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--params", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--out-report", required=True)
    a = ap.parse_args()
    params = json.load(io.open(a.params, encoding="utf-8"))
    recs, usable, excluded, blocking = classify(params.get("reports", []))

    out = {"phase": "preflight", "blocked": bool(blocking), "blocking": blocking,
           "reports": recs,
           "usable_annual": [{"pdf": r["pdf"], "period": r["period"], "year": r["year"],
                              "company": r.get("company"), "code": r.get("code"),
                              "kind": r["kind"]} for r in usable],
           "excluded": [{"pdf": r["pdf"], "kind": r["kind"], "kind_cn": r["kind_cn"],
                         "period": r["period"], "year": r["year"],
                         "company": r.get("company")} for r in excluded]}
    json.dump(out, io.open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    io.open(a.out_report, "w", encoding="utf-8", newline="\n").write(
        render(recs, usable, excluded, blocking))

    print("认料：%d 份文件" % len(recs))
    for r in recs:
        print("  %-46s %-10s %-12s %s"
              % (os.path.basename(r.get("pdf") or "-")[:46],
                 r.get("kind_cn") or "认不出",
                 r.get("period") or "-", r.get("company") or "-"))
    if excluded:
        print("非年报 %d 份（不写入 年度数据）：%s"
              % (len(excluded), "、".join("%s %s" % (r["kind_cn"], r["period"]) for r in excluded)))
    for b in blocking:
        print("  [BLOCK] %s" % b["why"])
    print("preflight ->", a.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())