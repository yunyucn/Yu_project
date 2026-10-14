"""[1] Parser: annual-report PDF -> structured JSON.

Locates statements by title (not fixed page numbers) so the same rules work
across issuers, and switches statement ownership mid-page because the
consolidated and parent-company statements often share a page.
"""
import argparse
import hashlib
import json
import os
import re

import pdfplumber

CONFIG = None
CFG_SHA = None


def _load_cfg(path):
    global CONFIG, CFG_SHA
    raw = open(path, "rb").read()
    CFG_SHA = hashlib.sha256(raw).hexdigest()[:12]
    CONFIG = json.loads(raw.decode("utf-8"))
    return CONFIG


def _self_sha():
    """Hash of this file, so a parser change invalidates every cached parse.

    A cache keyed only on the PDF would survive a change to the extraction rules
    and quietly serve yesterday's reading of the same document -- the cache would
    then be the thing producing a stale number.
    """
    return hashlib.sha256(open(os.path.abspath(__file__), "rb").read()).hexdigest()[:12]


def cache_key(pdf_path, period):
    h = hashlib.sha256()
    h.update(open(pdf_path, "rb").read())
    h.update(b"|" + str(period).encode("utf-8"))
    h.update(b"|" + str(_self_sha()).encode("ascii"))
    h.update(b"|" + str(CFG_SHA).encode("ascii"))
    return h.hexdigest()[:24]


def cache_get(cache_dir, key):
    if not cache_dir:
        return None
    path = os.path.join(cache_dir, key + ".json")
    if not os.path.exists(path):
        return None
    try:
        return json.load(open(path, encoding="utf-8"))
    except Exception:
        return None


def cache_put(cache_dir, key, value):
    if not cache_dir:
        return
    os.makedirs(cache_dir, exist_ok=True)
    json.dump(value, open(os.path.join(cache_dir, key + ".json"), "w",
                          encoding="utf-8"), ensure_ascii=False)


def assemble_rows(lines):
    """Pair label-only and number-only groups into logical rows.

    A statement row is frequently split into two text groups, because the label
    and the numbers sit a fraction of a point apart. Either order occurs, so both
    are handled. Trailing label-only groups are deliberately NOT appended to the
    previous row: section headers ("流动资产：", "流动负债：") have exactly that
    shape, and appending them corrupts the preceding line item. A wrapped label is
    instead matched by adding its truncated form to that row's `match` list.
    """
    out = []
    i = 0
    n = len(lines)
    while i < n:
        cur = lines[i]
        nxt = lines[i + 1] if i + 1 < n else None
        if cur["label"] and not cur["nums"] and nxt and not nxt["label"] and nxt["nums"]:
            out.append({"label": cur["label"], "nums": nxt["nums"]})
            i += 2
            continue
        if not cur["label"] and cur["nums"] and nxt and nxt["label"] and not nxt["nums"]:
            out.append({"label": nxt["label"], "nums": cur["nums"]})
            i += 2
            continue
        out.append(cur)
        i += 1
    return out


def is_caption(raw):
    """True when a line is a table caption that happens to read like a title.

    The restatement note ("追溯调整") lists the affected statements as captions
    above their adjustment rows:

        项目 变更前 变更后 变更影响
        合并资产负债表：
        其他应付款 ...
        合并利润表：
        销售费用 ...

    norm_title strips the trailing colon, so those captions matched a statement
    title exactly and started a capture that ran on until the end of the file.
    Measured across three 兴源环境 annual reports: every real statement title is
    numbered and colon-free ("3、合并利润表"), and every caption is
    colon-terminated. The colon is the discriminator.
    """
    return str(raw).rstrip().endswith(("：", ":"))


def norm_title(s):
    """Fold a candidate line down to a bare statement title.

    Annual reports mention statement names in prose and in tables of contents
    ("合并资产负债表主要科目变动"), so a substring test attributes half the MD&A to
    the balance sheet. Only an exact title, optionally numbered and wrapped over
    two lines, starts a statement.
    """
    s = re.sub(r"[\s\u3000]", "", str(s))
    s = re.sub(r"^[0-9]+[\u3001.\uff0e]", "", s)
    s = re.sub(r"^[\uff08(][0-9\u4e00-\u9fa5]{1,3}[\uff09)]", "", s)
    s = re.sub(r"[\uff1a:]+$", "", s)
    s = re.sub(r"[\uff08(]续[\uff09)]$", "", s)
    return s


def statement_complete(rows, terminals):
    """True when the capture ran to the statement's own total row.

    Absence of a line item only means zero if the statement was captured in
    full. Without this check a truncated capture (or a run that started in the
    wrong place) would write zeros over real figures, which is the failure mode
    this whole pipeline exists to prevent.
    """
    # Containment rather than equality: labels carry section numbering and
    # qualifiers ("六、期末现金及现金等价物余额", "归属于母公司所有者的净利润"),
    # and any of those appearing means the capture reached the end of the table.
    have = "\n".join(re.sub(r"[\s\u3000]", "", str(r["label"])) for r in rows)
    return any(re.sub(r"[\s\u3000]", "", t) in have for t in terminals)


def _group_unclaimed(entries, max_gap, min_rows, anchors):
    """Runs of number-carrying lines that no statement claimed.

    A title the table does not know is not a parse error: the capture never
    starts, the rows fall through unowned, and the statement arrives as zero
    rows with nothing said anywhere. Group the orphans into blocks, keep the
    ones long enough to be a statement and carrying a statement end-marker, and
    report them. Reported only -- whether a title belongs in the table is a
    judgement for a human, and a fuzzy fallback here is what once mapped
    归母净利润 onto plain 净利润.
    """
    blocks, cur = [], None
    for e in entries:
        if (cur is not None and e["page"] == cur["page"]
                and e["line"] - cur["last_line"] <= max_gap):
            cur["last_line"] = e["line"]
            cur["last_label"] = e["label"]
            cur["labels"].append(e["label"])
            cur["rows"] += 1
            continue
        if cur is not None:
            blocks.append(cur)
        cur = {"page": e["page"], "first_label": e["label"],
               "last_label": e["label"], "last_line": e["line"],
               "labels": [e["label"]], "rows": 1}
    if cur is not None:
        blocks.append(cur)
    out = []
    for b in blocks:
        if b["rows"] < min_rows:
            continue
        hits = [a for a in anchors if any(a in lab for lab in b["labels"])]
        if not hits:
            continue
        b["anchors"] = sorted(set(hits))
        b.pop("labels", None)
        out.append(b)
    return out


def scan_statements(pdf, cfg):
    numre = re.compile(cfg["number_pattern"])
    skipre = re.compile("|".join(cfg["skip_line_patterns"]))
    tol = cfg["row_group_tolerance"]
    titles = {}
    for name, spec in cfg["statements"].items():
        for t in spec["titles"]:
            titles[t] = name
    stop = {norm_title(t): t for t in cfg["stop_titles"]}
    want = {norm_title(t): name for t, name in titles.items()}
    # The supplementary schedule lives inside the notes, so the note table that
    # follows it would otherwise be swallowed. Each statement can declare the
    # patterns that end its capture.
    ends = {name: ([re.compile(p) for p in spec.get("stop_patterns", [])],
                   spec.get("stop_after_rows"))
            for name, spec in cfg["statements"].items()}
    out = {name: [] for name in cfg["statements"]}
    # A line the issuer lists but leaves blank is a stated nil, which is different
    # from a line the issuer does not list at all. Keep both facts: dropping the
    # blank ones (as the row loop does) made every nil balance look like a missing
    # one. `single` is the same idea for a row that carries only the current
    # period: the prior-period column is then nil.
    nil = {name: [] for name in cfg["statements"]}
    single = {name: [] for name in cfg["statements"]}
    # A figure that does not match number_pattern stays glued to its label
    # ("（一）基本每股收益-0.1462-0.1945"), so the line looks like a header and the
    # row vanishes with no error. Count that shape so it is visible when it happens.
    glued = []
    gluedre = re.compile(r"[-0-9,]+[.][0-9]{2,}$")
    # Rows no statement owns, and captures that ran past their own total row.
    # Both mean a statement title was not recognised; without them the run
    # reports a shorter column and raises nothing.
    ucfg = cfg.get("unclaimed") or {}
    max_gap = int(ucfg.get("max_line_gap", 4))
    min_rows = int(ucfg.get("min_rows", 12))
    overrun_tol = int(ucfg.get("overrun_after_rows", 8))
    anchors = [re.sub(r"[\s\u3000]", "", t)
               for terms in (cfg.get("terminal_rows") or {}).values()
               for t in terms]
    unclaimed = []
    suppressed = False
    cur = None
    for page_no, page in enumerate(pdf.pages, 1):
        words = sorted(page.extract_words(), key=lambda w: (w["top"], w["x0"]))
        groups = []
        for w in words:
            if groups and abs(w["top"] - groups[-1][0]) <= tol:
                groups[-1][1].append(w)
            else:
                groups.append([w["top"], [w]])
        lines = []
        for _, g in groups:
            g.sort(key=lambda x: x["x0"])
            label = "".join(w["text"] for w in g if not numre.match(w["text"])).strip()
            nums = [float(w["text"].replace(",", "")) for w in g if numre.match(w["text"])]
            lines.append({"label": label, "nums": nums})
        lines = assemble_rows(lines)
        for j, r in enumerate(lines):
            if cur is not None:
                pats, cap = ends.get(cur, ([], None))
                if any(p.search(r["label"]) for p in pats):
                    cur = None
                    continue
                if cap is not None and len(out[cur]) >= cap:
                    # A cap is a safety net, not a normal exit: reaching it means
                    # the capture probably started in the wrong place. Say so
                    # rather than silently truncating.
                    print("  [WARN] %s hit stop_after_rows=%d; capture stopped"
                          % (cur, cap))
                    cur = None
                    continue
            # Titles wrap across lines too ("...经营活动现金流" / "量：").
            probe = r["label"]
            if r["label"] and not r["nums"] and j + 1 < len(lines):
                nxt = lines[j + 1]
                if nxt["label"] and not nxt["nums"]:
                    probe = r["label"] + nxt["label"]
            # The bare line is tried first: the line after a title is often an
            # unrelated "编制单位：" line, which the concatenation would swallow.
            key = norm_title(r["label"])
            key2 = norm_title(probe) if probe != r["label"] else None
            # A caption is neither a boundary nor a start: skipping it leaves an
            # in-progress capture running, which is what a mid-notes caption
            # ("合并资产负债表：" inside the restatement note) actually means.
            if not is_caption(r["label"]):
                if key in stop or key2 in stop:
                    # A parent-company statement is a deliberate boundary, not a
                    # miss: its rows are not orphans to be reported.
                    cur = None
                    suppressed = True
                    continue
                if key in want:
                    cur = want[key]
                    suppressed = False
                    continue
                if key2 is not None and key2 in want:
                    cur = want[key2]
                    suppressed = False
                    continue
            if (cur is not None and r["label"] and not skipre.search(r["label"])
                    and not any(p.search(r["label"]) for p in ends.get(cur, ([], None))[0])):
                if not r["nums"]:
                    nil[cur].append(r["label"])
                elif len(r["nums"]) == 1:
                    single[cur].append(r["label"])
            if cur is not None and not r["nums"] and r["label"] and gluedre.search(r["label"]):
                glued.append((cur, page_no, r["label"][-40:]))
            if (cur is None and not suppressed and r["nums"] and r["label"]
                    and not skipre.search(r["label"])):
                unclaimed.append({"page": page_no, "line": j,
                                  "label": r["label"], "nums": len(r["nums"])})
            if cur is None or not r["nums"] or not r["label"]:
                continue
            if skipre.search(r["label"]):
                continue
            out[cur].append({"page": page_no, "label": r["label"], "nums": r["nums"]})
    if glued:
        print("  [WARN] %d label(s) end in a decimal number but carry no value; "
              "a figure may not match number_pattern:" % len(glued))
        for stmt, pno, lab in glued[:6]:
            print("         p%-4s %-16s %r" % (pno, stmt, lab))
    # A capture that ran past its own total row means the title of the statement
    # that follows was not recognised: the two were read as one. Measured on
    # 兴源环境's 2025 Q3 report, whose income statement is captioned
    # 合并年初到报告期末利润表 -- the balance sheet swallowed it whole, and the
    # capture still stopped short of stop_after_rows, so not even the cap warned.
    overrun = {}
    for name, rows in out.items():
        terms = [re.sub(r"[\s\u3000]", "", t)
                 for t in (cfg.get("terminal_rows") or {}).get(name, [])]
        if not terms or not rows:
            continue
        last = None
        for i, row in enumerate(rows):
            lab = re.sub(r"[\s\u3000]", "", str(row["label"]))
            if any(t in lab for t in terms):
                last = i
        if last is None:
            continue
        extra = len(rows) - 1 - last
        if extra > overrun_tol:
            overrun[name] = {"at_row": last + 1, "extra": extra,
                             "page": rows[last]["page"],
                             "after": [r["label"] for r in rows[last + 1:last + 5]]}

    findings = {"overrun": overrun,
                "blocks": _group_unclaimed(unclaimed, max_gap, min_rows, anchors),
                "empty": sorted(n for n, r in out.items() if not r)}
    for name, o in sorted(overrun.items()):
        print("  [WARN] %s 在合计行之后仍多抓 %d 行（p%s 起：%s）——紧跟其后的报表标题"
              "可能没被认出来" % (name, o["extra"], o["page"],
                                  "、".join(str(x) for x in o["after"][:2])))
    for b in findings["blocks"]:
        print("  [WARN] p%s 有 %d 行数字没有任何报表认领：%s … %s"
              % (b["page"], b["rows"], b["first_label"][:18], b["last_label"][:18]))
    return out, nil, single, findings


NO_DIVIDEND = re.compile(
    r"不派发现金红利|不派发现金股利|不进行现金分红|不分配现金红利"
    r"|不进行利润分配|不分配利润|不分配、不转增|不分配不转增")
NO_BONUS = re.compile(r"不送红股|不送股|不进行送股")
NO_TRANSFER = re.compile(
    r"不以公积金转增股本|不进行资本公积金转增股本|不进行公积金转增股本"
    r"|不转增股本|不进行转增股本")

# A dividend phrase only counts when it sits in dividend wording: a bare
# "不派发" elsewhere in the report is not the distribution plan.
DIVIDEND_CONTEXT = re.compile(r"分红|股利|红利|利润分配|转增|派发|送红股")
YEAR_IN_TEXT = re.compile(r"(20\d{2})\s*年")

# The issuer prints its own arithmetic one row below the per-10-share figure:
#   每10股派息数（元）（含税） 3.50
#   分配预案的股本基数（股） 419,145,267
#   现金分红金额（元）（含税） 146,700,843.45
# Reproducing the third from the first two is the only independent check the PDF
# itself offers, so it is the one that gates the write.
SHARE_BASE = re.compile(r"分配预案的股本基数[^0-9\-]{0,12}(-?[\d,]+\.?\d*)")
CASH_AMOUNT = re.compile(r"现金分红金额[^0-9\-]{0,20}(-?[\d,]+\.?\d*)")

CROSS_CHECK_TOLERANCE = 0.02
EVIDENCE_SCAN_PAGES = 80


def _snippet(text, start, end, pad=45):
    """A short single-line quote around a match, for the delivery note."""
    lo, hi = max(0, start - pad), min(len(text), end + pad)
    return re.sub(r"\s+", " ", text[lo:hi]).strip()


def _speaks_of_period(text, start, end, period_year):
    """True when the wording around a phrase is about the period being reported.

    A report routinely restates the PRIOR year's plan inside the current year's
    narrative: the FY2025 annual report of 江波龙 prints "2024年度利润分配方案为：不派发现金红利" on the
    very page that states its FY2025 plan pays 3.50 per ten shares. Reading that
    as "this year pays nothing" turns a paying issuer into a zero-dividend one,
    and DDM then returns 0 without raising an error anywhere.

    Accept the wording when it names the current period, or names no year at all
    ("公司计划不派发现金红利") while sitting in a dividend context.
    """
    window = text[max(0, start - 80):min(len(text), end + 80)]
    if not DIVIDEND_CONTEXT.search(window):
        return False
    years = {int(y) for y in YEAR_IN_TEXT.findall(window)}
    if period_year is None:
        return True
    return period_year in years or not years


def _payout_cross_check(texts, pdf, extra_pages=3):
    """The issuer's own cash amount and share base, when it printed them."""
    base = amount = None
    last = max(texts) if texts else 0
    for page_no in range(1, min(last + extra_pages, len(pdf.pages)) + 1):
        text = texts.get(page_no)
        if text is None:
            text = pdf.pages[page_no - 1].extract_text() or ""
        if base is None:
            m = SHARE_BASE.search(text)
            if m:
                base = float(m.group(1).replace(",", ""))
        if amount is None:
            m = CASH_AMOUNT.search(text)
            if m:
                amount = float(m.group(1).replace(",", ""))
        if base is not None and amount is not None:
            break
    return base, amount


def extract_extras(pdf, cfg, period_year=None):
    """Dividend-per-share figures live in the profit-distribution section.

    A loss-making issuer has no number to report: it writes the absence in words
    ("不派发现金红利，不送红股，不以公积金转增股本"). Searching only for the
    numeric labels then yields nothing, the const rows stay unresolved, and the
    run reports a dictionary problem where the answer is simply zero. Scan the
    wording as well, and treat an explicit statement of absence as 0 -- but only
    when the wording is about the period being reported, and only when the
    issuer's own arithmetic agrees with the number.

    Every decision lands in `_evidence` (page, quote, basis) so the delivery note
    can show a human why the model says what it says.
    """
    out = {}
    want = {"每10股派息数": None, "每10股送红股数": None, "每10股转增数": None}
    absent = {"每10股派息数": False, "每10股送红股数": False, "每10股转增数": False}
    pairs = {"每10股派息数": NO_DIVIDEND, "每10股送红股数": NO_BONUS,
             "每10股转增数": NO_TRANSFER}
    evidence = {}
    texts = {}
    # The distribution plan sits in the front matter (business review / important
    # matters) or in the profit-distribution section, well inside the first
    # eighty pages of every report seen so far.
    for page_no, page in enumerate(pdf.pages[:EVIDENCE_SCAN_PAGES], 1):
        text = page.extract_text() or ""
        texts[page_no] = text
        for key in list(want):
            if want[key] is None:
                m = re.search(re.escape(key) + r"[^0-9\-]{0,12}(-?[\d,]+\.?\d*)", text)
                if m:
                    want[key] = float(m.group(1).replace(",", ""))
                    evidence[key] = {"page": page_no, "basis": "numeric",
                                     "quote": _snippet(text, m.start(), m.end())}
            if not absent[key]:
                for am in pairs[key].finditer(text):
                    if _speaks_of_period(text, am.start(), am.end(), period_year):
                        absent[key] = True
                        evidence[key] = {"page": page_no, "basis": "explicit_absence",
                                         "quote": _snippet(text, am.start(), am.end())}
                        break
        if all(v is not None for v in want.values()):
            break

    # A figure beats a sentence. The table row is the issuer's own number for the
    # period; a sentence about absence elsewhere in the report may be about
    # another year, and it must not veto a figure that does exist.
    for key, v in want.items():
        if v is not None and absent[key]:
            absent[key] = False

    dps = want["每10股派息数"]
    if dps is not None:
        base, amount = _payout_cross_check(texts, pdf)
        if base and amount:
            expect = dps / 10.0 * base
            if abs(expect - amount) / max(abs(amount), 1.0) > CROSS_CHECK_TOLERANCE:
                ev = evidence.get("每10股派息数", {})
                evidence["每10股派息数"] = {
                    "page": ev.get("page"), "basis": "cross_check_failed",
                    "quote": ev.get("quote", ""),
                    "expected": round(expect, 2), "reported": amount}
                # Neither figure can be trusted, and a wrong dividend is not a
                # small error: it is the whole DDM block. Refuse to write it, so
                # the row lands in the gap list instead of a confident 0.
                want["每10股派息数"] = None
            else:
                evidence["每10股派息数"]["cross_checked"] = {
                    "shares": base, "amount": amount}

    for k, v in want.items():
        if v is not None:
            out[k] = v
        elif absent[k]:
            out[k] = 0.0
    out["_evidence"] = evidence
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--params", required=True)
    ap.add_argument("--pages-config", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--cache-dir", help="reuse parses of unchanged PDFs")
    a = ap.parse_args()
    cfg = _load_cfg(a.pages_config)
    params = json.load(open(a.params, encoding="utf-8"))
    # 年报 and 季报/半年报 are read by the same rules and kept in separate lists.
    # They are not interchangeable: 年度数据 takes only the annual ones (a quarter
    # written into a year-end column is the silent failure this pipeline exists to
    # stop), while 报告期历史数据 is a quarter-granularity sheet and takes both.
    queue = ([(r, "annual") for r in params.get("reports", [])]
             + [(r, "interim") for r in params.get("reports_interim", [])])
    sources, sources_interim = [], []
    hits = 0
    for rep, bucket in queue:
        sink = sources if bucket == "annual" else sources_interim
        # Parsing a 190-page report takes about 25 s; a re-run that only added one
        # PDF should not pay for the others again. The key carries the PDF bytes,
        # the period, this file, and the page config, so any change to the reading
        # rules misses the cache instead of serving a stale parse.
        key = cache_key(rep["pdf"], rep.get("period"))
        cached = cache_get(a.cache_dir, key)
        if cached is not None:
            sink.append(dict(cached, pdf=rep['pdf'], period=rep['period'],
                             kind=rep.get('kind', 'annual')))
            hits += 1
            print("  %-52s [cached]" % rep["pdf"][-46:])
            continue
        with pdfplumber.open(rep["pdf"]) as pdf:
            stmts, nil, single, findings = scan_statements(pdf, cfg)
            extras = extract_extras(pdf, cfg, int(str(rep["period"])[:4]))
        terminal = cfg.get("terminal_rows", {})
        complete = {name: statement_complete(rows, terminal.get(name, []))
                    for name, rows in stmts.items()}
        src = {"pdf": rep["pdf"], "period": rep["period"],
               "kind": rep.get("kind", "annual"),
               "statements": stmts, "extras": extras,
               "complete": complete,
               "nil_labels": {k: sorted(set(v)) for k, v in nil.items()},
               "single_labels": {k: sorted(set(v)) for k, v in single.items()},
               "unclaimed": findings}
        cache_put(a.cache_dir, key, src)
        sink.append(src)
        counts = {k: len(v) for k, v in stmts.items()}
        print("  %-52s [%s] %s" % (rep["pdf"][-46:], bucket, counts))
        print("      complete=%s extras=%s" % (complete, extras))
    if hits:
        print("  %d/%d report(s) served from cache" % (hits, len(queue)))
    json.dump({"sources": sources, "sources_interim": sources_interim},
              open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    if sources_interim:
        print("  %d 份季报/半年报 -> 报告期历史数据（不写入 年度数据）" % len(sources_interim))
    print("wrote", a.out)


if __name__ == "__main__":
    main()

