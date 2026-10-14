"""[2] Mapper: extracted PDF values -> 年度数据 cell values.

Each report supplies two years (current period and prior period), so a row is
resolved by finding the report whose period year equals the target year (index 0)
or is one year later (index 1).

Matching is exact after normalisation plus an explicit alias list. Fuzzy
matching is deliberately absent: a similarity fallback previously mapped
"归属于母公司所有者净利润" onto plain "净利润", which is a silent wrong answer.
"""
import argparse
import json
import re

NO_PAREN = re.compile(r"[（(][^（()）]*[）)]")


def norm(s):
    s = re.sub(r"[\s\u3000]", "", str(s))
    s = NO_PAREN.sub("", s)
    s = re.split(r"[（(]", s)[0]
    s = re.sub(r"[\u201c\u201d\"\'\u2019]", "", s)
    s = re.sub(r"^[\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]+[\u3001.]?", "", s)
    s = re.sub(r"^(\u51cf|\u52a0|\u5176\u4e2d)[:\uff1a]?", "", s)
    s = re.sub(r"[:\uff1a,\uff0c\u3002\u3001_\-\uff0d\u2014]", "", s)
    s = re.sub(r"\d+(\.\d+)?$", "", s)
    s = re.sub(r"^\d+(\.\d+)?", "", s)
    return s.strip()


def index_sources(sources):
    """(statement, normalised label) -> list of (year, index, value, source_pdf)."""
    idx = {}
    for src in sources:
        year = int(str(src["period"])[:4])
        for stmt, rows in src["statements"].items():
            for r in rows:
                key = (stmt, norm(r["label"]))
                for i, v in enumerate(r["nums"][:2]):
                    idx.setdefault(key, []).append((year - i, i, v, src["pdf"]))
    return idx


def lookup(idx, stmt, match_labels, year):
    for lab in match_labels:
        for key in ((stmt, norm(lab)), (stmt, norm(lab).lstrip("一、二、三、四、五、六、七、八、九、十"))):
            for y, i, v, pdf in idx.get(key, []):
                if y == year:
                    return v, {"matched": lab, "pdf": pdf, "year_index": i}
    return None, None



def const_reason(ev):
    """Why a constant row could not be resolved, in words a reader can act on."""
    if (ev or {}).get("basis") == "cross_check_failed":
        return (u"分红数据与年报自己的现金分红金额对不上"
                u"（每10股派息数/10 × 股本基数 = %.2f，年报写 %.2f）"
                u"，已拒绝写入"
                % (ev.get("expected", 0.0), ev.get("reported", 0.0)))
    return u"const value not supplied"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--extract", required=True)
    ap.add_argument("--mapping", required=True)
    ap.add_argument("--params", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    mapping = json.load(open(a.mapping, encoding="utf-8"))
    params = json.load(open(a.params, encoding="utf-8"))
    sources = json.load(open(a.extract, encoding="utf-8"))["sources"]

    # The run_pipeline preflight already refuses to pass a non-annual report down,
    # so reaching here with one means map_fields was invoked directly. It is the
    # last place that can stop a quarter's figures being averaged into a year, so
    # it stops rather than trusting the caller.
    non_annual = [s for s in sources if (s.get("kind") or "annual") != "annual"]
    if non_annual:
        print("REFUSED: %d report(s) are not annual reports:" % len(non_annual))
        for s_ in non_annual:
            print("  %s  kind=%s period=%s" % (s_.get("pdf"), s_.get("kind"), s_.get("period")))
        print("  年度数据 is annual only. A quarter or a half-year written into it")
        print("  produces a workbook that looks complete and is not.")
        return 1

    period_year = int(str(params["company"]["report_period"])[:4])
    cols = mapping["columns"]
    years = {}
    for c, back in (("C", 5), ("D", 4), ("E", 3), ("F", 2), ("G", 1), ("H", 0)):
        years[c] = period_year - back

    idx = index_sources(sources)

    # A line item that a fully captured statement does not mention has a balance
    # of zero: A-share statements follow a standard format and issuers drop the
    # lines that are nil. Only a complete capture licenses that reading, so the
    # parser records completeness per statement and it is carried here.
    complete = {}
    for src in sources:
        y = int(str(src["period"])[:4])
        for stmt, ok in (src.get("complete") or {}).items():
            if ok:
                complete[(stmt, y)] = True
                complete[(stmt, y - 1)] = True

    extras = {}
    for src in sources:
        if int(str(src["period"])[:4]) == period_year:
            extras = src.get("extras", {})
    market = params.get("market", {})

    # Which calendar years any supplied report can speak to. A report whose
    # period is Y supplies Y (index 0) and Y-1 (index 1). A year no report covers
    # is a COVERAGE gap, not a mapping failure: the two need different fixes
    # (fetch another PDF vs. extend the dictionary) and must never be merged.
    covered_years = set()
    for rep in params.get("reports", []):
        y = int(str(rep["period"])[:4])
        covered_years.update({y, y - 1})

    # Which rows this issuer reports anywhere among the covered years. A row that
    # is reported in some years but absent in another is a suspected extraction
    # miss rather than a nil balance -- a line that is genuinely zero is still
    # listed, with its value column left blank.
    # Positive evidence of a nil balance: the issuer listed the line but left the
    # column empty. A row carrying only the current period says the same about the
    # prior period. Both are zero, and both are distinguishable from a line that
    # is simply not in the statement.
    stated_nil = {}
    for src in sources:
        y = int(str(src["period"])[:4])
        for stmt, labels in (src.get("nil_labels") or {}).items():
            for lab in labels:
                stated_nil.setdefault((stmt, y), set()).add(norm(lab))
                stated_nil.setdefault((stmt, y - 1), set()).add(norm(lab))
        for stmt, labels in (src.get("single_labels") or {}).items():
            for lab in labels:
                stated_nil.setdefault((stmt, y - 1), set()).add(norm(lab))

    def is_stated_nil(stmt, match_labels, year):
        pool = stated_nil.get((stmt, year))
        if not pool:
            return False
        return any(norm(lab) in pool for lab in match_labels)

    found_any = set()
    for row in mapping["rows"]:
        if row["mode"] != "pdf":
            continue
        for year in sorted(set(years.values())):
            if year in covered_years and lookup(idx, row["statement"], row["match"], year)[0] is not None:
                found_any.add(row["row"])
                break

    cells, coverage_gaps, unresolved, warnings, absent_zeroed = {}, [], [], [], []
    optional_blank = []
    for row in mapping["rows"]:
        mode = row["mode"]
        # A formula row is the workbook's own arithmetic, not an input: there is
        # no value to write and, more importantly, no gap to report. Listing it as
        # a missing market multiple sends the analyst after data that does not
        # exist, because 历史EV/EBITDA is derived, not observed.
        if mode == "formula":
            continue
        for c, year in years.items():
            key = "%s!%s%d" % ("年度数据", c, row["row"])
            if mode == "zero":
                cells[key] = {"row": row["row"], "col": c, "value": 0.0, "mode": "zero"}
                continue
            if mode == "market":
                # Market multiples are in no annual report. Without them the
                # 相对估值 sheet returns #DIV/0! while the model still looks fine,
                # so a missing value is an error unless the caller opts out.
                got = None
                for period, vals in (market.get("history") or {}).items():
                    if int(str(period)[:4]) == year:
                        got = (vals or {}).get(row.get("field"))
                        break
                if got is None:
                    item = {"row": row["row"], "label": row["label"], "col": c,
                            "year": year, "field": row.get("field"),
                            "reason": "market data not supplied (B 类：不在年报 PDF 中)"}
                    (warnings if market.get("allow_missing_history") else unresolved).append(item)
                    continue
                cells[key] = {"row": row["row"], "col": c, "value": float(got),
                              "mode": "market", "field": row.get("field")}
                continue
            if mode == "calc":
                cells[key] = {"row": row["row"], "col": c, "mode": "calc",
                              "op": row["op"], "rows": row["rows"]}
                continue
            if mode == "const":
                v = None
                ev = None
                if row.get("source") == "利润分配" and row.get("field"):
                    v = extras.get(row["field"])
                    ev = (extras.get("_evidence") or {}).get(row["field"])
                    if v is not None and row.get("scale"):
                        v = v * row["scale"]
                if v is None:
                    v = row.get("value")
                if v is None:
                    item = {"row": row["row"], "label": row["label"], "col": c,
                            "field": row.get("field"), "reason": const_reason(ev)}
                    # A row whose only source is a report this run does not have
                    # (an interim dividend lives in a half-year report, never in
                    # an annual one) is left blank on purpose: writing 0 would
                    # state "this company pays nothing" as a fact. That is a
                    # to-do for the analyst, not a defect in the facts, so it
                    # goes to the warnings and leaves the verdict alone.
                    if row.get("optional"):
                        warnings.append(item)
                        optional_blank.append(item)
                    else:
                        unresolved.append(item)
                    continue
                entry = {"row": row["row"], "col": c, "value": v, "mode": "const"}
                if ev:
                    entry["evidence"] = ev
                cells[key] = entry
                continue
            v, meta = lookup(idx, row["statement"], row["match"], year)
            if v is None and complete.get((row["statement"], year)):
                stated = is_stated_nil(row["statement"], row["match"], year)
                if stated or row["row"] not in found_any:
                    cells[key] = {"row": row["row"], "col": c, "value": 0.0,
                                  "mode": "absent_zero", "statement": row["statement"],
                                  "basis": "listed_blank" if stated else "never_listed"}
                    absent_zeroed.append({
                        "row": row["row"], "label": row["label"], "col": c,
                        "year": year, "statement": row["statement"],
                        "basis": "listed_blank" if stated else "never_listed"})
                    continue
            if v is None:
                item = {"row": row["row"], "label": row["label"], "col": c,
                        "year": year, "statement": row["statement"],
                        "match": row["match"]}
                if row["row"] in found_any and complete.get((row["statement"], year)):
                    item["reason"] = ("statement captured, line not listed in %d while "
                                      "other years report it -- suspect a parse miss, "
                                      "not a nil balance" % year)
                    unresolved.append(item)
                    continue
                if year not in covered_years:
                    item["reason"] = "no report covers %d" % year
                    coverage_gaps.append(item)
                else:
                    item["reason"] = "no PDF line matched"
                    unresolved.append(item)
                continue
            # 年度数据 holds millions of yuan, but per-share rows (每股EPS) hold
            # yuan per share. Scaling those by 1e-6 turned 兴源环境's -0.64 into
            # -1e-06; the template's own sample row reads 0.48/0.31/0.6, so the
            # conversion is declared per row rather than assumed.
            scale = row.get("scale", 1e-6)
            cells[key] = {"row": row["row"], "col": c, "value": round(v * scale, 6),
                          "mode": "pdf", "raw_yuan": v, "scale": scale, **meta}

    out = {"version": "1.0.0", "years": years, "period_year": period_year,
           "cells": cells, "coverage_gaps": coverage_gaps,
           "unresolved": unresolved, "warnings": warnings,
           "absent_zeroed": absent_zeroed,
           "covered_years": sorted(covered_years),
           "market": market, "a_prime": mapping.get("a_prime", {}),
           "dividend_evidence": extras.get("_evidence", {}),
           "optional_blank": optional_blank}
    json.dump(out, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    total = sum(6 for r in mapping["rows"] if r.get("mode") != "formula")
    print("years:", years, " covered:", sorted(covered_years))
    print("resolved cells: %d / %d" % (len(cells), total))
    gap_years = sorted({g["year"] for g in coverage_gaps})
    print("COVERAGE GAPS %d  (years %s -- need more reports, not dictionary work)"
          % (len(coverage_gaps), gap_years))
    print("MAPPING FAILURES %d  (PDF on hand, label unmatched -- extend mapping.json)"
          % len(unresolved))
    print("ABSENT->0 %d  (statement captured in full, line not reported = nil balance)"
          % len(absent_zeroed))
    for u in unresolved[:25]:
        print("   r%-4s %-30s %s  %s" % (u["row"], str(u["label"])[:30],
                                         u.get("year", ""), u["reason"]))
    print("wrote", a.out)


if __name__ == "__main__":
    main()


