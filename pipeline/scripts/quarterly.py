"""[5] Quarterly channel: 年报 + 季报 -> 报告期历史数据.

年度数据 holds six year-ends. The template keeps a second, quarter-granularity
history on 报告期历史数据, and that sheet is the only thing 报告期利润变化 reads:
twenty income-statement rows across twelve dated columns, 240 formulas in all.
Nothing wrote to it, so it arrived blank and the chart drew no line -- the
complaint was "报告期利润变化没有数据和折线图".

Three facts make the mapping safe, and each was measured on the template:

  * A column is chosen by the date its own header row names
    (=CONCATENATE(数据输入!O5,"-12-","31") ...), never by position. The header is
    driven by 目录!C8 = TODAY()-1, so the window slides with the calendar; a
    positional write would put a quarter in the wrong column on New Year's Day.
  * The values are year-to-date cumulative, which is the basis a quarterly
    income statement is published on. A quarter's figures therefore go in as
    printed -- no subtraction.
  * The second numeric column of a balance sheet is the previous year-END, while
    the second column of an income or cash-flow statement is the same period one
    year earlier. Keying the index by the date a number actually refers to (not
    by year) is what keeps 2024-03-31, read off a 2025 Q1 report's 上年同期, from
    being written into 2024-12-31, read off its 年初余额.

A value is only ever written for a period some report actually speaks to. A
column with no report behind it stays empty and is listed in 待补清单; it is
never filled with zero.
"""
import argparse
import io
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from map_fields import norm

BALANCE = "合并资产负债表"
DATE_RE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})")


def parse_date(s):
    m = DATE_RE.match(str(s))
    return (int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else None


def iso(d):
    return "%04d-%02d-%02d" % d


def last_year_end(d):
    return (d[0] - 1, 12, 31)


def same_period_last_year(d):
    return (d[0] - 1, d[1], d[2])


def comparative(stmt, d):
    """The date the second numeric column of this statement refers to."""
    return last_year_end(d) if stmt == BALANCE else same_period_last_year(d)


def build_index(sources):
    """(statement, normalised label, ISO date) -> the numbers that claim that date.

    Each entry remembers whether it came from the report whose own period is that
    date (index 0) or from the following year's comparative (index 1). Index 0
    always wins: it is the year the issuer was reporting on, not a restated
    look-back.
    """
    idx = {}
    for src in sources:
        end = parse_date(src.get("period"))
        if end is None:
            continue
        for stmt, rows in (src.get("statements") or {}).items():
            for row in rows:
                lab = norm(row.get("label"))
                for i, v in enumerate((row.get("nums") or [])[:2]):
                    when = end if i == 0 else comparative(stmt, end)
                    idx.setdefault((stmt, lab, iso(when)), []).append(
                        {"index": i, "value": v, "pdf": src.get("pdf"),
                         "kind": src.get("kind") or "annual"})
    return idx


def build_complete(sources):
    """(statement, ISO date, index) -> the report captured that statement whole.

    Only a whole capture licenses reading "the line is absent" as "the balance is
    nil". A balance sheet's comparative column is the previous year-end, so
    completeness carries to that date as well.
    """
    out = {}
    for src in sources:
        end = parse_date(src.get("period"))
        if end is None:
            continue
        for stmt, ok in (src.get("complete") or {}).items():
            if ok:
                out[(stmt, iso(end), 0)] = True
                out[(stmt, iso(comparative(stmt, end)), 1)] = True
    return out


def build_stated_nil(sources):
    """(statement, ISO date, index) -> the issuer listed the line and left it blank.

    Positive evidence of a nil balance. It is what separates "the line is there
    and reads zero" from "the line is not in this statement", and only the first
    of those may be written as 0.
    """
    out = {}
    for src in sources:
        end = parse_date(src.get("period"))
        if end is None:
            continue
        for stmt, labels in (src.get("nil_labels") or {}).items():
            for lab in labels:
                out.setdefault((stmt, iso(end), 0), set()).add(norm(lab))
                out.setdefault((stmt, iso(comparative(stmt, end)), 1), set()).add(norm(lab))
        for stmt, labels in (src.get("single_labels") or {}).items():
            for lab in labels:
                out.setdefault((stmt, iso(comparative(stmt, end)), 1), set()).add(norm(lab))
    return out


def lookup(idx, stmt, labels, when):
    for want in (0, 1):
        for lab in labels:
            for key in ((stmt, norm(lab), when),
                        (stmt, norm(lab).lstrip("一二三四五六七八九十"), when)):
                for hit in idx.get(key, []):
                    if hit["index"] == want:
                        return hit
    return None


def resolve(rows, idx, complete, stated_nil, targets):
    """One pass per target date; nothing is written for a date nobody speaks to."""
    # A row this issuer reports in some other period is a suspected extraction or
    # dictionary miss where it does not resolve -- not a nil balance. Without this
    # the 2025 Q1 report, which words the line 归属于母公司所有者的净利润 while the
    # dictionary knew only 归属于母公司股东的净利润, would have had a loss of
    # 25.87 百万元 written as a confident 0. The annual path has always refused
    # that reading; the same refusal belongs here.
    found_any = set()
    for row in rows:
        if row.get("mode") != "pdf":
            continue
        for when in targets:
            if lookup(idx, row["statement"], row.get("match") or [], when) is not None:
                found_any.add(row["row"])
                break

    # Which (statement, date, index) some report actually speaks to. Without it
    # "the line is missing" and "the statement is missing" read the same, and
    # every balance-sheet row of a column that only has an income statement --
    # 2024-03-31 comes from a 2025 Q1 comparative -- would be reported as a
    # suspected extraction miss. It is not: there is simply no balance sheet for
    # that date in this batch, which the column list already says.
    spoken = set()
    for (stmt, _lab, when), hits in idx.items():
        for h_ in hits:
            spoken.add((stmt, when, h_["index"]))

    def is_stated_nil(stmt, labels, when, index):
        pool = stated_nil.get((stmt, when, index))
        return bool(pool) and any(norm(lab) in pool for lab in labels)

    by_period, unresolved, absent = {}, [], []
    for when in targets:
        cells = []
        for row in rows:
            if row.get("mode") != "pdf":
                continue
            labels = row.get("match") or []
            hit = lookup(idx, row["statement"], labels, when)
            if hit is None:
                own = (row["statement"], when, 0) in spoken
                # A nil line is only readable from the period the report was
                # actually about. Off a comparative column the statement may be
                # abbreviated, and a zero there would overwrite a figure the next
                # report states outright.
                if own and complete.get((row["statement"], when, 0)):
                    stated = is_stated_nil(row["statement"], labels, when, 0)
                    if stated or row["row"] not in found_any:
                        cells.append({"row": row["row"], "value": 0.0,
                                      "mode": "absent_zero",
                                      "basis": "listed_blank" if stated else "never_listed",
                                      "statement": row["statement"]})
                        absent.append({"row": row["row"], "label": row["label"],
                                       "period": when})
                        continue
                if own:
                    reason = ("该科目在别的报告期能取到，本期取不到——怀疑是解析或映射漏掉，"
                              "不是真的为 0，故留空"
                              if row["row"] in found_any else "本期报表没抓全，不敢判 0")
                elif (row["statement"], when, 1) in spoken:
                    reason = "本期没有这张报表，只有上年同期可比列；不从可比列判 0"
                else:
                    reason = "本期没有这张报表"
                unresolved.append({"row": row["row"], "label": row["label"],
                                   "period": when, "statement": row["statement"],
                                   "reason": reason})
                continue
            scale = row.get("scale", 1e-6)
            cells.append({"row": row["row"], "value": round(hit["value"] * scale, 6),
                          "mode": "pdf", "raw_yuan": hit["value"], "scale": scale,
                          "matched": row["match"][0] if row.get("match") else None,
                          "pdf": hit["pdf"],
                          "via": "本期报告" if hit["index"] == 0 else "上年同期"})
        if not cells:
            continue
        # calc rows are sums of their components in the same column; skipped
        # unless every component resolved, so a half-sum is never written
        seen = {c["row"]: c["value"] for c in cells}
        for row in rows:
            if row.get("mode") != "calc":
                continue
            parts = [seen.get(r) for r in row.get("rows", [])]
            if not parts or any(p is None for p in parts):
                continue
            cells.append({"row": row["row"], "value": round(sum(parts), 6),
                          "mode": "calc", "rows": row["rows"]})
        by_period[when] = cells
    return by_period, unresolved, absent


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--extract", required=True)
    ap.add_argument("--mapping", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    mapping = json.load(io.open(a.mapping, encoding="utf-8"))
    ex = json.load(io.open(a.extract, encoding="utf-8"))
    annual = ex.get("sources") or []
    interim = ex.get("sources_interim") or []
    sources = annual + interim

    idx = build_index(sources)
    complete = build_complete(sources)
    stated_nil = build_stated_nil(sources)
    targets = sorted({k[2] for k in idx})
    by_period, unresolved, absent = resolve(mapping["rows"], idx, complete,
                                            stated_nil, targets)

    # What each report actually yielded, and what the parser could not place.
    # A quarterly column can be written from a partial capture -- the balance
    # sheet parsed, the income statement did not -- and the column then exists
    # and is short with nothing saying so. 报告期利润变化 draws its line from the
    # income-statement rows, so the newest point is simply absent.
    #
    # Only the statements that report kind is meant to carry are judged. A
    # quarterly report publishes the three main statements and no 现金流量补充资料;
    # that is the format, not a shortfall, and counting it would train the reader
    # to ignore the warning that matters.
    expectation = mapping.get("statement_expectation") or {}
    coverage, incomplete, parser_warnings = {}, [], []
    for s in sources:
        per = s.get("period")
        if not per:
            continue
        kind = s.get("kind") or "annual"
        want = expectation.get(kind)
        comp = s.get("complete") or {}
        cov = {}
        for name, rows in (s.get("statements") or {}).items():
            if want is not None and name not in want:
                continue
            cov[name] = {"rows": len(rows), "complete": bool(comp.get(name))}
        coverage[per] = cov
        missing = sorted(k for k, v in cov.items()
                         if not v["rows"] or not v["complete"])
        if missing:
            incomplete.append({"period": per, "kind": kind, "missing": missing})
        u = s.get("unclaimed") or {}
        if u.get("overrun") or u.get("blocks"):
            parser_warnings.append({"pdf": s.get("pdf"), "period": per,
                                    "overrun": u.get("overrun") or {},
                                    "blocks": u.get("blocks") or []})

    out = {"version": "1.0.0", "sheet": mapping["sheet"],
           "by_period": by_period, "unresolved": unresolved,
           "absent_zeroed": absent,
           "statement_coverage": coverage,
           "incomplete": incomplete,
           "parser_warnings": parser_warnings,
           "source_reports": [{"pdf": s.get("pdf"), "period": s.get("period"),
                               "kind": s.get("kind") or "annual"} for s in sources]}
    json.dump(out, io.open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    print("季度通道：年报 %d 份 + 季报/半年报 %d 份" % (len(annual), len(interim)))
    for when in sorted(by_period):
        cs = by_period[when]
        direct = sum(1 for c in cs if c.get("via") == "本期报告")
        via = sum(1 for c in cs if c.get("via") == "上年同期")
        print("  %s  %3d 格（本期报告 %d，上年同期 %d）" % (when, len(cs), direct, via))
    print("报告期历史数据：可填 %d 列；填不了的格 %d 个（其中怀疑漏抽 %d 个）"
          % (len(by_period), len(unresolved),
             sum(1 for u in unresolved if "怀疑" in str(u.get("reason")))))
    for u in unresolved[:12]:
        if "怀疑" in str(u.get("reason")):
            print("   ! h%-4d %-28s %s" % (u["row"], str(u["label"])[:28], u["period"]))
    for item in incomplete:
        print("  注意：%s（%s）缺 %s —— 这一列只填了一部分"
              % (item["period"], item["kind"], "、".join(item["missing"])))
    if parser_warnings:
        print("  注意：%d 份 PDF 里有报表标题没被认出来，整张表可能没抓全"
              % len(parser_warnings))
    print("wrote", a.out)


if __name__ == "__main__":
    main()
