"""[3a] Validator (data side): assertions on the resolved field map.

Run before writing. The template swallows missing history into zeros instead of
raising errors, so these checks must be explicit.
"""
import argparse
import json
import sys

COLORDER = ["C", "D", "E", "F", "G", "H"]


def check_depth(fm, spec):
    """Every row must have all of D:H; a single missing year zeroes the derived
    assumptions and flips the DCF sign."""
    need = set(spec.get("cols", ["D", "E", "F", "G", "H"]))
    by_row = {}
    for key, c in fm["cells"].items():
        by_row.setdefault(c["row"], {})[c["col"]] = c
    market = {r["row"] for r in []}
    bad = []
    for row, cols in by_row.items():
        if "market" in {c.get("mode") for c in cols.values()}:
            continue
        missing = [c for c in need if c not in cols]
        if missing:
            bad.append({"row": row, "missing_cols": sorted(missing)})
    return {"id": spec["id"], "passed": not bad, "detail": bad[:20], "count": len(bad)}


def check_reconcile(fm, spec):
    tol = spec.get("tol", 0.005)
    bad = []
    for key, c in fm["cells"].items():
        if c.get("mode") != "pdf" or "raw_yuan" not in c:
            continue
        # The scale is per row: 年度数据 holds millions of yuan, except the
        # per-share rows, which hold yuan per share. Reconciling against a
        # hardcoded 1e-6 flagged every EPS cell as a mismatch.
        expect = round(c["raw_yuan"] * c.get("scale", 1e-6), 6)
        if abs(c["value"] - expect) > tol:
            bad.append({"cell": key, "written": c["value"], "from_pdf": expect})
    return {"id": spec["id"], "passed": not bad, "detail": bad[:20], "count": len(bad)}


def check_identity(fm, spec):
    tol = spec.get("tol", 0.005)
    by_row_col = {(c["row"], c["col"]): c["value"] for c in fm["cells"].values() if "value" in c}
    bad = []
    for col in COLORDER:
        left = by_row_col.get((spec["left"], col))
        rights = [by_row_col.get((r, col)) for r in spec["right"]]
        if left is None or any(v is None for v in rights):
            continue
        if abs(left - sum(rights)) > tol:
            bad.append({"col": col, "left": left, "right_sum": sum(rights)})
    return {"id": spec["id"], "passed": not bad, "detail": bad, "count": len(bad)}


HANDLERS = {"no_gaps": check_depth, "reconcile": check_reconcile, "identity": check_identity}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--fieldmap", required=True)
    ap.add_argument("--assertions", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    fm = json.load(open(a.fieldmap, encoding="utf-8"))
    cfg = json.load(open(a.assertions, encoding="utf-8"))
    results = []
    for spec in cfg["pre_write"]:
        fn = HANDLERS.get(spec["type"])
        if fn is None:
            results.append({"id": spec["id"], "passed": None, "skipped": spec["type"]})
            continue
        results.append(fn(fm, spec))
    # Coverage gaps and mapping failures are both "we do not have this value",
    # so both belong to L2: they are reported and they move the verdict, but
    # they do not stop the run. What stops the run is data that is present and
    # wrong (identities) or a value that does not match its source (reconcile).
    # Blocking here would also discard the fact layer, which stays useful.
    gaps = fm.get("coverage_gaps", [])
    unresolved = fm.get("unresolved", [])
    for tag, items in (("coverage_gaps", gaps), ("mapping_failures", unresolved)):
        if items:
            print("  [INFO] %s %d (L2 owns this; see compat.md)" % (tag, len(items)))
    if unresolved:
        for u in unresolved[:10]:
            print("        r%-4s %-30s %s" % (u["row"], str(u["label"])[:30],
                                              u.get("reason", "")))
    failed = [r for r in results if r["passed"] is False]
    out = {"phase": "pre_write", "passed": not failed, "results": results}
    json.dump(out, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    for r in results:
        mark = "PASS" if r["passed"] else ("SKIP" if r["passed"] is None else "FAIL")
        extra = (" (%d)" % r["count"]) if r.get("count") else ""
        print("  [%s] %s%s" % (mark, r["id"], extra))
        if r["passed"] is False:
            for d in r["detail"][:6]:
                print("        ", d)
    print("pre-write validation:", "PASSED" if out["passed"] else "FAILED")
    sys.exit(0 if out["passed"] else 1)


if __name__ == "__main__":
    main()


