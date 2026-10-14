"""Resolve human-written analyst inputs to template cells.

The six judgement cells are the one part of the model a person supplies. Asking
that person to find `假设!I3` is asking them to speak the template's language
instead of their own, so `params.analyst` accepts the label they already use:

    "analyst": {"营业收入增长率(2026E)": 0.03}

Lives apart from write_model on purpose: the gap list needs the same resolution to
know which of the six are already answered, and gaps.py must run without Excel.
"""
import re

FULLWIDTH = (("\uff08", "("), ("\uff09", ")"), ("\uff1a", ":"), ("\uff0c", ","),
             ("\uff1b", ";"), ("\uff05", "%"))


def norm(s):
    """Fold the ways a human might type a cell name into one key."""
    s = str(s).strip()
    for a_, b_ in FULLWIDTH:
        s = s.replace(a_, b_)
    return s.replace(" ", "").replace("\u3000", "").lower()


def resolve(cells_cfg, wanted):
    """{name: value} -> ({cell: value}, [names that matched nothing]).

    A name that matches two different cells (e.g. the label without its year) is
    treated as unmatched rather than guessed: writing 2027E's growth rate into
    2026E would be a silent, plausible, wrong number.
    """
    by_key, ambiguous = {}, set()
    for it in cells_cfg.get("true_inputs", []):
        cell = it["cell"]
        keys = [cell, cell.split("!")[-1],
                "%s(%s)" % (it["label"], it.get("year", "")),
                "%s%s" % (it["label"], it.get("year", "")),
                it["label"]]
        for k in keys:
            nk = norm(k)
            if nk in by_key and by_key[nk] != cell:
                ambiguous.add(nk)
            else:
                by_key[nk] = cell
    for k in ambiguous:
        by_key.pop(k, None)

    out, unknown = {}, []
    for key, value in (wanted or {}).items():
        if value is None:
            # A key left as null in the params file means "not supplied". The
            # example ships every key that way, so copying it verbatim must be a
            # no-op rather than a refusal.
            continue
        cell = by_key.get(norm(key))
        if cell is None:
            unknown.append(key)
        else:
            out[cell] = value
    return out, unknown
