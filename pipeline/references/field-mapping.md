# Field mapping

`config/mapping.json` holds the rules that turn extracted PDF lines into `年度数据`
cell values. It contains **no company data** — values are supplied at run time.

## Structure

```json
{
  "version": "1.0.0",
  "template": { "file": "...", "sha256": "..." },
  "units": { "target": "百万元", "pdf": "元", "scale": 1e-6 },
  "columns": { "C": 0, "D": 1, "E": 2, "F": 3, "G": 4, "H": 5 },
  "a_prime": { "cells": { "I55": "商誉" }, "source_rows": { "商誉": 78 } },
  "rows": [ ... 197 entries ... ]
}
```

`columns` maps the history column to how many years back from the report period it is:
`C` is the oldest, `H` the newest.

## Row modes

| mode | Meaning |
|---|---|
| `pdf` | Take from a statement. Fields: `statement`, `match` (candidate labels, in order), `index` (0 = current period, 1 = prior period) |
| `zero` | The item does not appear in the PDF (old-standard accounts, or genuinely absent). Write `0` |
| `calc` | Sum of other rows in the same column. Fields: `op`, `rows` |
| `const` | Constant, or taken from `extras`. Fields: `value`, or `source`/`field`/`scale` |
| `market` | Market data, never in the PDF. Supplied through the params file |
| `formula` | The template computes the row itself. Nothing is written, and no gap is reported |

### `formula` rows

`年度数据!D147:H147` (历史EV/EBITDA) holds `=财务分析!D65`, derived from `目录`,
`资产负债` and `利润表`. It is the only formula row inside the data block, and it is
**not** market data: `相对估值!D82` reads `年度数据!I130`, which averages `D147:H147`.
Treating it as `market` erased it twice over — `reset_template` and `write_model` both
array-write the whole `C:H` block, and an array write puts `None` in every row it
covers — after which the EV/EBITDA history read empty for good and no market feed
could fill it. `com_excel.write_block_keep_formulas` now carries existing formulas
through the write, and `compat` excludes `formula` rows from its coverage denominator
so they are neither "filled" nor "missing".

## How a year is resolved

Each report supplies two years. For target year `Y`, the mapper looks for a report
whose period year equals `Y` (uses index 0) or `Y + 1` (uses index 1). So FY2021,
FY2023 and FY2025 reports cover 2020-2025 with no overlap needed.

## Normalisation

Applied to both template labels and PDF labels before comparison:

- strip whitespace
- drop balanced parenthetical content, then truncate at a leftover opening parenthesis
  (PDF labels get cut mid-phrase)
- drop leading 一、二、… enumerators
- drop leading 减 / 加 / 其中 followed by a colon
- drop punctuation
- drop leading and trailing footnote digits

## Adding a match

When `map_fields.py` reports `no PDF line matched` for a row, read the actual label from
the PDF and append it to that row's `match` list:

```json
{ "row": 17, "label": "加：投资净收益", "mode": "pdf",
  "statement": "合并利润表", "match": ["投资收益", "投资净收益"], "index": 0 }
```

Prefer adding the exact PDF wording over loosening normalisation. Never add a fuzzy or
similarity fallback: it once mapped 归母净利润 onto plain 净利润.

## Regenerating

`config/mapping.json` was generated from the template's `年度数据` row index plus the
calibrated rule set. If the template version changes, check `sha256` first — row numbers
and defined names may have moved, in which case the dictionary needs recalibration
before anything else will work.
