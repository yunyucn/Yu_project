# Domain model

The vocabulary and rules of the domain this skill operates in. Read this before
changing the mapping dictionary or the compatibility thresholds — most of the
constraints in `README.md` follow from something on this page.

## Entities

| Concept | Meaning | Where it lands |
|---|---|---|
| Company | the valuation subject | `目录!C6` / `F6` |
| Reporting period | annual or half-year | `目录!C7` drives the column mapping |
| Statement | a structured set of accounts | 年度数据 segments |
| Statement scope | consolidated vs parent-only | only consolidated is a data source |
| Line item | an account name, a **natural-language label** | `年度数据!A` |
| Fact | company x item x **past** year = value | `年度数据!C:H` |
| Assumption | company x item x **future** year = value | `假设` (human judgement) |
| Valuation | f(facts, assumptions) | `绝对估值!D40`, `相对估值!D6/D25` |

## The fact/assumption line is the automation boundary

| | Fact | Assumption |
|---|---|---|
| Time | happened | not yet |
| Source | annual report | a person |
| Automatable | yes | **no, and it should not be** |
| Failure mode | detectable by reconciliation | unfalsifiable, only challengeable |
| Owner | the tool | the person who signs |

The DCF is dominated by assumptions, not by history: a terminal growth rate of 3%
versus 5% can double the output. Automate facts; hand assumptions over with a
default and a flag.

**Exception — A′ items.** `假设!I55` (goodwill), `I68` (contract liabilities),
`I79` (deferred tax liability), `I81` (provisions) sit on the *assumptions* sheet
but are the sample company's **actual balance-sheet figures**, hard-coded with no
link to `年度数据`. They must be written from the new company's data or the old
company's balance sheet leaks into the new valuation — silently.

## Three boundaries

1. **Fact vs assumption** — above.
2. **Consolidated vs parent-only** — the two often share a page. Switch ownership on
   the title; never split by page.
3. **Units, three orders of magnitude apart** — PDF is 元, `年度数据` is 百万元
   (x 1e-6), per-share items stay in 元/股, and 每10股派息数 needs x 0.1. Declare
   the scale per row; a per-share value scaled by 1e-6 becomes 0.0000034 and raises
   nothing.

## Invariants

Held on real data (2024/2025, see `references/invariants.md`):

- 资产总计 = 负债合计 + 股东权益合计
- 资产总计 = 负债和股东权益合计
- 间接法经营净额 = 直接法经营净额
- 期末现金 = 净增加 + 期初

These are **data-correctness** evidence and are independent of how many years are
available. Keep them separate from depth checks.

Sign convention: the same economic event carries different signs per statement.
`固定资产折旧、油气资产折耗、生产性生物资产折旧` is positive in the cash-flow
supplement (added back) and a cost in the income statement.

## Traps

**Same name, different meaning.** `净利润` appears in both the income statement and
as the indirect-method starting point in the cash-flow supplement. Key the mapping on
`(statement, normalised label)` — keyed on the label alone, one overwrites the other.

**Attribution.** `净利润` includes minorities; `归属于母公司股东的净利润` does not;
the valuation uses the latter. A fuzzy matcher once mapped one onto the other. Do not
reintroduce similarity matching.

**Old vs new accounting standards.** 可供出售金融资产 / 持有至到期投资 were replaced
by 交易性金融资产 / 债权投资 / 其他债权投资 / 其他权益工具投资; 应收利息 and
应付利息 folded into 其他应收款 / 其他应付款. These rows are `mode: zero` because the
company applies the new standard — not because the PDF search failed.

**Variant characters.** 其它/其他, 盈余公积金/盈余公积, 资本公积金/资本公积. Fold
these in normalisation; keep the alias list for genuinely different wording.

**Generic statement format.** A non-financial company's income statement still carries
financial-sector rows (已赚保费, 退保金, 赔付支出净额, 发放贷款和垫款 ...) at zero.
Decide these by company type, not by whether the PDF shows the row.

## Domain features that force the architecture

| Feature | Consequence |
|---|---|
| Account names are an **open set** | a mapping dictionary plus an adapter layer; never exhaustive |
| Arithmetic is **strict** (one missing year flips the DCF) | cannot be relaxed away; only completed or explicitly degraded |
| Failures are **silent** (errors become 0) | assert actively; a human eye cannot tell |
| Reports are **restated** | the same year can differ across reports; record which report and index each value came from |
| Half-year and annual periods **differ** | never mix them into one annual column |
| Market data is **not in the report** | PE/PS/PB/EV-EBITDA/PCF need a separate supply path |

## Time semantics

`C:H` runs oldest to newest, driven by `目录!C7`. Each report supplies two years
(current and prior), so three reports taken every other year cover six columns.

For a target year `Y`, the mapper prefers a report whose period is `Y` (index 0) over
one whose period is `Y+1` (index 1). Accounting-policy changes and error corrections
can restate a prior year, so the two may disagree; the choice above prefers the
company's own annual statement over a later report's comparative figure. Record the
source either way so a disagreement stays traceable.

