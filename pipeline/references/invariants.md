# Invariants and measured failure modes

Every number here comes from a controlled experiment on a copy of the template.
Raw data lives in the project's `_analysis/*.json`.

## 1. History depth: D:H is mandatory

Baseline = the shipped template with all six years present. Variants delete the
earliest years. Same workbook, same formulas, one variable.

| Metric | 6y | 5y | 4y | 3y | 2y |
|---|---|---|---|---|---|
| `假设!I19` tax rate | 0.1529 | 0.1529 | **0** | **0** | **0** |
| `假设!I6` expense ratio | 0.0896 | 0.0896 | **0** | **0** | **0** |
| `绝对估值!D40` DCF | **+1527.56** | +1421.81 | **-1009.68** | -1082.42 | -1220.57 |
| `相对估值!D5` mean PE | 28.19 | 28.19 | 23.27 | 15.84 | **8.62** |
| Error values | 0 | 0 | 0 | 0 | 0 |

Three mechanisms:

1. **Derived assumptions collapse to zero.** They read
   `假设!I5 = IF(ISNUMBER(AVERAGE(H5,G5,F5,E5,D5)), AVERAGE(...), 0)`. A missing year
   yields `0/0 = #DIV/0!`, `AVERAGE` propagates it, `ISNUMBER` is false, and the whole
   assumption returns `0`. Tax rate, tax/surcharge ratio, expense ratio, minority
   interest and dividend payout ratio all collapse. `绝对估值!B8 = 假设!I19`, so the
   effective tax rate becomes zero. **The `IF(ISNUMBER(...), ..., 0)` idiom converts an
   error into a plausible-looking zero.**
2. **Mean multiples get diluted.** `年度数据!I127 = IF(ISNUMBER(AVERAGE(D144:H144)),
   AVERAGE(D144:H144), AVERAGE(F144:H144))`. Missing years count as zero: −17% for one,
   −44% for two, −69% for three.
3. **The DCF changes sign.** +1527.56 becomes −1009.68.

**No error is raised in any variant.** The `输出` sheet shows only recent years, so it
looks identical whether the history is complete or not.

Column usage across the `假设` sheet: `C` 2 references, `D` 14, `E` 14, `F` 14, `G` 14,
`H` 16. All `AVERAGE`-based assumptions read `D:H` only, so column `C` is optional —
the 5-year variant keeps every derived assumption intact and the DCF within −7%.

## 2. Residual data leaks silently

The template ships with 科顺股份. Variants: `base` untouched; `root_only` clears
`年度数据!C:H` and the identity cells and sets the period to 2025; `clean_all` also
clears the residual assumption cells.

| Cell | base | root_only | clean_all |
|---|---|---|---|
| `假设!I3` revenue growth | 0.0921 | **0.0921** | blank |
| `假设!I4` cost ratio | 0.8231 | **0.8231** | blank |
| `假设!I55` goodwill | 72.88 | **72.88** | blank |
| `假设!I68` contract liabilities | 426.13 | **426.13** | blank |
| `假设!I79` deferred tax liability | 16.43 | **16.43** | blank |
| `假设!I81` provisions | 114.37 | **114.37** | blank |
| `绝对估值!D40` DCF | 1527.5614 | **23.0301** | 0 |

Replacing only the history sheet is not enough. The previous company's growth rate,
cost ratio, goodwill and other balance-sheet items keep driving the forecast, and the
result is a non-zero number with no meaning attached. Nothing is reported as an error.

## 3. Write performance: batch the array

1176 cells into `年度数据!H`:

| Method | Write | Total |
|---|---|---|
| per-cell `Cells(r,c).Value=` | **11.03 s** | 11.13 s |
| single array assignment | 0.12 s | 0.26 s |
| manual calc + array assignment | **0.01 s** | 0.25 s |

Baseline: open 1.79 s, recalc 0.07 s, save 0.83 s. Roughly 3 s per company.

Array assignment needs a contiguous range, so write the whole `lo..hi` block per column
and let unmapped rows (title rows, already empty) receive `None`.

## 4. The read-only attribute defeats Save()

The shipped template has the DOS read-only attribute (`mode 0444`). `shutil.copy2`
propagates it, Excel opens the copy read-only, `Save()` returns without error, and the
file hash is unchanged. `wb.ReadOnly` is `True` in this state. Call
`os.chmod(path, stat.S_IWRITE)` on the copy and assert `wb.ReadOnly is False`.

## 5. Match labels exactly

PDF labels can be truncated mid-phrase (`投资收益（损失以“－”号填`), so normalisation
truncates at an unmatched opening parenthesis and strips a trailing footnote marker.
Matching is exact after normalisation plus an explicit alias list. A similarity
fallback previously mapped `归属于母公司所有者净利润` onto plain `净利润` at 0.96 similarity —
a wrong answer that looks right. Unmatched labels must fail loudly.

## 6. COM from Python only

PowerShell's `New-Object -ComObject Excel.Application` fails with
`TYPE_E_CANTLOADLIBRARY`. Use Python `win32com.client.Dispatch`.

## 7. Known pre-existing defects

Not caused by this skill, worth flagging in any report: 140 `#NUM!` on `相对估值`,
17 `#REF!` on `估值`, 7 `_xll.BD_COLLECTION` = `#NAME?` (the vendor XLL is gone), 5 stale
external links, and leftover sample strings. These do not affect filling, but they
pollute the output.

## 8. Errors arrive as negative integers

Excel hands error values back through COM as `CVErr` integers, not strings:
`#DIV/0!` is `-2146826281`, `#REF!` `-2146826265`, `#VALUE!` `-2146826273`. A
`startswith("#")` scan therefore reported a clean workbook while 211 error values sat
in it across 8 sheets before any work, and 699 after a bare reset. `com_excel.is_error()`
covers both forms, and `no_new_errors` compares each sheet against the template as
opened rather than against a magic constant.

## 9. The market rows are blanked and were never refilled

`年度数据` rows 144-148 hold six years of historical PE / PS / PB / EV-EBITDA / PCF.
`相对估值!D5` reads `年度数据!I127`, which averages row 144. The reset blanks rows
3-224, and these five rows are `mode: market`, so nothing rewrote them: `相对估值`
returned `#DIV/0!` and no assertion noticed. They must be supplied through
`params.market.history`; a missing year aborts unless `allow_missing_history` is set.

## 10. Statement titles must match exactly, and a statement needs a boundary

Substring title matching let a prose mention of 合并资产负债表 in the MD&A (pages 44,
53, 63) capture 160 rows of narrative into the balance sheet. In the same run the
supplementary schedule swallowed the note tables after it: 292 rows instead of 29.
`norm_title()` now demands an exact title (leading numbering and a title wrapped over
two lines are tolerated), and a statement declares `stop_patterns` / `stop_after_rows`
for its own end.

## 11. Collapse is per-assumption, and it is measured rather than inferred

Clearing one `年度数据` column at a time and reading all 65 derived assumptions:

| Missing column | C | D | E | F | G | H |
|---|---|---|---|---|---|---|
| Assumptions that zero out | **0** | 9 | 21 | 22 | 23 | 25 |

`C` is genuinely optional; `D:H` is not. `假设!I19` (effective tax rate) collapses on
any of `D:H`, `绝对估值!B8` reads it, and the DCF flips sign.

Inferring this from formulas gets it wrong in both directions, because the wrapper is
`IF(ISNUMBER(AVERAGE(...)), AVERAGE(...), 0)`:

- an **unguarded division** (`=利润表!D23/利润表!D22`) turns a missing input into
  `#DIV/0!`, `ISNUMBER` is false, and the assumption becomes `0`;
- a **plain reference** (`=年度数据!D33`) turns a missing input into `0`, which is then
  averaged in and merely dilutes the result — no collapse;
- a **guarded division** (`=IF(资产负债!C49+资产负债!D49<>0, ..., 0)`) returns `0` by
  design — no collapse.

`config/assumption_deps.json` therefore carries the measured `collapse_on_missing` list
per assumption, not a derived column span.

That experiment answers the coverage-gap question only. See section 14 for the
mapping-failure case, which it hides.

## 12. Coverage gaps and mapping failures were one list

`map_fields.py` once emitted a single `unresolved` list. On two 江波龙 reports it held
532 entries, of which **498 were years with no report at all** and **34 were genuine
label mismatches**. Read together this looked like 45% coverage when the covered years
were 100% mapped, and it left the operator unable to tell whether to fetch a PDF or
edit the dictionary. They are now `coverage_gaps` and `unresolved`.

## 13. A hard depth gate threw away the fact layer

The depth check originally lived in the pre-write assertions and aborted the run. That
is the wrong place: it discards work that is still usable, and it reports a blunt
"needs five years" where the useful answer is "fetch FY2021". Depth and mapping
failures now belong to the L2 compatibility layer, which reports and moves a verdict;
only data that is present and wrong (identities) or a value that disagrees with its
source (reconcile) stops the run.

## 14. A missing cell is a different failure from a missing column

Clearing a whole column cannot tell you what one missing cell does: it removes the
fatal cell and the harmless ones together, and it removes every cell of any range an
assumption averages over. A mapping failure is exactly one cell, so it was measured
that way — clear one cell, recalculate, read all 65 assumptions, restore — across the
214 cells the assumptions actually reference:

| Outcome | Cells |
|---|---|
| Fatal on its own (an assumption collapses to 0) | **24** |
| No effect at all | **55** |
| Dilutes an average, no collapse | 135 |

The fatal cells are the **denominators**, and there are only six rows' worth:

| 年度数据 row | Item | Assumptions it destroys |
|---|---|---|
| **4** | 营业收入 | 14 — every margin and expense ratio, including 4 that drive the DCF |
| 25 | 利润总额 | 2 — the effective tax rate, so the DCF sign flips |
| 28 | 净利润 | 4 |
| 30 | 归属于母公司所有者净利润 | 2 |
| 67 | 投资性房地产 | 2 |
| 118 | 实收资本(股本) | 2 |

So **营业收入 is the one row that must never be missing in any of `D:H`.** One unmatched
revenue line is enough to zero 14 assumptions.

Cross-check: the cells that are fatal singly are a subset of the cells whose column is
fatal, for every column, with one exception — `假设!J105` averages a range, so it
tolerates a blank cell but errors when the whole column is blank. It is recorded in
`range_collapse_only`. Both lists are needed; neither derives the other.

The verdict consequence is not cosmetic. Before this was measured, a single missing
revenue cell produced verdict `full`: `collapsed` stayed empty, `dcf_reliable` stayed
true, and the run wrote a workbook whose DCF was built on a zeroed effective tax rate.
The partial-gap path now consults `collapse_cells` and reaches `facts_only`.

Evidence: `_analysis/cell_collapse.json` (214 probes), `_analysis/collapse_experiment.json`
(6 column probes).

## 15. Calibrated against three real annual reports (兴源环境 300266.SZ)

Three consecutive annual reports (FY2023, FY2024, FY2025; 273/295/265 pages) cover
FY2022-FY2025, i.e. four of the six history columns. Result after the fixes below:

| Outcome | Cells |
|---|---|
| Value taken from the PDF | 552 |
| Listed by the issuer but blank -> 0 | 108 |
| Structural zero (row is always zero) | 132 |
| Dividend constant (issuer pays none) | 18 |
| Summed inside 年度数据 | 12 |
| Coverage gap (2020, 2021 -- no report) | 330 |
| Market multiples (not in any PDF) | 30 |
| **Unmatched labels** | **0** |

For every year a report covers, every PDF-sourced row is accounted for. All five
pre-write assertions pass, including the balance-sheet identity and the
indirect-vs-direct cash-flow reconciliation. The run still refuses to value,
because 2021 is missing and `D` is mandatory.

## 16. A decimal limit in `number_pattern` made whole rows disappear

`number_pattern` accepted 1-2 decimals, because some issuers write EPS as `1.2`.
兴源环境 FY2025 writes it as `-0.1462`. The word then failed the number test, stayed
glued to the label, and the assembled row read
`'（一）基本每股收益-0.1462-0.1945'` with **no numbers at all** — so the row was
dropped as a label-only line and the value vanished without any error.

Widened to 1-4 decimals. The general shape is now counted and printed: any line
under capture whose label ends in a decimal number but carries no value raises a
`[WARN]`, because that is what a pattern miss looks like from the inside.

## 17. A colon made a table caption look like a statement title

The restatement note captions its adjustment table with the statement names:

```
项目 变更前 变更后 变更影响
合并资产负债表：
其他应付款 ...
合并利润表：
销售费用 ...
```

`norm_title` strips a trailing colon, so `合并利润表：` matched the title exactly and
started a capture that never ended: **763 rows** of 合并利润表 against the 36 the
statement actually has, running from page 10 to page 245 and silently polluting
合并资产负债表 as well (63 rows against 57).

The discriminator is the colon. Across the three reports: every real title is
numbered and colon-free (`3、合并利润表`); all 6 false positives were
colon-terminated. A caption is now neither a start nor a stop, and every statement
carries `stop_after_rows` as a cap that prints a warning when it fires.

## 18. "Absent" is only zero when something says so

Reading a missing line as a nil balance is correct for A-share statements — they
follow a standard format and issuers drop nil lines. But applied bluntly it turns
an extraction miss into a confident `0`. Two safeguards, both evidence-based:

- **Listed but blank** is positive evidence of nil: the issuer printed the line and
  left the value column empty. The parser now records those labels instead of
  discarding them, and they map to `0`. Same for a row carrying only the current
  period — the prior-period column is nil.
- **Reported elsewhere but missing here** is a suspected parse miss. If a row
  resolves in one covered year and not another, it is reported rather than zeroed.
  This is what caught the EPS case in section 16.

With both in place the 兴源环境 run has 108 evidence-based zeros and 0 unexplained
misses. Without them it had 109 zeros of which one (EPS) was wrong, and 28 items
were flagged as suspicious that were in fact ordinary nil balances.

## 19. Per-share rows are not millions of yuan

`年度数据` holds millions of yuan, so PDF values are scaled by 1e-6. Row 31
(每股EPS) holds **yuan per share**: the template's own sample row reads
0.48/0.31/0.6/0.82/0.596/0.15. Scaling it turned 兴源环境's -0.64 into -1e-06, and
the reconcile assertion — which also hardcoded 1e-6 — flagged every EPS cell.

The conversion is now declared per row (`scale` in `mapping.json`, default 1e-6) and
read by both the mapper and the reconcile check. Rows 33-35 (每股派现/红股/转增)
were already correct because they carry an explicit `scale: 0.1`.



## 20. The oldest column of the window is a gap that breaks nothing

Clearing one column at a time and reading all 65 derived assumptions gives
missing `C` -> **0** collapsed, `D` -> 9, `E` -> 21, `F` -> 22, `G` -> 23,
`H` -> 25. `C` is inside the six-year window but outside every `AVERAGE` the
assumptions use, so `compat.py` reports it as a coverage gap while nothing
depends on it. The gap list therefore marks it **non-blocking**: a run that is
missing only `C` still produces a usable DCF. Reporting both missing years with
the same weight would send the analyst after a report that changes no conclusion.

Because a report for FY(N) carries both the closing balance of N and the opening
balance of N-1, the minimum fetch is smaller than the list of missing years. A
gap of 2020+2021 needs **one** FY2021 report. `gaps.py` computes that set
(`minimal_reports`), taking the newest missing year first.

## 21. 137 cells look like a finished forecast and are a formula

`假设!I:L` holds 146 literal blanks. Six are real inputs (`I3`, `J3`, `I4`, `I7`,
`I8`, `I105`), 124 are chapter headings and captions (`II﹑资产负债表假设`,
`流动资产：`, `（1）IPO募集资金:`), and 137 are neither blank nor input: they carry
`IF(ISNUMBER(AVERAGE(历史)), ..., 0)`, which the template fills with the
company's own historical average as soon as a company is loaded. `K3`/`L3` are
one step further removed -- they average `I3`/`J3`, so a blank `J3` drags them down.

Two failure modes follow. Reporting the 137 as "to fill" tells the analyst to redo
work the template already did; reporting nothing lets them read an average as a
judgement. The card prints them in a separate table headed 模板公式（非人工输入）
and `gaps.md` lists all of them. `config/analyst_inputs.json` carries the split so
no script has to re-derive it.

The six real inputs are the mirror image of the same trap: blank is not neutral.
`假设!I4` blank makes `利润表!I3 = I2*0`, a 100% gross margin, silently.

## 22. `年度数据!O:T` is a hidden mirror of the sample company

197 rows x 6 columns (O3:T224), hidden (column width 0), outside the print area,
carrying 科顺's values cell-for-cell. A reference scan over all 27,912 formulas,
all charts, and all 231 defined names finds **zero** references. Writing the same
values, clearing and restoring them, and editing an unrelated sheet all produce
an identical 591-cell recalc drift, so the block is decorative.

It is cleared anyway (P1 in `cleanup.json`): leaving it means a new company's
workbook still contains the previous issuer's numbers, and it costs 5 pre-existing
`#DIV/0!` cells. Measured effect of clearing it on the 兴源环境 run: workbook
errors 1809 -> 1804, all 5 in `年度数据`; every downstream value unchanged
(`绝对估值!D40` 3067.598129944243, `输出!B58/B60/B62` identical), zip fidelity
identical.


## 23. 分红口径是唯一一个「0」两种含义重合的地方

`绝对估值` 的 DDM 块在公司不分红时是 0，而一个漏取的洞也长成 0：
`绝对估值!D53 = 利润表!I33 = 假设!I22 × 假设!I136`，而
`假设!I22 = IF(ISNUMBER(AVERAGE(H22,G22,F22,E22)), AVERAGE(...), 0)`。两者在表里完全一样，
工作簿不报任何错。由此得出三条规则，均用真实年报测过：

- **「不分红」的话只有在说本期时才算数。** 江波龙（301308） 2025 年报在同一页上，
  一边写「公司2024年度利润分配方案为：不派发现金红利」（讲去年），
  一边写 2025 年度每10股派 3.50 元。没有期间门控（`_speaks_of_period`）时，
  前一句足以把一家分红公司算成零分红。门控接受“点名本期”或“完全不提年份且在分红语境内”
  的句子，拒绝“点名其他年份”的句子。
- **数字压过句子，且必须能复现公司自己的算术。** 利润分配表会同时印
  每10股派息数、分配预案的股本基数（股）、现金分红金额（元）（含税）；
  `派息数 / 10 × 股本基数` 必须等于印出的金额（2% 容差）。江波龙：
  3.50 / 10 × 419,145,267 = 146,700,843.45。对不上就说明两个数里有一个读错了，
  该行宁可拒写（进待补清单），也不能写下去。
- **行 36-38（中报分红）是空白，不是 0。** 模板出厂时把它们写成字面量 `0`，
  等于对每家公司断言“没有中期分红”。中期分红只在半年报里披露，
  只吃年报的运行拿不到，所以 `mapping.json` 把这几行标为 `optional` + `value: null`：
  单元格保持空白，并作为待办列出。这里空白与 0 在算术上等价
  （`假设!I138 = H136×H34 + (…)×年度数据!H37`，是乘法），下游不变。

每个判定都连同页码与原文记入 `_evidence`，经 `fieldmap.json` 的
`dividend_evidence` 流到 `请先看这个.md` 与 `gaps.md`：读者不用打开 PDF 就能分清
“事实”和“漏洞”。只有一家分红样本（江波龙），
“排版不同就抽不到”的概率仍未估计；再补银行/白酒/家电各一家再说。


## 24. A report's period comes from its cover, not from the caller

`period` was typed into params.json and never checked. The pipeline took
`int(period[:4])`, so `2025-03-31` and `2025-12-31` were the same input, and
nothing opened the document. A first-quarter report declared as `2025-12-31` was
therefore read as FY2025: its three months went into column `H` of 年度数据,
every derived assumption averaged them, and no stage reported anything. The
workbook that came out looked complete.

`detect.py` now reads the cover first and the period follows from the kind:

| Cover | Period end | Written to 年度数据 |
|---|---|---|
| 年度报告 | YYYY-12-31 | yes, YYYY and YYYY-1 |
| 半年度报告 | YYYY-06-30 | no |
| 第一季度报告 | YYYY-03-31 | no |
| 第三季度报告 | YYYY-09-30 | no |

Three measured details, all of which cost something to learn:

- **半年度报告 contains 年度报告 as a substring.** Matching 年度报告 first
  labels every half-year report as an annual one -- the exact mistake the module
  exists to prevent. The half-year pattern is tried first, and the earliest match
  position in the title block wins.
- **Only the first 500 characters of the cover count.** A report mentions other
  filings further down the page; those mentions are not its identity.
- **A disagreement stops the run.** If `period` and the cover differ, neither is
  preferred: the caller's belief and the document are in conflict, and picking one
  silently is how a wrong period gets in.

Non-annual reports are named on the card and in `gaps.md` rather than dropped.
Dropping a file the caller handed in without saying so is its own quiet error.

Verified on all 13 PDFs in the workspace: 5 annual, 4 half-year, 2 first-quarter,
plus duplicates -- every one classified correctly, including 江波龙's
2026 half-year and 2025 first-quarter reports.

## 25. The template build is part of the contract

`config/mapping.json` keys every rule by **row number**. A different build of the
workbook moves rows, and nothing in the write path would notice: every value lands
somewhere plausible, every cell shows a number, and no error is raised. So
`mapping.json` records the template's `sha256` and `reset_template` refuses to run
against a different one. `params.allow_template_mismatch = true` overrides it, and
should only ever be used after diffing the two builds.

The same reasoning makes `run_for_human.py` locate the template by **sha256** rather
than by filename: the folder usually also holds "已解锁" / "副本" rewrites, and
picking one by name builds the whole deliverable on a modified engine.

## 26. Not every cell in the `C:H` block is ours to write

Row 147 (`历史EV/EBITDA`) is a formula the template owns (`=财务分析!D65`). Both the
reset and the loader write `年度数据!C:H` as **one array**, and an array write puts
`None` in every row it covers — so a formula inside the block is erased, not
overwritten with a wrong number. The result reads as *blank*, which is why it looked
like nothing for several builds.

`com_excel.write_block_keep_formulas` carries existing formulas through the write,
and the `formulas_intact` post-write assertion fails the run if one is lost. The row
is also declared `mode: formula` in the mapping, so `compat.py` neither counts it as
filled nor reports it as a gap.

