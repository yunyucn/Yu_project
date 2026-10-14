# Template map

## At a glance

34 worksheets, 231 defined names, 166 ActiveX controls, 31 charts, 5 external links,
VBA project, `calcPr iterate="1" iterateCount="10"` (circular references resolved by
iteration). The VBA/ActiveX layer is the vendor's data-import mechanism
(`MSForms`, `GetService`, `aspx`, `_xll.BD_COLLECTION`) talking to the 贝格/BDC
terminal — this skill replaces that mechanism. The calculation chain is pure
formulas; no buttons need pressing.

## The real input surface

Isolation experiment: multiply each input sheet by 10, recalculate, watch `绝对估值!D40`.

| Changed | `绝对估值!D40` | Verdict |
|---|---|---|
| `收入成本预测` product forecast columns filled with 9999 | 1527.5614 → 1527.5614 | isolated |
| `数据输入` data region ×10 (541 cells) | 1527.5614 → 1527.5614 | isolated |
| `报告期历史数据` ×10 (652 cells) | 1527.5614 → 1527.5614 | isolated |
| `年度数据!C:H` rewritten | 1527.56 → 2168.92 | **drives the valuation** |

Dependency closure of `绝对估值`: `假设, 利润表, 目录, 营运资本, 资产负债`.
The isolated sheets are not in it.

## Cells that matter

| Location | Meaning |
|---|---|
| `年度数据!C:H` | Six years of history, rows 3-224 (197 data rows), all literal values |
| `年度数据!C1:H1` | Header formulas — **never overwrite** |
| `年度数据!D147:H147` | The only formulas inside the data block: `=财务分析!D65`, the derived 历史EV/EBITDA. Not a market ratio, not writable — see `field-mapping.md` |
| `目录!C6` / `F6` | Ticker / company name |
| `目录!C7` | Latest annual report period — drives which six years are in play |
| `目录!F8` / `F9` / `F12` | Total shares / float shares / close price (millions of shares, CNY) |
| `目录!F15` / `F16` | HKD and USD exchange rates |
| `假设!I3, J3, AN3, AO3` | Revenue growth rate (analyst input) |
| `假设!I4` | Cost ratio (analyst input) |
| `假设!I7` / `I8` | Admin / R&D expense ratios (analyst input) |
| `假设!I55, I68, I79, I81` | Goodwill, contract liabilities, deferred tax liability, provisions — **hard-coded facts**, see below |
| `假设!I88, I92, I93, I104` | Construction period, salvage rate, interest capitalisation, repayment term |
| `假设!I109, I110, I111` | Long-term / short-term loan and deposit rates |
| `绝对估值!B2:B9` | Transition years, transition growth, terminal growth g, beta, Rf, Rm, bond rate |

## A′ cells: facts disguised as assumptions

`假设!I55` (goodwill), `I68` (contract liabilities), `I79` (deferred tax liability),
`I81` (provisions) hold the sample company's **actual balance-sheet figures** but are
hard-coded rather than linked to `年度数据`. They do not follow the history sheet.

The same items are available from the PDFs at `年度数据` rows 78, 94, 112, 111.
The loader copies the latest year's value into these cells. Skipping this leaves the
previous company's goodwill inside the new company's valuation.

## Units

| Sheet | Unit |
|---|---|
| `年度数据` | 百万元 (millions) |
| `数据输入`, `数据输入季度数据` | 元 (yuan) |
| `财务对照表` | 元 |
| PDF statements | 元 |

The mapper converts yuan to millions with a scale of `1e-6`.

## The colour legend is not a write whitelist

`说明!B16` states orange = must modify, yellow = optional, red = result. Measured:

- Orange `收入成本预测` (960 cells, "must modify") — **no effect on the valuation**.
- Yellow `假设!I3/I4/I7/I8` ("optional") — **these decide the valuation**.

Derive the input surface from experiments, not from the legend.

## `财务对照表`

A ready-made Chinese line-item ↔ field-code dictionary (193 entries, unit 元).
Useful as a cross-check when extending the mapping dictionary.

## Isolated sheets

`主营业务, 估值, 单产品敏感性分析, 图形, 备用数据, 季度数据, 季度财务分析, 年度数据重构,
情景表, 报告期利润变化, 数据, 现金流量原始数据, 说明, 财务对照表, 资产规模指数化分析, 输出_BD`

`情景表` holds year-by-year forecast snapshots under 较好/一般情景 headings and is the
workbook's only traceability mechanism. The `目录` sheet lists sheets that do not exist
in this build (`假设记录表`, `同比报表`, `杜邦分析`).

## Verification anchors

| Cell | Meaning |
|---|---|
| `绝对估值!D40` | DCF value |
| `相对估值!D6` / `D25` | Per-share value, PE / PS method |
| `输出!B58`, `B60`, `B62` | Revenue, net profit, gross margin on the card |
| `假设!I19`, `I5`, `I6`, `I9` | Derived assumptions — must be non-zero |

## 年度数据 beyond column H

Columns `O:T` repeat all 197 history rows as literals (1176 cells). No formula
anywhere in the workbook references `J:T`, so it is a dead copy — but the reset does
not blank it either, so the sample company's numbers stay behind it.

Rows 142-148 are the market block: `历史股价`, `历史股价（前复权）`, then
`历史PE / 历史PS / 历史PB / 历史EV/EBITDA / 历史PCF`. All five are live —
`相对估值` and `图形` both read them — and none can come from a PDF. Four of them
(144/145/146/148) are literals supplied through `params.market.history`. The fifth,
row 147, is a **formula**: the template derives 历史EV/EBITDA as `=财务分析!D65`, and
`相对估值!D82` reaches it through `年度数据!I130`. It is therefore a `formula` row in
`mapping.json`, never written and never reported as a gap.

