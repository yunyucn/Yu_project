# pipeline —— 财报 PDF 到估值模型的流水线

Web 版（`app.py`）调的就是这里。也可以单独在命令行跑：

    python scripts/run_pipeline.py \
        --params params.json \
        --template ../原件/开源_A股估值模型(非金融)_20240109043703.xlsm \
        --workdir /tmp/run --name "某公司_2025年报_估值模型"

## 十一步做什么

    1  认料      每份 PDF 是什么报告（年报/半年报/季报）、报告期是哪天
    2  清理模板  把样板公司（科顺股份 300737.SZ）的数据清掉，公式一律保留
    3  解析 PDF  抽三张表 + 分红 + 股本；最慢的一步（大年报 1-3 分钟）
    4  字段映射  PDF 里的科目 -> 年度数据 的行号
    5  季度通道  季报/半年报 -> 报告期历史数据（季度粒度那一张）
    6  数据够不够  逐假设判断，缺哪年、缺哪格，会不会把假设打崩
    7  还差什么  列出要补的报告、要人工填的格
    8  写入前校验 会计恒等式、PDF 对账（不过就不写）
    9  写入并保存
    10 生成结论卡  给办公人员看的那一页
    11 验收      产出的工作簿还是不是原件那台引擎

退出码：`0` 完成；`1` 停下且没有产出工作簿；`2` 写出了工作簿但某项检查没过。

## 服务端不做什么（重要）

模板里有 876 个格子参与**循环引用**（`xl/workbook.xml` 的
`calcPr iterate="1"` 就是为它们开的）。实测：Excel 从它自己文件里存的迭代
起点出发才收敛到它的解，Python 冷启动算出来是另一个数（`绝对估值!D40`：
Excel 3067.60，冷启动 -936.63，而且多稳态）。

所以本流水线**不复现**这部分计算：

- 写工作簿时把所有公式格的缓存值清空，并在 `workbook.xml` 里写
  `fullCalcOnLoad="1"`；
- 办公人员用 Excel 打开产出文件时，Excel 自己重算一遍，结果与直接用
  Excel 跑这份模板一致。

由此产生的接口契约，改代码时必须照做：

- `com_excel.EVALUATES_FORMULAS = False`：公式格读回来的值是 `None`。
- 「某个公式格非零 / 不等于某个数」这类写入后断言在服务端**不可判定**，
  流水线把它们标成「待 Excel 重算」而不是失败（见 `write_model.py`）。
  当成失败会让每次正常运行都报假警报，而假警报会训练人忽略真警报。
- 验收里的「同数据同结果」要拿两份工作簿在 Excel 里各算一遍，服务端做不了，
  会明确写成「跳过（没有 Excel 后端）」，不是含糊地算过。

## 写工作簿的后端

`scripts/com_excel.py` 是唯一的出口。它有两套实现的历史：

- 原版用 COM 驱动本机装的 Excel（Windows + 真装了 Excel 才能跑）。
- 现在这份是纯 Python：直接改 xlsm 里的 XML（`../engine/xlsm_io.py`），
  公式、VBA、ActiveX、图表原样保留。`HAS_EXCEL = False`。

它对外提供的动作：`sha256_file` / `verify_template` / `make_writable` /
`excel_app` / `open_writable` / `write_block` / `write_block_keep_formulas` /
`write_literal` / `read_block` / `blank_literals` / `formula_counts` /
`formula_refs` / `header_dates` / `scan_errors` / `is_error` / `col_letter` /
`recalc`。

`header_dates` 值得单独说：`报告期历史数据` 的列是由表头自己写的日期定位的，
而表头是公式（`=CONCATENATE(数据输入!O2,"-03-","31")`，`数据输入!O2 =
YEAR(目录!C8)`，`目录!C8 = TODAY()-1`，窗口随日历滑动）。服务端没有 Excel
算它，所以 `com_excel` 里带了一个**只用于模板导航**的极小求值器
（`TODAY` / `YEAR` / `CONCATENATE` / 单元格引用 / `+ - &`）。认不出的公式
一律抛错 —— 把某一季的数写进另一年的列，工作簿不会报任何错。

## config/ 里都是什么

| 文件 | 管什么 |
|---|---|
| `mapping.json` | PDF 科目 -> `年度数据` 行号；某科目换了写法就往 `match` 列表加词 |
| `mapping_quarterly.json` | 季报/半年报的科目 |
| `pdf_pages.json` | 三张表在年报里大概在哪几页（先定位再抽，省时间） |
| `cleanup.json` | 清理样板公司时要清哪些格、哪些格要留着 |
| `assertions.json` | 写入前后的检查项 |
| `assumption_deps.json` | 65 个派生假设各自需要哪些格（逐格实测出来的） |
| `analyst_inputs.json` | 模板里 6 个「只能由人定」的判断格 |
| `params.example.json` | 参数 JSON 的样例 |

`mapping.json` 的 `template.sha256` 是校准过的那份模板的指纹。模板按指纹认，
不按文件名：文件夹里常有「已解锁」「副本」之类的改写版，按名字挑会挑错，
而挑错模板会让每个值落到错行上，还不报错。

## 维护手册

- `references/template-map.md` —— 模板结构（34 表 / 231 定义名称 / 166 ActiveX / 31 图表）
- `references/field-mapping.md` —— 映射字典的结构与写法
- `references/domain-model.md` —— 领域词汇与规则
- `references/invariants.md` —— 实测出来的失效模式与不变量（改阈值前必读）
