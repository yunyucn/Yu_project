"""上传后立刻能算出来的三个「行情」数，其实就写在财报里。

行情里真正买不到的只有股价和历史倍数。总股本写在合并资产负债表的「股本」
（实收资本）行，流通股写在「股份变动情况」表的「无限售条件股份」行，股票
代码写在年报「第二节 公司简介」里。这三样能算出来的，就不该让办公人员去查、
去抄 —— 抄错一位，整份估值就偏了，而工作簿不会报错。

解析复用 skill 自己的 parse_pdf.py，并和流水线共用同一个解析缓存：预填解析过
的 PDF，正式流水线直接从缓存读，不会重复解析。

这里还算历史估值倍数（PE/PS/PB/PCF）：它们等于「年末总市值 ÷ 当年的财务数」，
而财务数就在年报里。所以办公人员只需要查 6 个年末收盘价，倍数由系统算出来，
不必再去行情软件里抄 24 个数。
"""
import io
import json
import os
import re
import subprocess
import sys

import config
import logging
import runtime

# 合并资产负债表里代表股本的标签。A 股面值 1 元，所以「元」除以 1e6 就是百万股。
SHARE_LABELS = ("实收资本", "股本")

# 年报封面不写代码，代码在「公司简介」那一页；扫前若干页就够，不必扫完 265 页。
#
# 这个正则必须比「看到 股票代码 就取后面六位」严得多。年报的「释义」一节会把
# 供应商和同行也写进去，格式一模一样：「三星电子 指 韩国证券交易所上市公司，
# 股票代码005930.KS」。只认六位数字的话，江波龙的年报会被读成三星（005930），
# 而工作簿不会报错 —— 这正是要防的那类静默错误。
#
# 两道闸：
#   1. 只在同时出现「简称」的页上取（「股票简称 江波龙 股票代码 301308」，
#      这是公司简介一节的固定写法，释义一节没有「简称」）；
#   2. 代码后面不许跟 `.`（境外上市写作 005930.KS / 8299.TWO，A 股代码后面
#      不会跟小数点）。
CODE_RE = re.compile(r"(?:证券代码|股票代码)\s*[：:]?\s*(\d{6})(?![\d.])")
CODE_HEAD_PAGES = 20

# 流通股那张表在「股份变动及股东情况」一节，通常在报告前 2/3。设上限是为了
# 万一认不出时别把整本报告翻完。
FLOAT_PAGE_CAP = 200


def _py():
    return sys.executable


def _env():
    return dict(os.environ, PYTHONIOENCODING="utf-8")


def _norm(s):
    """和 map_fields.norm 同一套：去空白、去括号后内容、去编号与标点。"""
    s = re.sub(r"[\s\u3000]", "", str(s or ""))
    s = re.split(r"[（(]", s)[0]
    s = re.sub(r"^[\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]+[\u3001.]?", "", s)
    s = re.sub(r"^(\u51cf|\u52a0|\u5176\u4e2d)[:\uff1a]?", "", s)
    s = re.sub(r"[:\uff1a,\uff0c\u3002\u3001_\-\uff0d\u2014]", "", s)
    s = re.sub(r"\d+(\.\d+)?$", "", s)
    s = re.sub(r"^\d+(\.\d+)?", "", s)
    return s.strip()


def parse_reports(reports, workdir, cache_dir=None):
    """调用 skill 的 parse_pdf.py。reports: [{"pdf","period","kind"}]。

    返回 {pdf 路径: 解析结果}。失败返回空 dict —— 预填是加分项，不能因为它
    挂了就让整个上传失败。
    """
    if not reports:
        return {}
    params_path = os.path.join(workdir, "_prefill_params.json")
    out_path = os.path.join(workdir, "_prefill_extract.json")
    with io.open(params_path, "w", encoding="utf-8") as fh:
        json.dump({"reports": reports}, fh, ensure_ascii=False, indent=1)
    cmd = [_py(), config.PARSE_PDF, "--params", params_path,
           "--pages-config", config.PDF_PAGES, "--out", out_path]
    if cache_dir:
        cmd += ["--cache-dir", cache_dir]
    returncode, tail = runtime.run_command(
        cmd, os.path.join(workdir, '_prefill.log'), config.PREFILL_TIMEOUT_SECONDS,
        context={'task_id': os.path.basename(workdir)})
    if returncode != 0 or not os.path.exists(out_path):
        logging.getLogger(__name__).warning('prefill_parse_failed',
            extra={'task_id': os.path.basename(workdir), 'exit_code': returncode})
        return {}
    try:
        with io.open(out_path, encoding="utf-8") as fh:
            data = json.load(fh)
    except ValueError:
        logging.getLogger(__name__).exception('prefill_json_invalid',
            extra={'task_id': os.path.basename(workdir)})
        return {}
    # 按输入顺序回填，不按结果里的 "pdf" 字段：命中解析缓存时，缓存里存的是
    # 首次解析时的路径，和这次传进来的路径可能不是同一个字符串（同一份 PDF
    # 从「测试财报」换成「财报」目录就会命中缓存），按路径取会一个都取不到。
    sources = data.get("sources", [])
    out = {}
    for rep, src in zip(reports, sources):
        out[rep["pdf"]] = src
    for src in sources:
        out.setdefault(os.path.basename(src.get("pdf") or ""), src)
    return out


def total_shares_mn(source):
    """合并资产负债表「股本」行的期末数（元）-> 百万股。"""
    rows = ((source or {}).get("statements") or {}).get("合并资产负债表") or []
    for row in rows:
        if _norm(row.get("label")) in SHARE_LABELS:
            nums = row.get("nums") or []
            if nums:
                return round(nums[0] / 1e6, 6)
    return None


def _to_int(cell):
    if cell is None:
        return None
    s = re.sub(r"[\s,]", "", str(cell))
    return int(s) if re.fullmatch(r"\d+", s) else None


def float_shares_mn(pdf_path):
    """「股份变动情况」表：无限售条件股份的「本次变动后」数量 -> 百万股。

    这张表跨页、单元格里的数字还带换行，按行号取容易取错列，所以用 pdfplumber
    的表格还原，再按「本次变动后」那一列的数值反着找第一个纯数字单元格。
    """
    try:
        import pdfplumber
    except Exception:
        return None
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[:FLOAT_PAGE_CAP]:
                text = page.extract_text() or ""
                # 行名在表里是折行、而且数字还夹在中间（「二、无限 / 1,552,914 /
                # 售条件股 …」），所以「无限售条件股」这个整串在原文里根本不连续。
                # 只认「售条件股」这个不会出现在别处的片段，再配上表头「本次变动」。
                flat = re.sub(r"\s", "", text)
                if "售条件股" not in flat or "本次变动" not in flat:
                    continue
                table = page.extract_table()
                if not table:
                    continue
                for row in table:
                    cells = row or []
                    # 标签不在固定的列上：兴源环境在第 1 列，江波龙前面多一个
                    # 空列、标签落在第 2 列。整行找，不认死列号。
                    # 必须用完整的「无限售条件股」：「有限售条件股份」也含
                    # 「售条件股」，松一档就会把限售股当成流通股（差了十倍）。
                    # 表格单元格里这个串是完整的，只有 extract_text 才会被折行拆开。
                    labels = [_norm(c) for c in cells if c and _to_int(c) is None]
                    if not any("无限售条件股" in x for x in labels):
                        continue
                    # 「本次变动后」的数量是这一行最后一个纯数字格。比例带 %，
                    # _to_int 不认，所以倒着找第一个整数就是它。
                    for cell in reversed(cells):
                        n = _to_int(cell)
                        if n:
                            return round(n / 1e6, 6)
    except Exception:
        return None
    return None


# 算历史倍数要用的四个财务数，行号取自 mapping.json（不写死标签，
# 标签的匹配规则用 skill 自己的 map_fields.lookup，保证和流水线同一套口径）。
FINANCIAL_ROWS = {"revenue": 3, "net_profit": 30, "equity": 129, "ocf": 161}


def _financial_specs():
    """(行号 -> 行定义, 元转百万的缩放)。"""
    with io.open(config.MAPPING, encoding="utf-8") as fh:
        m = json.load(fh)
    by_row = {r.get("row"): r for r in (m.get("rows") or [])}
    scale = (m.get("units") or {}).get("scale") or 1.0
    return by_row, scale


def financials_by_year(sources):
    """{年份: {"revenue","net_profit","equity","ocf"}}，单位百万元。

    一份年报提供两年（期末 = 报告年，期初 = 上一年），所以三份隔年年报就能
    覆盖模板要的六年。取不到的年份值一律 None —— 倍数宁可不给，也不能拿 0 算。
    """
    if not sources:
        return {}
    import sys
    if config.SCRIPTS_DIR not in sys.path:
        sys.path.insert(0, config.SCRIPTS_DIR)
    import map_fields

    by_row, scale = _financial_specs()
    idx = map_fields.index_sources(sources)
    years = set()
    for src in sources:
        y = int(str(src["period"])[:4])
        years.update((y, y - 1))

    out = {}
    for year in sorted(years):
        rec = {}
        for key, row_no in FINANCIAL_ROWS.items():
            spec = by_row.get(row_no)
            if not spec:
                rec[key] = None
                continue
            value, _ev = map_fields.lookup(
                idx, spec.get("statement"),
                spec.get("match") or [spec.get("label")], year)
            rec[key] = round(value * scale, 6) if value is not None else None
        out[year] = rec
    return out


def stock_code(pdf_path):
    """扫前若干页找公司简介里的「股票简称 X 股票代码 YYYYYY」。

    认不出就返回 None。宁可留空让办公人员填，也不能猜一个 —— 猜错六位数字，
    整份估值都挂在别人公司的股本上。
    """
    try:
        import pdfplumber
    except Exception:
        return None
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[:CODE_HEAD_PAGES]:
                text = page.extract_text() or ""
                # 闸门 1：这一页必须同时出现「简称」，才可能是公司简介一节。
                if "简称" not in text:
                    continue
                m = CODE_RE.search(text)
                if m:
                    return m.group(1)
    except Exception:
        return None
    return None


# 历史倍数的四个字段，以及各自的分母取哪个财务数。
HISTORY_FIELDS = ("PE", "PS", "PB", "PCF")
HISTORY_BASIS = {"PE": "net_profit", "PS": "revenue",
                 "PB": "equity", "PCF": "ocf"}


def history_from_prices(prices, financials, shares_mn):
    """{年份: 年末收盘价} + 各年财务数 + 总股本 -> {期间: {PE,PS,PB,PCF}}。

    口径与模板一致（模板自己就是这么算现值的）：
      PE  = 总市值 / 归母净利润      PS  = 总市值 / 营业收入
      PB  = 总市值 / 归母净资产      PCF = 总市值 / 经营活动现金流量净额
    总市值 = 年末收盘价 × 总股本（都是百万元）。分母缺失或为 0 的那一项留空 ——
    宁可不给，也不能拿 0 算出一个假的倍数。
    """
    if not shares_mn:
        return {}
    try:
        shares_mn = float(shares_mn)
    except (TypeError, ValueError):
        return {}
    out = {}
    for year, price in (prices or {}).items():
        try:
            price = float(str(price).replace(",", "").strip())
        except (TypeError, ValueError):
            continue
        if price <= 0:
            continue
        # 内存里键可能是 int，存进 JSON 又变成 str，两种都认。
        fin = ((financials or {}).get(str(year))
               or (financials or {}).get(year) or {})
        cap = price * shares_mn
        vals = {}
        for field, key in HISTORY_BASIS.items():
            base = fin.get(key)
            if base in (None, 0):
                continue
            vals[field] = round(cap / float(base), 4)
        if vals:
            out["%s-12-31" % year] = vals
    return out


def prefill(reports, workdir, cache_dir=None):
    """reports 是 detect 的结果：[{"pdf","period","kind","code"}]。

    返回 {"total_shares_mn", "float_shares_mn", "stock_code", "source"}，
    认不出来的字段一律 None —— 绝不能拿 0 冒充「不知道」。
    """
    reports = [r for r in (reports or []) if r.get("pdf")]
    out = {"total_shares_mn": None, "float_shares_mn": None,
           "stock_code": None, "financials": {}, "source": {}}

    # 股票代码：先用 detect 已经认出来的，认不出再自己扫（年报封面不写代码）。
    for r in reports:
        if r.get("code"):
            out["stock_code"] = r["code"]
            out["source"]["stock_code"] = os.path.basename(r["pdf"])
            break
    if not out["stock_code"]:
        for r in reports:
            c = stock_code(r["pdf"])
            if c:
                out["stock_code"] = c
                out["source"]["stock_code"] = os.path.basename(r["pdf"])
                break

    annuals = [r for r in reports if r.get("kind") == "annual"]
    annuals.sort(key=lambda r: str(r.get("period") or ""), reverse=True)
    if not annuals:
        return out

    # 全部年报都要解析：股本只要最新那份，但历史倍数要按年算，六个年份的
    # 财务数散在各年年报里（一份年报含两年）。解析结果进同一个缓存，
    # 正式流水线第 3 步会直接命中，不会重复解析。
    parsed = parse_reports(
        [{"pdf": r["pdf"], "period": r.get("period"), "kind": "annual"}
         for r in annuals], workdir, cache_dir)

    latest = annuals[0]
    src = parsed.get(latest["pdf"])
    total = total_shares_mn(src)
    if total is not None:
        out["total_shares_mn"] = total
        out["source"]["total_shares_mn"] = os.path.basename(latest["pdf"])
    fl = float_shares_mn(latest["pdf"])
    if fl is not None:
        out["float_shares_mn"] = fl
        out["source"]["float_shares_mn"] = os.path.basename(latest["pdf"])

    out["financials"] = financials_by_year(
        [parsed[r["pdf"]] for r in annuals if r["pdf"] in parsed])
    return out
