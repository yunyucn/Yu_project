"""Excel 写入后端 —— 纯 Python 版（无 COM / 无 Excel / 无 Windows 依赖）。

原版用 COM 驱动本机装的 Excel：只能在 Windows 上跑，而且必须真的装了
Excel。这一版把值直接写进 xlsm 的 XML（engine/xlsm_io.py），公式、VBA、
ActiveX、图表一律原样保留，Linux 上一样跑得动。

一个刻意的差别：公式的**计算结果**不在服务端算。

模板里有 876 个格子参与循环引用（workbook.xml 的 calcPr iterate="1"
就是为它们开的），Excel 从它自己文件里存的迭代起点出发才收敛到它的解，
Python 冷启动算出来的不是同一个数（实测 绝对估值!D40：Excel 3067.60，
冷启动 -936.63）。所以写文件时把所有公式格的缓存值清空，并在
workbook.xml 打上 fullCalcOnLoad="1" —— 办公人员用 Excel 打开时 Excel
自己重算一遍，结果与直接用 Excel 跑这份模板一致。

由此带来的接口契约（上层必须照做，否则会误判）：

  EVALUATES_FORMULAS = False
      公式格读回来的 Value 一律是 None，因为服务端没有算过它。
      「某个公式格的值非零 / 不等于某个数」这类断言在这里无法成立，
      不能当失败处理，只能标成「由 Excel 重算」（见 write_model.py）。
      文字格、数字格（也就是我们真正写进去的数据）读回来是准的。
"""
import contextlib
import datetime
import hashlib
import os
import re
import stat
import sys
import zipfile

# 本文件在 <项目>/pipeline/scripts/ 下，engine 在 <项目>/engine/
_PROJECT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if _PROJECT not in sys.path:
    sys.path.insert(0, _PROJECT)
from engine import Workbook as _Xlsm  # noqa: E402

HAS_EXCEL = False
EVALUATES_FORMULAS = False

XL_CALC_MANUAL = -4135
XL_CALC_AUTO = -4105
XL_CELL_CONSTANTS = 2

_ERR_STRINGS = ("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!",
                "#N/A", "#GETTING_DATA")


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def verify_template(path, expected, allow_mismatch):
    """Refuse to work on a template other than the one the dictionary describes.

    config/mapping.json is keyed by 年度数据 row numbers. A different build of the
    workbook moves rows and defined names; the run then writes every value to the
    wrong row and reports nothing, because the template raises no error and a
    plausible number still lands in every cell. The recorded sha256 was in the
    config all along and nothing ever checked it, so this is the check.
    """
    actual = sha256_file(path)
    if expected and actual != expected and not allow_mismatch:
        raise RuntimeError(
            "template mismatch: %s\n  expected sha256 %s\n  actual   sha256 %s\n"
            "The mapping dictionary is keyed by row number and cannot be trusted "
            "against a different build. Recalibrate config/mapping.json against "
            "this workbook, or pass allow_template_mismatch to proceed anyway "
            "(the row numbers may be wrong and nothing will warn you)."
            % (os.path.basename(path), expected, actual))
    return actual


def make_writable(path):
    """Clear the read-only bit the shipped template carries.

    shutil.copy2 propagates it, and a read-only copy then silently refuses to be
    rewritten (the file hash stays identical and no error is raised). Always call
    this on a copy before writing to it.
    """
    if os.path.exists(path):
        os.chmod(path, os.stat(path).st_mode | stat.S_IWUSR | stat.S_IWGRP)
    return path


def col_letter(n):
    """1 -> A, 27 -> AA. Only for messages; the write paths use numbers."""
    s = ""
    while n:
        n, m = divmod(n - 1, 26)
        s = chr(65 + m) + s
    return s


def col_number(letter):
    n = 0
    for ch in letter:
        n = n * 26 + ord(ch) - 64
    return n


def formula_counts(path):
    """{sheet name: number of formula cells} read straight from the xlsx parts.

    Cheap enough to run on both the pristine template and the written output. The
    openpyxl equivalent costs ~5 s per workbook and tells the same story; this
    compares *counts* per sheet, which is what catches "a formula block was
    overwritten by an array write" without caring which cell it was.
    """
    with zipfile.ZipFile(path) as z:
        wb = z.read("xl/workbook.xml").decode("utf-8")
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        rid_to_target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
        out = {}
        for name, rid in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wb):
            target = rid_to_target.get(rid, "")
            part = "xl/" + target.lstrip("/") if not target.startswith("xl/") else target
            try:
                xml = z.read(part).decode("utf-8")
            except KeyError:
                continue
            # <f> and <f ...> are formula cells. <formula> (inside cfRule and
            # dataValidation) must not be counted, hence the delimiter.
            out[name] = len(re.findall(r"<f[ />]", xml))
        return out


def formula_cells(path):
    """{'工作表!D40'} —— 这份文件里所有公式格。用来判断某个格子的值能不能在
    服务端读出来：公式格不能，只有 Excel 打开时才算得出来。"""
    book = _Xlsm(path)
    return {"%s!%s" % k for k in book.formulas()}


ERR_BASE = -2146826288  # CVErr(2000) == #NULL!


def is_error(x):
    """True for an Excel error value.

    COM hands error values back as negative CVErr integers, not as strings:
    #DIV/0! is -2146826281, #REF! -2146826265, #VALUE! -2146826273. A string-only
    check therefore reports a clean workbook while it is full of #DIV/0!, which is
    how a blanked history block silently destroys 相对估值.
    """
    if isinstance(x, str):
        return x.startswith("#")
    if isinstance(x, bool):
        return False
    return isinstance(x, (int, float)) and ERR_BASE <= x < ERR_BASE + 100


# --------------------------------------------------------------------------
# 工作簿对象：只提供调用方真正用到的那一小撮 Excel 对象模型
# --------------------------------------------------------------------------
class _Range(object):
    __slots__ = ("book", "sheet", "r1", "c1", "r2", "c2")

    def __init__(self, book, sheet, r1, c1, r2=None, c2=None):
        self.book, self.sheet = book, sheet
        self.r1, self.c1 = r1, c1
        self.r2, self.c2 = (r1, c1) if r2 is None else (r2, c2)

    # COM 返回单格标量、多格二维元组，调用方就是按这个写的
    @property
    def Value(self):
        if self.r1 == self.r2 and self.c1 == self.c2:
            return self.book.value_at(self.sheet, self._ref(self.r1, self.c1))
        return tuple(tuple(self.book.value_at(self.sheet, self._ref(r, c))
                           for c in range(self.c1, self.c2 + 1))
                     for r in range(self.r1, self.r2 + 1))

    @Value.setter
    def Value(self, val):
        if self.r1 == self.r2 and self.c1 == self.c2:
            self.book.set_value(self.sheet, self._ref(self.r1, self.c1), val)
            return
        rows = val if isinstance(val, (list, tuple)) else [[val]]
        for i, r in enumerate(range(self.r1, self.r2 + 1)):
            line = rows[i] if i < len(rows) else ()
            if not isinstance(line, (list, tuple)):
                line = (line,)
            for j, c in enumerate(range(self.c1, self.c2 + 1)):
                v = line[j] if j < len(line) else None
                self.book.set_value(self.sheet, self._ref(r, c), v)

    @property
    def Formula(self):
        if self.r1 == self.r2 and self.c1 == self.c2:
            return self.book.formula_at(self.sheet, self._ref(self.r1, self.c1))
        return tuple(tuple(self.book.formula_at(self.sheet, self._ref(r, c))
                           for c in range(self.c1, self.c2 + 1))
                     for r in range(self.r1, self.r2 + 1))

    @Formula.setter
    def Formula(self, val):
        raise NotImplementedError(
            "这个后端不改公式，只改值。要改公式请改模板本身。")

    @staticmethod
    def _ref(r, c):
        return "%s%d" % (col_letter(c), r)

    def ClearContents(self):
        """清掉范围内的字面值，公式一律不动。"""
        for r in range(self.r1, self.r2 + 1):
            for c in range(self.c1, self.c2 + 1):
                ref = self._ref(r, c)
                if (self.sheet, ref) not in self.book._formulas:
                    self.book.set_value(self.sheet, ref, None)

    def SpecialCells(self, kind):
        # 只被 blank_literals 用来「清掉范围内的字面值」，和 ClearContents 同义
        return self


class _Sheet(object):
    def __init__(self, book, name):
        self.book, self.Name = book, name

    def Range(self, addr, end=None):
        if end is not None:
            # ws.Range(ws.Cells(3, 3), ws.Cells(3, 60))：两端都是 Range 对象
            r1, c1 = addr.r1, addr.c1
            r2, c2 = end.r2, end.c2
            return _Range(self.book, self.Name, r1, c1, r2, c2)
        m = re.match(r"^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$", str(addr).strip())
        if not m:
            raise ValueError("看不懂的单元格地址：%r" % (addr,))
        r1, c1 = int(m.group(2)), col_number(m.group(1))
        if m.group(3):
            r2, c2 = int(m.group(4)), col_number(m.group(3))
        else:
            r2, c2 = r1, c1
        return _Range(self.book, self.Name, r1, c1, r2, c2)

    def Cells(self, row, col):
        return _Range(self.book, self.Name, int(row), int(col))

    @property
    def UsedRange(self):
        cells = self.book.cells_of(self.Name)
        if not cells:
            return _Range(self.book, self.Name, 1, 1)
        rows, cols = [], []
        for ref in cells:
            m = re.match(r"([A-Z]+)(\d+)$", ref)
            rows.append(int(m.group(2)))
            cols.append(col_number(m.group(1)))
        return _Range(self.book, self.Name, min(rows), min(cols), max(rows), max(cols))


class _Workbook(object):
    def __init__(self, path):
        self.path = os.path.abspath(path)
        self._src = _Xlsm(self.path)
        self._formulas = self._src.formulas()
        self._base = self._src.all_values()
        self._writes = {}
        self._sheets = {n: _Sheet(self, n) for n in self._src.sheet_names()}
        self.ReadOnly = False

    # ---- 调用方用到的接口 ----
    def Worksheets(self, name):
        if name not in self._sheets:
            raise KeyError("没有这张表：%s" % name)
        return self._sheets[name]

    @property
    def Sheets(self):
        return [self._sheets[n] for n in self._src.sheet_names()]

    def Save(self):
        values = {}
        # 公式格的缓存值全部清掉：服务端算不出循环引用那块，留着旧缓存
        # （模板自带样板公司的数）比留空更危险 —— 打开它的工具不重算时，
        # 用户看到的是一个像模像样但是错的数。
        for cell in self._formulas:
            values[cell] = None
        for (sheet, ref), v in self._writes.items():
            if (sheet, ref) in self._formulas:
                continue          # 公式格不接受字面值覆盖
            values[(sheet, ref)] = v
        self._src.save(self.path, values, full_calc_on_load=True)

    def Close(self, SaveChanges=False):
        # 真的把文件句柄放掉：Windows 上调用方紧接着就要重命名这份产出，
        # 句柄还开着的时候改名会直接失败。
        self._src.close()

    # ---- 取值 / 写值 ----
    def value_at(self, sheet, ref):
        if (sheet, ref) in self._writes:
            return self._writes[(sheet, ref)]
        if (sheet, ref) in self._formulas:
            return None          # 公式值由 Excel 算，服务端没有
        return self._base.get((sheet, ref))

    def formula_at(self, sheet, ref):
        f = self._formulas.get((sheet, ref))
        if f is not None:
            return "=" + f
        return self.value_at(sheet, ref)

    def set_value(self, sheet, ref, val):
        if (sheet, ref) in self._formulas:
            raise RuntimeError(
                "%s!%s 是模板自己的公式，不能用值覆盖。真要在那里写字面值，"
                "得先改模板。" % (sheet, ref))
        if isinstance(val, float) and val == int(val) and abs(val) < 1e15:
            val = int(val)
        self._writes[(sheet, ref)] = val

    def drop_formula(self, sheet, ref):
        self._formulas.pop((sheet, ref), None)

    # ---- 模板导航用的极小公式求值 ------------------------------------------
    # 这不是在服务端复现 Excel 的估值计算（那个明确不做，见模块开头的说明）。
    # 它只回答一个导航问题：「报告期历史数据 的这一列是哪一季」。列是由表头
    # 自己写的日期决定的，而表头是
    #     =CONCATENATE(数据输入!O5,"-12-","31")
    # 数据输入!O2 = YEAR(目录!C8)、目录!C8 = TODAY()-1，所以窗口随日历滑动。
    # 认不出的公式一律抛错：把某一季的数写进另一年的列，工作簿不会报任何错。
    def _today(self):
        if getattr(self, "_today_cache", None) is None:
            self._today_cache = datetime.date.today()
        return self._today_cache

    @staticmethod
    def _split_top(text, seps):
        """从右往左找不在括号 / 引号里的分隔符，返回 (左, 分隔符, 右) 或 None。"""
        depth, inq = 0, False
        for i in range(len(text) - 1, -1, -1):
            ch = text[i]
            if ch == '"':
                inq = not inq
                continue
            if inq:
                continue
            if ch == ")":
                depth += 1
            elif ch == "(":
                depth -= 1
            elif depth == 0 and ch in seps:
                if ch in "+-" and (i == 0 or text[i - 1] in "+-*/^&(,"):
                    continue                      # 一元正负号，不是运算符
                return text[:i], ch, text[i + 1:]
        return None

    def _split_args(self, text):
        out, depth, inq, start = [], 0, False, 0
        for i, ch in enumerate(text):
            if ch == '"':
                inq = not inq
            elif inq:
                continue
            elif ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif ch == "," and depth == 0:
                out.append(text[start:i])
                start = i + 1
        if text.strip() or out:
            out.append(text[start:])
        return out

    @staticmethod
    def _serial_to_date(n):
        return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(n))

    def _eval(self, sheet, expr, depth=0):
        if depth > 12:
            raise ValueError("公式嵌套太深：%r" % expr)
        e = (expr or "").strip()
        if e.startswith("="):
            e = e[1:]
        if not e:
            raise ValueError("空公式")
        if e.startswith('"') and e.endswith('"') and len(e) >= 2:
            return e[1:-1].replace('""', '"')
        for cast in (int, float):
            try:
                return cast(e)
            except ValueError:
                pass
        m = re.match(r"^([A-Za-z][A-Za-z0-9_.]*)\((.*)\)$", e, re.S)
        if m:
            name, inner = m.group(1).upper(), m.group(2)
            args = self._split_args(inner) if inner.strip() else []
            if name == "TODAY" and not args:
                return self._today()
            if name == "YEAR" and len(args) == 1:
                v = self._eval(sheet, args[0], depth + 1)
                if isinstance(v, datetime.date):
                    return v.year
                if isinstance(v, (int, float)):
                    return self._serial_to_date(v).year
                raise ValueError("YEAR() 认不出 %r" % (v,))
            if name == "CONCATENATE" and args:
                return "".join(_text(self._eval(sheet, a, depth + 1)) for a in args)
            raise ValueError("认不出的函数：%s" % name)
        sp = self._split_top(e, "+-&")
        if sp:
            left, op, right = sp
            a = self._eval(sheet, left, depth + 1)
            b = self._eval(sheet, right, depth + 1)
            if op == "&":
                return _text(a) + _text(b)
            if isinstance(a, datetime.date) or isinstance(b, datetime.date):
                base = a if isinstance(a, datetime.date) else b
                n = b if isinstance(a, datetime.date) else a
                if not isinstance(n, (int, float)):
                    raise ValueError("日期只能加减天数，得到 %r" % (n,))
                return base + datetime.timedelta(days=int(n) * (1 if op == "+" else -1))
            if not isinstance(a, (int, float)) or not isinstance(b, (int, float)):
                raise ValueError("%r %s %r 算不了" % (a, op, b))
            return (a + b) if op == "+" else (a - b)
        m = re.match(r"^'?([^'!]*)'?!?([A-Z]{1,3})(\d+)$", e)
        if m:
            sh = m.group(1) or sheet
            ref = m.group(2) + m.group(3)
            f = self._formulas.get((sh, ref))
            if f is not None:
                return self._eval(sh, f, depth + 1)
            v = self.value_at(sh, ref)
            if v is None:
                raise ValueError("%s!%s 是空的，表头日期解不出来" % (sh, ref))
            return v
        raise ValueError("认不出的公式：%r" % expr)

    def formula_refs(self):
        """{'工作表!D40'} —— 这份工作簿里所有公式格。它们的值要 Excel 才算得出来。"""
        return {"%s!%s" % k for k in self._formulas}

    def cells_of(self, sheet):
        out = {r for (s, r) in self._formulas if s == sheet}
        out |= {r for (s, r) in self._base if s == sheet}
        out |= {r for (s, r) in self._writes if s == sheet}
        return out


@contextlib.contextmanager
def excel_app(visible=False):
    """原版这里起一个私有 Excel 实例。纯 Python 后端不需要它，但保留这个
    上下文管理器，四个调用方的 with 语句就不用改。"""
    class _App(object):
        Calculation = XL_CALC_AUTO

        def CalculateFull(self):
            pass

    yield _App()


def open_writable(app, path):
    return _Workbook(path)


# --------------------------------------------------------------------------
# 调用方用的动作，语义与原版一致
# --------------------------------------------------------------------------
def _text(v):
    if isinstance(v, datetime.date):
        return v.isoformat()
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v)


def header_dates(ws, row, first_col, last_col):
    """读一行「日期表头」，公式的也算。返回 {日期: 列号}。

    原版靠 Excel 算好这些表头再读 .Value；这里用上面那个极小求值器算，因为
    整张表的列是靠表头自己写的日期定位的，找不到列就等于静默丢掉一整列数据。
    """
    out = {}
    for c in range(first_col, last_col + 1):
        ref = "%s%d" % (col_letter(c), row)
        f = ws.book._formulas.get((ws.Name, ref))
        v = ws.book._eval(ws.Name, f) if f is not None else ws.book.value_at(ws.Name, ref)
        if v in (None, ""):
            continue
        out[_text(v)[:10]] = c
    return out


def write_block(ws, first_row, last_row, col, values_by_row):
    """整列写。和原版一样是「无条件覆盖」：块内的公式会被写掉。"""
    for r in range(first_row, last_row + 1):
        ref = "%s%d" % (col_letter(col), r)
        ws.book.drop_formula(ws.Name, ref)
        ws.book.set_value(ws.Name, ref, values_by_row.get(r))


def write_block_keep_formulas(ws, first_row, last_row, col, values_by_row):
    """整列写，但块内模板自己的公式原样留着。

    年度数据 的 3-224 行不是纯字面值：147 行是 =财务分析!D65（历史
    EV/EBITDA），相对估值!D82 经 年度数据!I130 读它。无条件整列赋值会把
    它写成空，而空不报错 —— 相对估值 就永远少一条线。

    原版靠 COM 的 .Formula 读一遍再写回去；这里公式本来就还在文件里，
    所以只要「跳过去不写」就够了。返回值保持原样：(行号, "=公式")。
    """
    kept = []
    for r in range(first_row, last_row + 1):
        ref = "%s%d" % (col_letter(col), r)
        f = ws.book._formulas.get((ws.Name, ref))
        if f is not None:
            kept.append((r, "=" + f))
            continue
        ws.book.set_value(ws.Name, ref, values_by_row.get(r))
    return kept


def write_literal(ws, addr, value):
    """写标量格，但拒绝覆盖公式。

    这个模板上每个标量目标格今天都是字面值，但「今天恰好是」不是保证：
    换个版本的模板可能把它做成公式，而一句 .Value = x 会把它抹掉，还留下
    一个看着挺对的数。抛异常把静默错误变成响亮的失败。
    """
    cur = ws.Range(addr).Formula
    if isinstance(cur, str) and cur.startswith("="):
        raise RuntimeError(
            "%s!%s holds the template formula %s; refusing to overwrite it with a "
            "value. If the template really should carry a literal here, remove the "
            "formula deliberately." % (ws.Name, addr, cur))


def read_block(ws, first_row, last_row, col):
    return {r: ws.Cells(r, col).Value for r in range(first_row, last_row + 1)}


def blank_literals(ws, addr):
    """Clear literal cells in a range, leaving formulas untouched."""
    try:
        ws.Range(addr).SpecialCells(XL_CELL_CONSTANTS).ClearContents()
    except Exception:
        pass  # no literal cells in range


def recalc(app, rounds=2):
    """原版是 app.CalculateFull()。公式值不在这里算，见模块开头的说明。"""
    return None


def scan_errors(wb):
    """Count error values across the workbook.

    公式格的错误值要等 Excel 算完才存在，这里数不到；能数到的是我们写进去的
    字面值里本来就带 #DIV/0! 这种字符串的格子。调用方必须把「一个错误都没有」
    理解成「服务端可见的范围内没有」，而不是「整本工作簿是干净的」。
    """
    total = 0
    per_sheet = {}
    for sh in wb.Sheets:
        n = 0
        try:
            v = sh.UsedRange.Value
            if isinstance(v, tuple):
                for row in v:
                    if isinstance(row, tuple):
                        for x in row:
                            if is_error(x):
                                n += 1
                    elif is_error(row):
                        n += 1
            elif is_error(v):
                n = 1
        except Exception:
            pass
        if n:
            per_sheet[sh.Name] = n
            total += n
    return total, per_sheet