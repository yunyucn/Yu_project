# -*- coding: utf-8 -*-
# 纯 Python 读写 xlsm/xlsx：直接操作 zip 内的 sheet XML。
# 不依赖 Excel、pywin32、COM，Windows / Linux 都能跑。
#
# 写回时：
#   公式格 -> 只动 <v> 缓存值，<f> 原样保留
#   普通格 -> 写成字面值（数字 / 布尔 / inlineStr 字符串 / 错误值）
#   目标格在 XML 里不存在 -> 就地新建（不建的话写进去的值会静默消失）
#   VBA、ActiveX、图表、样式等部件一律原样复制

import os
import re
import zipfile
from xml.sax.saxutils import escape, unescape

_MISSING = object()

_REF_SHIFT = re.compile(
    r"(?<![A-Za-z0-9_.!$])(\$?)([A-Z]{1,3})(\$?)([1-9][0-9]{0,6})(?![A-Za-z0-9_(])")
# 自闭合格 <c r="C5" s="1"/> 和完整格 <c r="D40">..</c> 都要能匹配。
# 早先用 split('<c r="') 的写法会把连续的自闭合格并进同一块，
# 从而把后面格子的 <f>/<v> 误记到前面的格子上。
_CELL_ITER = re.compile(r'<c\b[^>]*/>|<c\b[^>]*>.*?</c>', re.S)
_CELL_REF = re.compile(r'\br="([A-Z]+[0-9]+)"')
_ROW_ITER = re.compile(r'<row\b[^>]*/>|<row\b[^>]*>.*?</row>', re.S)
_ROW_REF = re.compile(r'<row\b[^>]*?\br="(\d+)"')
_ERR_LIKE = re.compile(r"^#(NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|GETTING_DATA)$")


def col2n(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


def n2col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def split_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)", ref)
    return col2n(m.group(1)), int(m.group(2))


def shift_formula(f, dr, dc):
    # 展开共享公式：把主公式里的相对引用按偏移改写
    parts = f.split('"')
    for i in range(0, len(parts), 2):
        def rep(m):
            ac, cl, ar, rw = m.group(1), m.group(2), m.group(3), m.group(4)
            c, r = col2n(cl), int(rw)
            if not ac:
                c += dc
            if not ar:
                r += dr
            if c < 1 or r < 1:
                return "#REF!"
            return ac + n2col(c) + ar + str(r)
        parts[i] = _REF_SHIFT.sub(rep, parts[i])
    return '"'.join(parts)


def parse_cell(txt):
    """把一段 <c .../> 或 <c ...>...</c> 拆成 (ref, attrs, inner, selfclose)。"""
    m = _CELL_REF.search(txt)
    if not m:
        return None, None, None, None
    ref = m.group(1)
    gt = txt.index(">")
    head = txt[1:gt]
    if head.endswith("/"):
        head = head[:-1]
        selfclose = True
        inner = ""
    else:
        selfclose = False
        inner = txt[gt + 1:txt.rindex("</c>")]
    sp = head.find(" ")
    attrs = head[sp + 1:] if sp >= 0 else ""
    return ref, attrs, inner, selfclose


def iter_cells(xml):
    for m in _CELL_ITER.finditer(xml):
        yield m.start(), m.group(0)


def render_cell(ref, attrs, inner, val, has_formula):
    """一个 <c> 元素的序列化。attrs 里带不带 r= / t= 都行，这里统一重建。"""
    base = re.sub(r'\s*r="[^"]*"', "", attrs or "")
    body = re.sub(r"<v>.*?</v>", "", inner or "", flags=re.S)
    body = re.sub(r"<is>.*?</is>", "", body, flags=re.S)

    if val is None and has_formula:
        # 公式格：只删缓存值，t 属性必须原样留着。实测把 t="str" 一起去掉之后
        # Excel 直接打不开这份文件（"类 Workbooks 的 Open 方法无效"），而这份
        # XML 本身是完全合法的 —— 结构校验、openpyxl 都查不出来，只有真的
        # 拿 Excel 打开才会暴露。
        return '<c r="%s"%s>%s</c>' % (ref, base, body)

    base = re.sub(r'\s+t="[^"]*"', "", base)
    if val is None:
        # 清空：普通格整格变空（连带去掉过期的类型标记）
        return '<c r="%s"%s>%s</c>' % (ref, base, body)
    if isinstance(val, bool):
        return '<c r="%s"%s t="b">%s<v>%d</v></c>' % (ref, base, body, 1 if val else 0)
    if isinstance(val, (int, float)):
        return '<c r="%s"%s>%s<v>%s</v></c>' % (ref, base, body, repr(float(val)))
    text = str(val)
    if _ERR_LIKE.match(text):
        return '<c r="%s"%s t="e">%s<v>%s</v></c>' % (ref, base, body, escape(text))
    if has_formula:
        return '<c r="%s"%s t="str">%s<v>%s</v></c>' % (ref, base, body, escape(text))
    return ('<c r="%s"%s t="inlineStr">%s<is><t xml:space="preserve">%s</t></is></c>'
            % (ref, base, body, escape(text)))


def set_full_calc_on_load(wb_xml):
    """让 Excel 打开时无条件重算整本工作簿。

    交付物里的公式缓存值一律是空的（服务端算不出循环引用那块），
    所以必须让 Excel 自己算一遍；calcOnSave="0" 是模板原有的设置，保留。
    """
    m = re.search(r"<calcPr\b[^>]*/>", wb_xml)
    if not m:
        return wb_xml.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
    tag = m.group(0)
    if "fullCalcOnLoad" in tag:
        tag = re.sub(r'\s*fullCalcOnLoad="[^"]*"', "", tag)
    tag = tag[:-2].rstrip() + ' fullCalcOnLoad="1"/>'
    return wb_xml[:m.start()] + tag + wb_xml[m.end():]


class Workbook(object):
    def __init__(self, path):
        self.path = path
        self.zip = zipfile.ZipFile(path)
        self.sheets = self._sheet_map()
        self._xml = {}
        self._shared = None

    def _z(self):
        # 句柄按需开：Windows 上文件还被 zipfile 占着就没法改名 / 替换，
        # 而调用方保存完就要把它搬走。
        if self.zip is None:
            self.zip = zipfile.ZipFile(self.path)
        return self.zip

    def close(self):
        if self.zip is not None:
            self.zip.close()
            self.zip = None

    # ---------- 读取 ----------
    def _sheet_map(self):
        wbx = self.zip.read("xl/workbook.xml").decode("utf-8", "replace")
        rels = self.zip.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
        rmap = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
        out = {}
        for m in re.finditer(r'<sheet[^>]*name="([^"]*)"[^>]*r:id="([^"]+)"', wbx):
            t = rmap.get(m.group(2), "").replace("worksheets/", "xl/worksheets/")
            if not t.startswith("xl/"):
                t = "xl/" + t
            out[unescape(m.group(1))] = t
        return out

    def sheet_names(self):
        return list(self.sheets)

    def xml(self, sheet):
        if sheet not in self._xml:
            self._xml[sheet] = self._z().read(self.sheets[sheet]).decode("utf-8", "replace")
        return self._xml[sheet]

    def shared_strings(self):
        if self._shared is None:
            self._shared = []
            try:
                raw = self._z().read("xl/sharedStrings.xml").decode("utf-8", "replace")
            except KeyError:
                return self._shared
            for si in raw.split("<si>")[1:]:
                si = si.split("</si>")[0]
                self._shared.append("".join(
                    unescape(t) for t in re.findall(r"<t[^>]*>(.*?)</t>", si, re.S)))
        return self._shared

    def formulas(self):
        # {(sheet, ref): 公式文本}，共享公式已展开为每格的实际公式
        out = {}
        for sh in self.sheets:
            x = self.xml(sh)
            shared = {}
            for _pos, txt in iter_cells(x):
                ref, _attrs, inner, selfclose = parse_cell(txt)
                if ref is None or selfclose:
                    continue
                m = re.search(r"<f([^>]*)>(.*?)</f>", inner, re.S)
                m2 = re.search(r"<f([^>]*)/>", inner)
                if m:
                    at, txt = m.group(1), m.group(2)
                    si = re.search(r'si="(\d+)"', at)
                    if si and "shared" in at:
                        shared[si.group(1)] = (ref, txt)
                    out[(sh, ref)] = unescape(txt)
                elif m2:
                    at = m2.group(1)
                    si = re.search(r'si="(\d+)"', at)
                    if si and si.group(1) in shared:
                        mr, mt = shared[si.group(1)]
                        c0, r0 = split_ref(mr)
                        c1, r1 = split_ref(ref)
                        out[(sh, ref)] = unescape(shift_formula(mt, r1 - r0, c1 - c0))
        return out

    @staticmethod
    def _cell_value(attrs, inner, shared):
        # t="s" 这类类型标记在 <c> 标签上，不在标签内部：只把 inner 递进来
        # 会让每个共享字符串格都读成它的索引数字（目录!C6 读成 2886 而不是
        # 股票代码），而且不报错。
        m = re.search(r"<v>(.*?)</v>", inner, re.S)
        raw = unescape(m.group(1)) if m else None
        tm = re.search(r'\bt="([^"]+)"', attrs or "")
        t = tm.group(1) if tm else None
        if t == "inlineStr":
            im = re.search(r"<is>(.*?)</is>", inner, re.S)
            return "".join(unescape(x) for x in
                           re.findall(r"<t[^>]*>(.*?)</t>", im.group(1), re.S)) if im else None
        if raw is None:
            return None
        if t == "s":
            try:
                return shared[int(raw)]
            except (ValueError, IndexError):
                return None
        if t == "str" or t == "e":
            return raw
        if t == "b":
            return raw not in ("0", "false", "FALSE")
        try:
            f = float(raw)
            return int(f) if f.is_integer() and abs(f) < 1e15 else f
        except ValueError:
            return raw

    def all_values(self):
        # {(sheet, ref): 值}，字面值和公式缓存值都包含
        out = {}
        shared = self.shared_strings()
        for sh in self.sheets:
            for _pos, txt in iter_cells(self.xml(sh)):
                ref, _attrs, inner, selfclose = parse_cell(txt)
                if ref is None or selfclose:
                    continue
                v = self._cell_value(_attrs, inner, shared)
                if v is not None:
                    out[(sh, ref)] = v
        return out

    def qualify(self, sheet, ref):
        return "'[%s]%s'!%s" % (os.path.basename(self.path), sheet, ref)

    # ---------- 写入 ----------
    def _patch_existing(self, sheet, x, values):
        def repl(m):
            txt = m.group(0)
            ref, attrs, inner, _sc = parse_cell(txt)
            if ref is None:
                return txt
            val = values.get(ref, _MISSING)
            if val is _MISSING:
                return txt
            return render_cell(ref, attrs, inner, val, "<f" in (inner or ""))
        return _CELL_ITER.sub(repl, x)

    @staticmethod
    def _merge_row(txt, refs_vals):
        """把新格子插进一行里，按列号排序。行里除 <c> 之外不该有别的东西。"""
        if txt.rstrip().endswith("/>"):
            # 空行 <row r="5" s="1"/>：展开成有内容的行
            head = txt.rstrip()[:-2].rstrip() + ">"
            cells = []
            tail = "</row>"
        else:
            head_end = txt.index(">") + 1
            head = txt[:head_end]
            body_end = txt.rindex("</row>")
            body = txt[head_end:body_end]
            leftover = _CELL_ITER.sub("", body).strip()
            if leftover:
                raise RuntimeError("行里出现了非 <c> 内容，拒绝重建：%r" % leftover[:80])
            cells = [(col2n(parse_cell(m.group(0))[0]), m.group(0))
                     for m in _CELL_ITER.finditer(body)]
            tail = "</row>"
        cells += [(col2n(ref), render_cell(ref, "", "", val, False))
                  for ref, val in refs_vals]
        cells.sort(key=lambda t: t[0])
        return head + "".join(t for _c, t in cells) + tail

    @classmethod
    def _add_cells(cls, x, by_row):
        def repl(m):
            txt = m.group(0)
            rm = _ROW_REF.search(txt)
            if not rm:
                return txt
            n = int(rm.group(1))
            if n not in by_row:
                return txt
            return cls._merge_row(txt, by_row.pop(n))
        return _ROW_ITER.sub(repl, x)

    @staticmethod
    def _add_rows(x, by_row):
        """整行都不存在时新建一行，行号必须保持升序。"""
        def build(n):
            refs = by_row.pop(n)
            cells = "".join(render_cell(ref, "", "", val, False) for ref, val in refs)
            return '<row r="%d">%s</row>' % (n, cells)

        m = re.search(r"<sheetData\s*>", x)
        if not m:
            raise RuntimeError("工作表里没有 <sheetData>")
        head_end = m.end()
        body_end = x.rindex("</sheetData>")
        body = x[head_end:body_end]
        rows = [(int(_ROW_REF.search(mm.group(0)).group(1)), mm.group(0))
                for mm in _ROW_ITER.finditer(body)]
        rows += [(n, build(n)) for n in sorted(by_row)]
        rows.sort(key=lambda t: t[0])
        return x[:head_end] + "".join(t for _n, t in rows) + x[body_end:]

    def save(self, out_path, values=None, full_calc_on_load=True):
        """values: {(sheet, ref): 值}；None 表示清空。未给出的格子原样不动。

        目标格在 XML 里不存在时会新建 —— 否则「写进去了」只是内存里的错觉，
        文件里什么都没有，而调用方读回内存还会以为成功了。
        """
        values = values or {}
        patched = {}
        for sh in self.sheets:
            x = self.xml(sh)
            mine = {ref: v for (s, ref), v in values.items() if s == sh}
            if not mine:
                patched[self.sheets[sh]] = x
                continue
            x = self._patch_existing(sh, x, mine)
            have = {parse_cell(m.group(0))[0] for m in _CELL_ITER.finditer(x)}
            by_row = {}
            for ref, val in mine.items():
                if ref in have or val is None:
                    continue          # 清空一个不存在的格子等于没写
                by_row.setdefault(int(re.match(r"[A-Z]+(\d+)", ref).group(1)), []).append((ref, val))
            if by_row:
                x = self._add_cells(x, by_row)
            if by_row:
                x = self._add_rows(x, by_row)
            patched[self.sheets[sh]] = x

        wbx = None
        if full_calc_on_load:
            wbx = set_full_calc_on_load(
                self._z().read("xl/workbook.xml").decode("utf-8", "replace"))

        # 源文件先整个读进内存再关掉句柄：Windows 上文件还被 zipfile 占着的时候
        # 没法替换它，而调用方经常是「就地保存」（out_path 就是自己这份文件）。
        z = self._z()
        infos = z.infolist()
        raw = {it.filename: z.read(it.filename) for it in infos}
        z.close()
        self.zip = None
        tmp = out_path + ".tmp-xlsm"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in infos:
                if wbx is not None and item.filename == "xl/workbook.xml":
                    data = wbx.encode("utf-8")
                else:
                    data = patched.get(item.filename)
                    data = data.encode("utf-8") if data is not None else raw[item.filename]
                zout.writestr(item, data)
        os.replace(tmp, out_path)
        # 就地保存之后对象自己还得能用：下次读的时候会重新打开新文件，
        # 旧的 XML / 共享串缓存必须丢掉。
        self.path = out_path
        self._xml = {}
        self._shared = None
        return out_path