# -*- coding: utf-8 -*-

import re
import os
import sys
import datetime
import functools
import itertools
import collections

import openpyxl
import dateutil.parser as dateparser

from . import utils

Ctx = utils.Ctx


RANGE_SEP = ':'
RANGE_NIL = '...'
COLKEY_TOKEN = '@'
COLKEY_REGEX = \
    re.compile(r'^(@[^,:]+)$', re.IGNORECASE)
CELLXX_REGEX = \
    re.compile(r'^([A-Z]+)([1-9][0-9]*)$', re.IGNORECASE)
COLUMN_REGEX = \
    re.compile(r'^([A-Z]+)$', re.IGNORECASE)
VINDEX_REGEX = \
    re.compile(r'^([1-9][0-9]*)$', re.IGNORECASE)
RANGE1_REGEX = \
    re.compile(
        r'^(?:(@[^,:]+)(,[1-9][0-9]*)?|([A-Z]+)?([1-9][0-9]*)?)'
        r':(?:(@[^,:]+)(,[1-9][0-9]*)?|([A-Z]+)?([1-9][0-9]*)?)$',
        re.IGNORECASE)
RANGE2_REGEX = \
    re.compile(
        r'^(?:(@[^,:]+)|([A-Z]+))$',
        re.IGNORECASE)
RANGE3_REGEX = VINDEX_REGEX

DEFAULT_DATETIME = datetime.datetime(1900, 1, 1)


def build_column(col):
    digits = []
    while col > 0:
        digits.append(chr((col - 1) % 26 + 65))
        col //= 26
    return ''.join(reversed(digits))


def parse_column(expr):
    if isinstance(expr, str):
        matches = COLUMN_REGEX.match(expr.upper())
        if matches:
            col = 0
            for ich in matches.group(1):
                col = col * 26 + ord(ich) - 64
            return col
    Ctx.throw('错误的列号格式：{0!r}', expr)


RangeArgs = collections.namedtuple(
    'RangeArgs',
    ['hoff', 'voff', 'hnum', 'vnum'],
)


def range_args(lcol, lrow, hcol, hrow):
    return RangeArgs(lcol - 1, lrow - 1, hcol - lcol + 1, hrow - lrow + 1)


def range_expr(lcol, lrow, hcol, hrow):
    return RANGE_SEP.join([
        build_column(lcol) + str(lrow),
        build_column(hcol) + str(hrow),
    ])


def parse_ranges(expr, max_col, max_row, headers=None, min_col=1, min_row=1):
    def make_return(lcol, lrow, hcol, hrow):
        return (
            range_expr(lcol, lrow, hcol, hrow),
            range_args(lcol, lrow, hcol, hrow),
        )

    def make_range1(matches):
        if matches.group(1):
            lkey = matches.group(1)
            lcol = headers.get(lkey[1:], 0) if headers else 0
            if lcol == 0:
                Ctx.throw('无法定位指定列：{0!r}', lkey)
            lrow = int(matches.group(2)[1:]) if matches.group(2) else min_row
        else:
            lcol = parse_column(matches.group(3)) \
                if matches.group(3) else min_col
            lrow = int(matches.group(4)) if matches.group(4) else min_row
        if matches.group(5):
            hkey = matches.group(5)
            hcol = headers.get(hkey[1:], 0) if headers else 0
            if hcol == 0:
                Ctx.throw('无法定位指定列：{0!r}', hkey)
            hrow = int(matches.group(6)[1:]) if matches.group(6) else max_row
        else:
            hcol = parse_column(matches.group(7)) \
                if matches.group(7) else max_col
            hrow = int(matches.group(8)) if matches.group(8) else max_row
        if lcol > hcol:
            lcol, hcol = hcol, lcol
        if lrow > hrow:
            lrow, hrow = hrow, lrow
        return make_return(lcol, lrow, hcol, hrow)

    def make_range2(matches):
        if matches.group(1):
            ikey = matches.group(1)
            icol = headers.get(ikey[1:], 0) if headers else 0
            if icol == 0:
                Ctx.throw('无法定位指定列：{0!r}', ikey)
        else:
            icol = parse_column(matches.group(2))
        return make_return(icol, min_row, icol, max_row)

    def make_range3(matches):
        irow = int(matches.group(1))
        return make_return(min_col, irow, max_col, irow)

    if isinstance(expr, str):
        if expr.strip() == RANGE_SEP:
            return make_return(min_col, min_row, max_col, max_row)

        for regex, make_range in [(RANGE1_REGEX, make_range1),
                                  (RANGE2_REGEX, make_range2),
                                  (RANGE3_REGEX, make_range3)]:
            matches = regex.match(expr.strip())
            if matches:
                return make_range(matches)

    Ctx.throw('错误的区域格式：{0!r}', expr)


def _fit_str_num(val1, val2):
    try:
        return float(val1.strip()), val2
    except ValueError:
        return val1.strip(), str(val2)


def _fit_str_bool(val1, val2):
    try:
        return float(val1.strip()), val2
    except ValueError:
        return val1.strip(), '1' if val2 else '0'


def _fit_str_datetime(val1, val2):
    try:
        fit1 = dateparser.parse(
            val1.strip(), fuzzy=False, ignoretz=True, default=DEFAULT_DATETIME)
    except (ValueError, OverflowError):
        return val1.strip(), str(val2)
    else:
        if isinstance(val2, datetime.date):
            return fit1, datetime.datetime.combine(val2, datetime.time())
        else:
            return fit1, val2


def _fit_date_datetime(val1, val2):
    return datetime.datetime.combine(val1, datetime.time()), val2


def _fit_datetime_date(val1, val2):
    return val1, datetime.datetime.combine(val2, datetime.time())


def _fit_nothing2(val1, val2):
    return val1, val2


XEQ_ALL_TYPES = {
    int: 10,
    float: 11,
    bool: 12,
    datetime.datetime: 21,
    datetime.date: 22,
    str: 30,
    tuple: 40,
    type(None): 50,
}
XEQ_FIT_TYPES = {
    # cast types
    (str, int):                 _fit_str_num,
    (str, float):               _fit_str_num,
    (str, bool):                _fit_str_bool,
    (str, datetime.date):       _fit_str_datetime,
    (str, datetime.datetime):   _fit_str_datetime,
    (str, type(None)):          (lambda v1, _: (v1.strip(), '')),
    (type(None), str):          (lambda _, v2: ('', v2.strip())),
    (str, str):                 (lambda v1, v2: (v1.strip(), v2.strip())),

    # compatible types
    (int, int):                 _fit_nothing2,
    (int, float):               _fit_nothing2,
    (int, bool):                _fit_nothing2,
    (float, float):             _fit_nothing2,
    (float, int):               _fit_nothing2,
    (float, bool):              _fit_nothing2,
    (bool, int):                _fit_nothing2,
    (bool, float):              _fit_nothing2,
    (bool, bool):               _fit_nothing2,
    (type(None), type(None)):   _fit_nothing2,

    # datetime types
    (datetime.date, datetime.date):         _fit_nothing2,
    (datetime.date, datetime.datetime):     _fit_date_datetime,
    (datetime.datetime, datetime.date):     _fit_datetime_date,
    (datetime.datetime, datetime.datetime): _fit_nothing2,
}


def xeq_(val1, val2):
    tp_val1 = type(val1)
    tp_val2 = type(val2)
    if tp_val1 not in XEQ_ALL_TYPES \
            or tp_val2 not in XEQ_ALL_TYPES:
        Ctx.throw('不支持该类型之间的比较：{0} <-> {1}',
                  tp_val1.__name__, tp_val2.__name__)
    elif val1 is val2:
        return True
    elif tp_val1 is tuple and tp_val2 is tuple:
        if len(val1) != len(val2):
            return False
        return all(itertools.starmap(xeq_, zip(val1, val2)))
    elif (tp_val1, tp_val2) in XEQ_FIT_TYPES:
        fit1, fit2 = XEQ_FIT_TYPES[(tp_val1, tp_val2)](val1, val2)
        return fit1 == fit2
    elif (tp_val2, tp_val1) in XEQ_FIT_TYPES:
        fit2, fit1 = XEQ_FIT_TYPES[(tp_val2, tp_val1)](val2, val1)
        return fit1 == fit2
    else:
        return False


def xlt_(val1, val2):
    tp_val1 = type(val1)
    tp_val2 = type(val2)
    if tp_val1 not in XEQ_ALL_TYPES \
            or tp_val2 not in XEQ_ALL_TYPES:
        Ctx.throw('不支持该类型之间的比较：{0} <-> {1}',
                  tp_val1.__name__, tp_val2.__name__)
    elif val1 is val2:
        return False
    elif tp_val1 is tuple and tp_val2 is tuple:
        for elm1, elm2 in zip(val1, val2):
            if not xeq_(elm1, elm2):
                return xlt_(elm1, elm2)
        return len(val1) < len(val2)
    elif (tp_val1, tp_val2) in XEQ_FIT_TYPES:
        fit1, fit2 = XEQ_FIT_TYPES[(tp_val1, tp_val2)](val1, val2)
        return fit1 < fit2
    elif (tp_val2, tp_val1) in XEQ_FIT_TYPES:
        fit2, fit1 = XEQ_FIT_TYPES[(tp_val2, tp_val1)](val2, val1)
        return fit1 < fit2
    else:
        return XEQ_ALL_TYPES[tp_val1] < XEQ_ALL_TYPES[tp_val2]


def xcmp_(val1, val2):
    if xeq_(val1, val2):
        return 0
    elif xlt_(val1, val2):
        return -1
    else:
        return 1


class ArrayView(object):

    def __init__(self, sheet, array, offset, vindex):
        self._sheet = sheet
        self._array = array
        self._offset = offset
        self._vindex = vindex

    def __len__(self):
        return len(self._array)

    def __iter__(self):
        for elm in self._array:
            yield elm.value

    def __getitem__(self, key):
        if isinstance(key, slice):
            if key.step is not None:
                Ctx.throw('行对象不支持间隔切片：{0!r}', key)
            if key.start is None and key.stop is None:
                return self
            hidx1 = 1 if key.start is None else key.start
            hidx2 = len(self._array) + 1 if key.stop is None else key.stop
            hidx1 = self.hidx(hidx1) - self._offset
            hidx2 = self.hidx(hidx2) - self._offset
            return self.cut(min(hidx1, hidx2), abs(hidx2 - hidx1) + 1)
        return self.valx(key)

    @property
    def vidx(self):
        return self._vindex

    def hidx(self, key):
        if not isinstance(key, int):
            return self._sheet.hidx(key)
        else:
            if key <= 0:
                Ctx.throw('无法定位指定列：{0!r}', key)
            return key + self._offset

    def expr(self, key):
        return build_column(self.hidx(key)) + str(self._vindex)

    def keys(self, *idxs, token=True):
        def impl(idxs, offset, maxidx):
            for hidx in idxs:
                if not isinstance(hidx, int) or hidx <= 0:
                    Ctx.throw('索引必须是大于零的整数：{0!r}', hidx)
                if hidx > maxidx:
                    Ctx.throw('索引超出区域范围：{0!r}', hidx)
                yield offset + hidx
        idxs = tuple(impl(idxs, self._offset, len(self._array)))
        return self._sheet.keys(*idxs, token=token)

    def vals(self, *keys):
        return tuple(self.valx(key) for key in keys)

    def valx(self, key):
        col = self.hidx(key) - self._offset
        if 0 < col <= len(self._array):
            return self._array[col - 1].value
        return None

    def cut(self, key, size):
        for elm in self.slc(key, size, 1):
            return elm
        return None

    def slc(self, key, size, num):
        if not isinstance(size, int) or size <= 0:
            Ctx.throw('分组大小必须是大于零的整数：{0!r}', size)
        if not isinstance(num, int) or num < 0:
            Ctx.throw('分组数量必须是不为负的整数：{0!r}', num)
        col = self.hidx(key) - self._offset
        if 0 < col <= num * size + col - 1 <= len(self._array):
            for idx in range(num):
                offset = idx * size + col - 1
                yield type(self)(
                    self._sheet,
                    self._array[offset:offset+size],
                    self._offset + offset,
                    self._vindex,
                )
        else:
            Ctx.throw('分组超出区域范围：{0!r},{1},{2}', key, size, num)

    def aslist(self):
        return [elm.value for elm in self._array]

    def asdict(self, *keys):
        if not keys:
            idx0 = self._offset + 1
            idxs = tuple(range(idx0, idx0 + len(self._array)))
            keys = self._sheet.keys(*idxs, token=False)
        return {str(key): elm.value for key, elm in zip(keys, self._array)}


class SheetView(object):

    def __init__(self, filepath, sheetname, worksheet,
                 headers=None, beg_col=None, beg_row=None, end_col=None, end_row=None):
        self._filepath = filepath
        self._filename = os.path.basename(filepath)
        self._sheetname = sheetname
        self._worksheet = worksheet
        self._headers = headers or {}
        self._beg_col = beg_col or 1
        self._beg_row = beg_row or 1
        self._end_col = end_col or self._worksheet.max_column
        self._end_row = end_row or self._worksheet.max_row
        self._cur_row = None

    def __str__(self):
        return '{0}#{1}'.format(self._filename, self._sheetname)

    def __iter__(self):
        return self[RANGE_SEP]

    def __getitem__(self, expr):
        if isinstance(expr, slice):
            if expr.step is not None:
                Ctx.throw('工作表不支持间隔切片：{0!r}', expr)
            expr = RANGE_SEP.join([
                str(expr.start or ''),
                str(expr.stop or ''),
            ])
        if expr.strip() == RANGE_NIL:
            return
        max_row = self._worksheet.max_row
        max_col = self._worksheet.max_column
        if max_row <= 0 or max_col <= 0:
            return
        slc_expr, args = parse_ranges(
            expr, max_col, max_row, self._headers, min_col=self._beg_col)
        for idx, row in xslice(self._worksheet[slc_expr], self._beg_row, self._end_row):
            Ctx.set_ctx(str(self), args.voff + idx)
            self._cur_row = ArrayView(self, row, args.hoff, args.voff + idx)
            yield self._cur_row

    @property
    def vobj(self, vidx=None):
        if vidx is None:
            return self._cur_row
        if self._worksheet.max_row <= 0 \
                or self._worksheet.max_column <= 0:
            return None
        if isinstance(vidx, int):
            if vidx <= 0:
                Ctx.throw('无法定位指定行：{0!r}', vidx)
        else:
            if not isinstance(vidx, str) or \
                    not VINDEX_REGEX.match(vidx.strip()):
                Ctx.throw('无法定位指定行：{0!r}', vidx)
            vidx = int(vidx.strip())
        return ArrayView(self, self._worksheet[str(vidx)], 0, vidx)

    def hidx(self, key):
        if isinstance(key, str):
            if key.startswith(COLKEY_TOKEN):
                col = self._headers.get(key[1:], 0)
            else:
                col = parse_column(key)
        elif isinstance(key, int):
            col = key if key > 0 else 0
        else:
            col = 0
        if col == 0:
            Ctx.throw('无法定位指定列：{0!r}', key)
        return col

    def keys(self, *idxs, token=True):
        def impl(idxs, headers):
            for hidx in idxs:
                if not isinstance(hidx, int) or hidx <= 0:
                    Ctx.throw('索引必须是大于零的整数：{0!r}', hidx)
                if hidx not in headers:
                    Ctx.throw('索引指向的列不包含表头：{0!r}', hidx)
                if token:
                    yield COLKEY_TOKEN + headers[hidx]
                else:
                    yield headers[hidx]
        return tuple(impl(idxs, self._headers))

    def valx(self, expr):
        if not isinstance(expr, str) \
                or not CELLXX_REGEX.match(expr.strip()):
            Ctx.throw('错误的单元格式：{0!r}', expr)
        return self._worksheet[expr.strip().upper()].value

    def rehead(self, vidx, hbeg=None, vbeg=None, hend=None, vend=None):
        max_row = self._worksheet.max_row
        max_col = self._worksheet.max_column
        if max_row <= 0 or max_col <= 0:
            return self
        if isinstance(vidx, int):
            if vidx <= 0:
                Ctx.throw('无法定位指定行：{0!r}', vidx)
        else:
            if not isinstance(vidx, str) or \
                    not VINDEX_REGEX.match(vidx.strip()):
                Ctx.throw('无法定位指定行：{0!r}', vidx)
            vidx = int(vidx.strip())
        slc_expr, _ = parse_ranges(
            str(vidx), max_col, max_row, min_col=self._beg_col)
        for head_row in self._worksheet[slc_expr]:
            headers = {
                str(cell.value).strip(): col
                for col, cell in enumerate(head_row, 1)
                if cell.value not in {None, ''}
            }
            headers.update((val, key) for key, val in list(headers.items()))
            return type(self)(
                self._filepath,
                self._sheetname,
                self._worksheet,
                headers=headers,
                beg_col=hbeg,
                beg_row=vbeg,
                end_col=hend,
                end_row=vend,
            )
        return self

    def select(self, expr):
        max_row = self._worksheet.max_row
        max_col = self._worksheet.max_column
        if max_row <= 0 or max_col <= 0:
            return
        slc_expr, args = parse_ranges(
            expr, max_col, max_row, self._headers, min_col=self._beg_col)
        if slc_expr is not None:
            for idx, row in xslice(self._worksheet[slc_expr], self._beg_row, self._end_row):
                yield ArrayView(self, row, args.hoff, args.voff + idx)

    def locate(self, ltag, htag, loff=1, hoff=-1):
        hbeg, vbeg, hend, vend = None, None, None, None
        for row in self.select(RANGE_SEP):
            for idx, val in enumerate(row, 1):
                if val is None:
                    continue
                if vbeg is None and xeq_(ltag, val):
                    vbeg = row.vidx
                    hbeg = row.hidx(idx)
                    continue
                if vend is None and xeq_(htag, val):
                    vend = row.vidx
                    hend = row.hidx(idx)
                    break
        if vbeg is None:
            Ctx.error('找不匹配的起始标签：{0}', ltag)
            return self[RANGE_NIL]
        if vend is None:
            Ctx.error('找不匹配的结束标签：{0}', htag)
            return self[RANGE_NIL]
        return self.rehead(
            vbeg,
            hbeg=hbeg + 1,
            vbeg=vbeg + loff,
            hend=hend + 0,
            vend=vend + hoff)

    def findone(self, val1, tab):
        for row in self.findall(val1, tab):
            return row
        Ctx.throw('指定区域\'{0}\'!{1}找不到对应值{2!r}', str(self), tab, val1)

    def findall(self, val1, tab):
        if isinstance(val1, tuple):
            if len(val1) <= 0:
                Ctx.throw('目标元素个数必须大于零：{0!r}', val1)
            keys = tuple(range(1, len(val1) + 1))
            valx = (lambda row: row.vals(*keys))
        else:
            valx = (lambda row: row.valx(1))
        for row in self.select(tab):
            if xeq_(val1, valx(row)):
                yield row

    def vlookup(self, val1, tab, idx):
        if isinstance(val1, tuple):
            if len(val1) <= 0:
                Ctx.throw('目标元素个数必须大于零：{0!r}', val1)
            keys = tuple(range(1, len(val1) + 1))
            valx = (lambda row: row.vals(*keys))
        else:
            valx = (lambda row: row.valx(1))
        for row in self.select(tab):
            if xeq_(val1, valx(row)):
                return row.valx(idx)
        Ctx.error('指定区域\'{0}\'!{1}找不到对应值{2!r}', str(self), tab, val1)
        return None

    def chkuniq(self, tab):
        max_row = self._worksheet.max_row
        max_col = self._worksheet.max_column
        if max_row <= 0 or max_col <= 0:
            return
        slc_expr, args = parse_ranges(
            tab, max_col, max_row, self._headers, min_col=self._beg_col)
        keys = tuple(range(1, args.hnum + 1))
        iterable = (
            ArrayView(self, row, args.hoff, args.voff + idx)
            for idx, row in xslice(self._worksheet[slc_expr], self._beg_row, self._end_row)
        )
        for key, group in xgroupby(iterable, *keys):
            first = None
            for row in group:
                if first is None:
                    first = row
                    continue
                Ctx.error('指定区域\'{0}\'!{1}含有重复值{2!r}：{3} <-> {4}',
                          str(self), tab, key, row.vidx, first.vidx)


def xslice(rows, vbeg=1, vend=sys.maxsize):
    for idx, row in enumerate(rows, 1):
        if idx > vend:
            break
        elif idx < vbeg:
            continue
        yield idx, row


def xrequire(rows, *keys):
    for row in rows:
        if all(not xeq_(row.valx(key), '') for key in keys):
            yield row


def xpickcol(value, to_int=False):
    if not isinstance(value, str) \
            or not CELLXX_REGEX.match(value.strip()):
        Ctx.throw('错误的单元格式：{0!r}', value)
    expr = CELLXX_REGEX.match(value.strip()).group(1)
    return parse_column(expr) if to_int else expr


def xpickrow(value):
    if not isinstance(value, str) \
            or not CELLXX_REGEX.match(value.strip()):
        Ctx.throw('错误的单元格式：{0!r}', value)
    return int(CELLXX_REGEX.match(value.strip()).group(2))


def xoffset(value, hoff=1, voff=0):
    if not isinstance(value, str) \
            or not CELLXX_REGEX.match(value.strip()):
        Ctx.throw('错误的单元格式：{0!r}', value)
    if not isinstance(hoff, int):
        Ctx.throw('水平位移量必须是整数：{0!r}', hoff)
    if not isinstance(voff, int):
        Ctx.throw('垂直位移量必须是整数：{0!r}', voff)
    if hoff == 0 and voff == 0:
        return value
    hpar, vpar = CELLXX_REGEX.match(value.strip()).groups()
    hidx, vidx = parse_column(hpar), int(vpar)
    if hidx + hoff <= 0 or vidx + voff <= 0:
        Ctx.throw('单元偏移超出范围：{0!r},{1},{2}', value, hoff, voff)
    return build_column(hidx + hoff) + str(vidx + voff)


def xgroupby(rows, *keys):
    def getkey(row):
        return tuple(row.valx(key) for key in keys)

    def rowcmp(row1, row2):
        key1 = getkey(row1)
        key2 = getkey(row2)
        if xeq_(key1, key2):
            return 0
        elif xlt_(key1, key2):
            return -1
        else:
            return 1

    origin_rows = list(rows)
    sorted_rows = sorted(origin_rows, key=functools.cmp_to_key(rowcmp))
    for key, group in itertools.groupby(sorted_rows, key=functools.cmp_to_key(rowcmp)):
        yield getkey(key.obj), group


def load_worksheet(filepath, sheetname, head=0):
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb[sheetname]
    except IOError:
        Ctx.abort('无法打开目标工作簿：{0}', filepath)
    except KeyError:
        Ctx.abort('无法打开目标工作表：{0}#{1}', filepath, sheetname)

    sheet_view = SheetView(filepath, sheetname, ws)
    if 0 < head <= ws.max_row:
        sheet_view = sheet_view.rehead(head)
    return sheet_view
