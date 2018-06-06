# -*- coding: utf-8 -*-

import re
import os
import datetime
import functools
import itertools
import collections

import openpyxl
import dateutil.parser as dateparser

from . import utils

Ctx = utils.Ctx


RANGE_SEP = ':'
COLUMN_NAME_MARK = '@'
CELLXX_REGEX = \
    re.compile(r'^([A-Z]+)([1-9][0-9]*)$')
COLUMN_REGEX = \
    re.compile(r'^([A-Z]+)$')
RANGE1_REGEX = \
    re.compile(r'^([A-Z]+)?:([A-Z]+)?$')
RANGE2_REGEX = \
    re.compile(r'^([1-9][0-9]*)?:([1-9][0-9]*)?$')
RANGE3_REGEX = \
    re.compile(r'^([A-Z]+)?([1-9][0-9]*)?:([A-Z]+)?([1-9][0-9]*)?$')
RANGE4_REGEX = COLUMN_REGEX
RANGE5_REGEX = \
    re.compile(r'^([1-9][0-9]*)$')

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


def parse_ranges(expr, max_col, max_row, min_col=1, min_row=1):
    def make_return(lcol, lrow, hcol, hrow):
        return (
            range_expr(lcol, lrow, hcol, hrow),
            range_args(lcol, lrow, hcol, hrow),
        )

    def make_range1(matches):
        lcol = parse_column(matches.group(1)) \
            if matches.group(1) else min_col
        hcol = parse_column(matches.group(2)) \
            if matches.group(2) else max_col
        return make_return(lcol, min_row, hcol, max_row)

    def make_range2(matches):
        lrow = int(matches.group(1)) if matches.group(1) else min_row
        hrow = int(matches.group(2)) if matches.group(2) else max_row
        return make_return(min_col, lrow, max_col, hrow)

    def make_range3(matches):
        lcol = parse_column(matches.group(1)) if matches.group(1) else min_col
        lrow = int(matches.group(2)) if matches.group(2) else min_row
        hcol = parse_column(matches.group(3)) if matches.group(3) else max_col
        hrow = int(matches.group(4)) if matches.group(4) else max_row
        if lcol > hcol:
            lcol, hcol = hcol, lcol
        if lrow > hrow:
            lrow, hrow = hrow, lrow
        return make_return(lcol, lrow, hcol, hrow)

    def make_range4(matches):
        icol = parse_column(matches.group(1))
        return make_return(icol, min_row, icol, max_row)

    def make_range5(matches):
        irow = int(matches.group(1))
        return make_return(min_col, irow, max_col, irow)

    if isinstance(expr, str):
        if expr.strip() == RANGE_SEP:
            return make_return(min_col, min_row, max_col, max_row)

        for regex, make_range in [(RANGE1_REGEX, make_range1),
                                  (RANGE2_REGEX, make_range2),
                                  (RANGE3_REGEX, make_range3),
                                  (RANGE4_REGEX, make_range4),
                                  (RANGE5_REGEX, make_range5)]:
            matches = regex.match(expr.strip().upper())
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

    def val(self, key):
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
        return {key: elm.value for key, elm in zip(keys, self._array)}


class SheetView(object):

    def __init__(self, filepath, sheetname, worksheet, headers=None):
        self._filepath = filepath
        self._filename = os.path.basename(filepath)
        self._sheetname = sheetname
        self._worksheet = worksheet
        self._headers = headers or {}
        self._cur_row = None

    def __str__(self):
        return '{0}#{1}'.format(self._filename, self._sheetname)

    def __iter__(self):
        return self[RANGE_SEP]

    def __getitem__(self, expr):
        if isinstance(expr, slice):
            Ctx.throw('工作表遍历不支持切片：{0!r}', expr)
        max_row = self._worksheet.max_row
        max_col = self._worksheet.max_column
        if max_row <= 0 or max_col <= 0:
            return
        slc_expr, args = parse_ranges(expr, max_col, max_row)
        for idx, row in enumerate(self._worksheet[slc_expr], 1):
            Ctx.set_ctx(str(self), args.voff + idx)
            self._cur_row = ArrayView(self, row, args.hoff, args.voff + idx)
            yield self._cur_row

    @property
    def cur(self):
        return self._cur_row

    def val(self, expr, key=''):
        if isinstance(expr, int):
            if expr <= 0:
                Ctx.throw('行号必须是大于零的整数：{0!r}', expr)
            row, col = expr, self.hidx(key)
            expr = build_column(col) + str(row)
        elif not isinstance(expr, str) \
                or not CELLXX_REGEX.match(expr):
            Ctx.throw('错误的单元格式：{0!r}', expr)
        return self._worksheet[expr].value

    def hidx(self, key):
        if isinstance(key, str):
            if key.startswith(COLUMN_NAME_MARK):
                col = self._headers.get(key, 0)
            else:
                col = parse_column(key)
        elif isinstance(key, int):
            col = key if key > 0 else 0
        else:
            col = 0
        if col == 0:
            Ctx.throw('无法定位指定列：{0!r}', key)
        return col

    def select(self, expr):
        max_row = self._worksheet.max_row
        max_col = self._worksheet.max_column
        if max_row <= 0 or max_col <= 0:
            return
        slc_expr, args = parse_ranges(expr, max_col, max_row)
        if slc_expr is not None:
            for idx, row in enumerate(self._worksheet[slc_expr], 1):
                yield ArrayView(self, row, args.hoff, args.voff + idx)

    def search(self, val1, tab):
        for row in self.select(tab):
            for idx, val2 in enumerate(row, 1):
                if val2 is not None and xeq_(val1, val2):
                    return build_column(row.hidx(idx)) + str(row.vidx)
        return None

    def vlookup(self, val1, tab, idx):
        for row in self.select(tab):
            if row.val(1) is not None and xeq_(val1, row.val(1)):
                return row.val(idx)
        Ctx.error('指定区域\'{0}\'!{1}找不到对应值{2!r}', str(self), tab, val1)
        return None

    def chkuniq(self, tab):
        max_row = self._worksheet.max_row
        max_col = self._worksheet.max_column
        if max_row <= 0 or max_col <= 0:
            return
        slc_expr, args = parse_ranges(tab, max_col, max_row)
        keys = tuple(range(1, args.hnum + 1))
        iterable = (
            ArrayView(self, row, args.hoff, args.voff + idx)
            for idx, row in enumerate(self._worksheet[slc_expr], 1)
        )
        for key, group in xgroupby(iterable, *keys):
            first = None
            for row in group:
                if first is None:
                    first = row
                    continue
                Ctx.error('指定区域\'{0}\'!{1}含有重复值{2!r}：{3} <-> {4}',
                          str(self), tab, key, row.vidx, first.vidx)


def xrequire(rows, *keys):
    for row in rows:
        if all(not xeq_(row.val(key), '') for key in keys):
            yield row


def xpickcol(value, to_int=False):
    if not isinstance(value, str) \
            or not CELLXX_REGEX.match(value):
        Ctx.throw('错误的单元格式：{0!r}', value)
    expr = CELLXX_REGEX.match(value).group(1)
    if not to_int:
        return expr
    return parse_column(expr)


def xpickrow(value):
    if not isinstance(value, str) \
            or not CELLXX_REGEX.match(value):
        Ctx.throw('错误的单元格式：{0!r}', value)
    return int(CELLXX_REGEX.match(value).group(2))


def xgroupby(rows, *keys):
    def getkey(row):
        return tuple(row.val(key) for key in keys)

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


def get_worksheet_headers(ws, head):
    slc_expr, _ = parse_ranges(str(head), ws.max_column, ws.max_row)
    rows = ws[slc_expr]
    head_row = rows[0] if isinstance(rows, tuple) else next(rows)
    return {
        COLUMN_NAME_MARK + str(cell.value).strip(): col
        for col, cell in enumerate(head_row, 1)
        if cell.value not in {None, ''}
    }


def load_worksheet(filepath, sheetname, head=0):
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb[sheetname]
    except IOError:
        Ctx.abort('无法打开目标工作簿：{0}', filepath)
    except KeyError:
        Ctx.abort('无法打开目标工作表：{0}#{1}', filepath, sheetname)

    if 0 < head <= ws.max_row:
        headers = get_worksheet_headers(ws, head)
    else:
        headers = None
    return SheetView(filepath, sheetname, ws, headers=headers)
