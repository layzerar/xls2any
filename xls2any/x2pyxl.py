# -*- coding: utf-8 -*-

import re
import os
import functools
import itertools

import openpyxl

from . import utils

Ctx = utils.Ctx


RANGE_SEP = ':'
COLUMN_NAME_MARK = '@'
COLUMN_REGEX = \
    re.compile(r'^([A-Z]+)$')
RANGE1_REGEX = \
    re.compile(r'^([A-Z]+)?:([A-Z]+)?$')
RANGE2_REGEX = \
    re.compile(r'^([1-9][0-9]*)?:([1-9][0-9]*)?$')
RANGE3_REGEX = \
    re.compile(r'^(([A-Z]+)([1-9][0-9]*))?:(([A-Z]+)([1-9][0-9]*))?$')
RANGE4_REGEX = COLUMN_REGEX
RANGE5_REGEX = \
    re.compile(r'^([1-9][0-9]*)$')


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
    raise ValueError('错误的列号格式：{0!r}'.format(expr))


def parse_ranges(expr, max_col, max_row, min_col=1, min_row=1):
    def make_range1(matches):
        lcol = parse_column(matches.group(1)) \
            if matches.group(1) else min_col
        hcol = parse_column(matches.group(2)) \
            if matches.group(2) else max_col
        return RANGE_SEP.join([
            build_column(lcol) + str(min_row),
            build_column(hcol) + str(max_row),
        ]), lcol - 1, min_row - 1

    def make_range2(matches):
        lrow = int(matches.group(1)) if matches.group(1) else min_row
        hrow = int(matches.group(2)) if matches.group(2) else max_row
        return RANGE_SEP.join([
            build_column(min_col) + str(lrow),
            build_column(max_col) + str(hrow),
        ]), min_col - 1, lrow - 1

    def make_range3(matches):
        if matches.group(1):
            lcol = parse_column(matches.group(2))
            lrow = int(matches.group(3))
        else:
            lcol = min_col
            lrow = min_row
        if matches.group(4):
            hcol = parse_column(matches.group(5))
            hrow = int(matches.group(6))
        else:
            hcol = max_col
            hrow = max_row
        if lcol > hcol:
            lcol, hcol = hcol, lcol
        if lrow > hrow:
            lrow, hrow = hrow, lrow
        return RANGE_SEP.join([
            build_column(lcol) + str(lrow),
            build_column(hcol) + str(hrow),
        ]), lcol - 1, lrow - 1

    def make_range4(matches):
        icol = parse_column(matches.group(1))
        return RANGE_SEP.join([
            build_column(icol) + str(min_row),
            build_column(icol) + str(max_row),
        ]), icol - 1, min_row - 1

    def make_range5(matches):
        irow = int(matches.group(1))
        return RANGE_SEP.join([
            build_column(min_col) + str(irow),
            build_column(max_col) + str(irow),
        ]), min_col - 1, irow - 1

    if isinstance(expr, str):
        if expr.strip() == RANGE_SEP:
            return RANGE_SEP.join([
                build_column(min_col) + str(min_row),
                build_column(max_col) + str(max_row),
            ]), min_col - 1, min_row - 1

        for regex, make_range in [(RANGE1_REGEX, make_range1),
                                  (RANGE2_REGEX, make_range2),
                                  (RANGE3_REGEX, make_range3),
                                  (RANGE4_REGEX, make_range4),
                                  (RANGE5_REGEX, make_range5)]:
            matches = regex.match(expr.strip().upper())
            if matches:
                return make_range(matches)

    raise ValueError('错误的区域格式：{0!r}'.format(expr))


def xeq_(val1, val2):
    tp_val1 = type(val1)
    tp_val2 = type(val2)
    try:
        if val1 is val2:
            return True
        elif tp_val1 is tp_val2:
            if tp_val1 is tuple:
                if len(val1) == len(val2):
                    return all(itertools.starmap(xeq_, zip(val1, val2)))
                else:
                    return False
            else:
                return val1 == val2
        elif tp_val1 is tuple or tp_val2 is tuple:
            return False
        elif tp_val1 is float:
            if tp_val2 is int:
                return val1 == val2
            elif tp_val2 is str:
                return val1 == float(val2.strip())
            else:
                return False
        elif tp_val2 is float:
            if tp_val1 is int:
                return val2 == val1
            elif tp_val1 is str:
                return val2 == float(val1.strip())
            else:
                return False
        elif tp_val1 is str:
            return val1.strip() == ('' if val2 is None else str(val2))
        elif tp_val2 is str:
            return val2.strip() == ('' if val1 is None else str(val1))
        else:
            return str(val1) == str(val2)
    except ValueError:
        return False


def xlt_(val1, val2):
    return False if xeq_(val1, val2) else id(val1) < id(val2)


def xcmp_(val1, val2):
    if xeq_(val1, val2):
        return 0
    elif id(val1) < id(val2):
        return -1
    else:
        return 1


class ArrayView(object):

    def __init__(self, sheet, array, offset):
        self._sheet = sheet
        self._array = array
        self._offset = offset

    def __len__(self):
        return len(self._array)

    def __iter__(self):
        for elm in self._array:
            yield elm.value

    def val(self, key):
        if not isinstance(key, int):
            col = self._sheet.column(key) - self._offset
        else:
            col = key
        if 0 < col <= len(self._array):
            return self._array[col - 1].value
        return None

    def cut(self, key, size):
        for elm in self.slc(key, size, 1):
            return elm
        return None

    def slc(self, key, size, num):
        if not isinstance(size, int) or size <= 0:
            raise ValueError('分组大小必须是大于零的整数：{0!r}'.format(size))
        if not isinstance(num, int) or num < 0:
            raise ValueError('分组数量必须是不为负的整数：{0!r}'.format(num))
        if not isinstance(key, int):
            col = self._sheet.column(key) - self._offset
        else:
            col = key
        if 0 < col <= num * size + col - 1 <= len(self._array):
            for idx in range(num):
                offset = idx * size + col - 1
                yield type(self)(
                    self._sheet,
                    self._array[offset:offset+size],
                    self._offset + offset,
                )
        else:
            raise ValueError('分组超出区域范围：{0!r},{1},{2}'.format(
                build_column(col), size, num))

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
        self._active = None

    def __str__(self):
        return '{0}#{1}'.format(self._filename, self._sheetname)

    def __iter__(self):
        return self[RANGE_SEP]

    def __getitem__(self, expr):
        if isinstance(expr, slice):
            raise ValueError('工作表遍历暂不支持切片：{0!r}'.format(expr))
        max_row = self._worksheet.max_row
        max_col = self._worksheet.max_column
        if max_row <= 0 or max_col <= 0:
            return
        slc, off_col, off_row = parse_ranges(expr, max_col, max_row)
        for idx, row in enumerate(self._worksheet[slc], 1):
            Ctx.set_ctx(str(self), off_row + idx)
            self._active = ArrayView(self, row, off_col)
            yield self._active

    @property
    def active(self):
        return self._active

    def column(self, key):
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
            raise ValueError('无法根据列名找到指定列：{0!r}'.format(key))
        return col

    def select(self, expr):
        max_row = self._worksheet.max_row
        max_col = self._worksheet.max_column
        if max_row > 0 and max_col > 0:
            slc, off_col, _ = parse_ranges(expr, max_col, max_row)
            if slc is not None:
                for row in self._worksheet[slc]:
                    yield ArrayView(self, row, off_col)

    def vlookup(self, val, tab, idx):
        for row in self.select(tab):
            if row.val(1) is not None \
                    and xeq_(val, row.val(1)):
                return row.val(idx)
        Ctx.error('指定区域{0}${1}找不到对应值{2!r}', str(self), tab, val)
        return None


def xgroupby(rows, *keys):
    def getkey(row):
        return tuple(row.val(key) for key in keys)

    def rowcmp(row1, row2):
        if xeq_(getkey(row1), getkey(row2)):
            return 0
        elif origin_rows.index(row1) < origin_rows.index(row2):
            return -1
        else:
            return 1

    origin_rows = list(rows)
    sorted_rows = sorted(origin_rows, key=functools.cmp_to_key(rowcmp))
    for key, group in itertools.groupby(sorted_rows, key=functools.cmp_to_key(rowcmp)):
        yield key.obj, group


def get_worksheet_headers(ws, head):
    expr, _, _ = parse_ranges(str(head), ws.max_column, ws.max_row)
    rows = ws[expr]
    head_row = rows[0] if isinstance(rows, tuple) else next(rows)
    return {
        COLUMN_NAME_MARK + str(cell.value).strip(): col
        for col, cell in enumerate(head_row, 1)
        if cell.value not in {None, ''}
    }


def load_worksheet(filepath, sheetname, head=1):
    try:
        wb = openpyxl.load_workbook(filepath)
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
