# -*- coding: utf-8 -*-

import io
import re
import datetime


LUA_IDENT_REGEX = \
    re.compile(r'^([a-zA-Z_][a-zA-Z0-9_]*)$')
LUA_KEYWORDS_REGEX = \
    re.compile(r'^(and|break|do|else|elseif|end|false|for|function|goto|if|in|local|nil|not|or|repeat|return|then|true|until|while)$')


def is_lua_ident(key):
    if not isinstance(key, str):
        return False
    if not LUA_IDENT_REGEX.match(key):
        return False
    return not LUA_KEYWORDS_REGEX.match(key)


def get_lua_escape_table():
    table = {}
    for i in range(32):
        table[i] = r'\x' + bytes([i]).hex()
    table[ord('\0')] = r'\0'
    table[ord('\a')] = r'\a'
    table[ord('\b')] = r'\b'
    table[ord('\f')] = r'\f'
    table[ord('\n')] = r'\n'
    table[ord('\r')] = r'\r'
    table[ord('\t')] = r'\t'
    table[ord('\v')] = r'\v'
    table[ord('\\')] = r'\\'
    table[ord('\"')] = r'\"'
    table[ord('\'')] = r"\'"
    table[ord('\x7f')] = r'\x7f'
    return table


def date_tolua(obj):
    return (
        '{{year = {0.year}, month = {0.month}, day = {0.day}}}'
    ).format(obj)


def datetime_tolua(obj):
    return (
        '{{year = {0.year}, month = {0.month}, day = {0.day}'
        ', hour = {0.hour}, min = {0.minute}, sec = {0.second}}}'
    ).format(obj)


def _check_ref(obj, refs):
    if type(obj) in (list, tuple, dict):
        if id(obj) in refs:
            return True
        refs.add(id(obj))
    return False


class LuaEncoder(object):
    LUA_ESCAPE_TABLE = get_lua_escape_table()

    def __init__(self, check_circular=True, indent=None, separators=(', ', ' = ')):
        self._indent = ' ' * (indent or 0)
        self._separators = separators
        self._check_circular = check_circular

    @classmethod
    def escape(cls, obj):
        if not isinstance(obj, str):
            obj = str(obj)
        return obj.translate(cls.LUA_ESCAPE_TABLE)

    def encode(self, obj):
        buf = io.StringIO()
        self._encode(obj, buf, set())
        return buf.getvalue()

    def _encode(self, obj, buf, refs, depth=0):
        indent = self._indent
        sep1 = self._separators[0]  # `,`
        sep2 = self._separators[1]  # `=`
        check_circular = self._check_circular
        if indent:
            newline = '\n'
            sep1 = sep1.rstrip() + newline
        else:
            newline = ''

        tp = type(obj)
        if tp is str:
            buf.write('"')
            buf.write(self.escape(obj))
            buf.write('"')
        elif tp in (int, float, complex):
            buf.write(str(obj))
        elif obj is None:
            buf.write('nil')
        elif obj is True:
            buf.write('true')
        elif obj is False:
            buf.write('false')
        elif tp in (list, tuple, dict):
            _check_ref(obj, refs)

            #buf.write(indent * depth)
            if len(obj) == 0:
                buf.write('{}')
                return

            buf.write('{')
            buf.write(newline)
            depth += 1

            nlast = 0
            if tp is dict:
                for k, v in obj.items():
                    if check_circular:
                        if _check_ref(k, refs) or _check_ref(v, refs):
                            continue
                    buf.write(indent * depth)
                    if is_lua_ident(k):
                        buf.write(k)
                        
                    else:
                        buf.write('[')
                        self._encode(k, buf, refs, depth)
                        buf.write(']')
                    buf.write(sep2)
                    self._encode(v, buf, refs, depth)
                    nlast = buf.write(sep1)
            else:
                for e in obj:
                    if check_circular:
                        if _check_ref(e, refs):
                            continue
                    buf.write(indent * depth)
                    self._encode(e, buf, refs, depth)
                    nlast = buf.write(sep1)
            if not indent and nlast > 0:
                buf.seek(buf.tell() - nlast)
                buf.truncate()

            depth -= 1
            buf.write(indent * depth)
            buf.write('}')
        elif tp is datetime.datetime:
            buf.write(datetime_tolua(obj))
        elif tp is datetime.date:
            buf.write(date_tolua(obj))
        else:
            raise TypeError('不能将{}转换为Lua类型'.format(tp.__name__))


def dumps(obj, **kwds):
    return LuaEncoder(**kwds).encode(obj)
