# -*- coding: utf-8 -*-

import sys
import datetime
import functools

import chardet
import colorama

colorama.init()


class Ctx(object):
    _ctx_arg = {}
    _log_fmt = '%(color)s%(asctime)s %(level)s [%(context)s] %(message)s%(reset)s'
    _at_abort = None
    _in_debug = False

    @classmethod
    def set_ctx(cls, filename, lineno):
        cls._ctx_arg.update(
            filename=filename,
            lineno=lineno,
        )

    @classmethod
    def set_debug(cls, flag):
        cls._in_debug = bool(flag)

    @classmethod
    def get_msg(cls, color, level, message, reset=colorama.Style.RESET_ALL):
        asctime = datetime.datetime.now().isoformat(timespec='seconds')
        if 'filename' in cls._ctx_arg:
            context = "%s:%s" % (
                cls._ctx_arg['filename'],
                cls._ctx_arg['lineno'],
            )
        else:
            context = '-'
        return cls._log_fmt % locals()

    @classmethod
    def debug(cls, msg, *args, **kwds):
        if not cls._in_debug:
            return
        line = cls.get_msg(
            colorama.Fore.LIGHTGREEN_EX,
            'DEBUG',
            str(msg).format(*args, **kwds),
        )
        print(line, file=sys.stderr)

    @classmethod
    def error(cls, msg, *args, **kwds):
        line = cls.get_msg(
            colorama.Fore.LIGHTYELLOW_EX,
            'ERROR',
            str(msg).format(*args, **kwds),
        )
        print(line, file=sys.stderr)

    @classmethod
    def throw(cls, msg, *args, exc_tp=ValueError, **kwds):
        raise exc_tp(str(msg).format(*args, **kwds))

    @classmethod
    def abort(cls, msg, *args, **kwds):
        line = cls.get_msg(
            colorama.Fore.LIGHTRED_EX,
            'FATAL',
            str(msg).format(*args, **kwds),
        )
        print(line, file=sys.stderr)
        at_abort = cls._at_abort or (lambda: sys.exit(1))
        at_abort()

    @classmethod
    def set_abort_handler(cls, func):
        cls._at_abort = func


def detect_encoding(binary_data, default='utf-8', confidence=0.75):
    result = chardet.detect(binary_data)
    if result['encoding'] and result['confidence'] >= confidence:
        return result['encoding']
    else:
        return default


def open_as_stdout(filename, encoding='utf-8'):
    try:
        sys.stdout = open(filename, 'w', encoding=encoding)
    except IOError:
        Ctx.throw('无法打开目标输出流：{0}', filename)
    except LookupError:
        Ctx.throw('错误的文件编码名称：{0}', encoding)


def cachedmethod(namespace, keyfmt, dumps=None, loads=None):
    def decorator(func):
        @functools.wraps(func)
        def decorate(*args, **kwds):
            cache_key  = keyfmt.format(*args, **kwds)
            cache_dict = getattr(args[0], namespace)
            if cache_key not in cache_dict:
                result = func(*args, **kwds)
                cache_val = result if dumps is None else dumps(result)
                cache_dict[cache_key] = cache_val
            else:
                cache_val = cache_dict[cache_key]
            return cache_val if loads is None else loads(cache_val)
        return decorate
    return decorator
