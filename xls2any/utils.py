# -*- coding: utf-8 -*-

import sys
import datetime

import chardet
import colorama

colorama.init()


class Ctx(object):
    _ctx_arg = {}
    _log_fmt = '%(color)s%(asctime)s %(level)s [%(context)s] %(message)s%(reset)s'
    _at_abort = None

    @classmethod
    def set_ctx(cls, filename, lineno):
        cls._ctx_arg.update(
            filename=filename,
            lineno=lineno,
        )

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
    def error(cls, msg, *args, **kwds):
        line = cls.get_msg(
            colorama.Fore.CYAN,
            'ERROR',
            msg.format(*args, **kwds),
        )
        print(line, file=sys.stderr)

    @classmethod
    def abort(cls, msg, *args, **kwds):
        line = cls.get_msg(
            colorama.Fore.RED,
            'FATAL',
            msg.format(*args, **kwds),
        )
        print(line, file=sys.stderr)
        at_abort = cls._at_abort or (lambda: sys.exit(1))
        at_abort()

    @classmethod
    def set_abort_handler(cls, func):
        cls._at_abort = func


def detect_encoding(binary_data, default='utf-8'):
    result = chardet.detect(binary_data)
    return result['encoding'] or default


def open_as_stdout(filename, encoding='utf-8'):
    try:
        sys.stdout = open(filename, 'w', encoding=encoding)
    except IOError:
        Ctx.error('无法打开目标输出流：{0}', filename)
    except LookupError:
        Ctx.error('错误的文件编码名称：{0}', encoding)
