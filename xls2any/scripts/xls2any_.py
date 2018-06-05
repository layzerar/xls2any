# -*- coding: utf-8 -*-

import io
import os
import re
import sys
import json
import functools
import traceback

import click
import jinja2
from jinja2 import filters
from jinja2 import defaults

from .. import utils
from .. import x2pylua
from .. import x2pyxl

Ctx = utils.Ctx


VARX_REGEX = re.compile(r'\bx\b')


def get_pyexc_msg():
    exc_type, exc_value, exc_tb = sys.exc_info()
    del exc_tb
    exc_msg = traceback.format_exception_only(exc_type, exc_value)[0]
    return exc_msg.strip()


def expand_check_expr(expr, value):
    cur = 0
    buf = io.StringIO()
    for match in VARX_REGEX.finditer(expr):
        buf.write(expr[cur:match.start()])
        buf.write(str(value))
        cur = match.end()
    buf.write(expr[cur:])
    return buf.getvalue()


def ignore_return(func):
    @functools.wraps(func)
    def wrap(*args, **kwds):
        func(*args, **kwds)
        return ''
    return wrap


def do_json(value, indent=None):
    options = {}
    options['sort_keys'] = True
    options['ensure_ascii'] = False
    if indent is not None:
        options['indent'] = indent
    else:
        options['separators'] = (',', ':')
    return json.dumps(value, **options)


def do_lua(value, indent=None):
    return x2pylua.dumps(value, indent=indent)


@filters.contextfilter
def do_check(ctx, value, expr, msg=None):
    global_ = ctx.get_exported()
    global_['__builtins__'] = dict(
        abs=abs, all=all, any=any, bin=bin, bool=bool, chr=chr,
        divmod=divmod, float=float, hex=hex, int=int, len=len, max=max,
        min=min, oct=oct, ord=ord, round=round, str=str, sum=sum,
    )
    try:
        rval = eval(expr, global_, {'x': value})
    except Exception:
        msg1 = '校验表达式异常 {0} -- ' + get_pyexc_msg()
    else:
        if not rval:
            msg1 = msg or '数值校验不通过 {0}'
        else:
            msg1 = None
    if msg1:
        Ctx.error(msg1, repr(expand_check_expr(expr, value)))
    return value


@filters.environmentfilter
def do_next(env, value, num=1):
    cur = None
    itr = iter(value)
    for _ in range(num):
        try:
            cur = next(itr)
        except StopIteration:
            return env.undefined('不能获取第{0}个元素'.format(num))
    return cur


FILTERS = {
    'abs':          defaults.DEFAULT_FILTERS['abs'],
    'check':        do_check,
    'choice':       defaults.DEFAULT_FILTERS['random'],
    'd':            defaults.DEFAULT_FILTERS['default'],
    'default':      defaults.DEFAULT_FILTERS['default'],
    'e':            defaults.DEFAULT_FILTERS['escape'],
    'escape':       defaults.DEFAULT_FILTERS['escape'],
    'f':            defaults.DEFAULT_FILTERS['float'],
    'float':        defaults.DEFAULT_FILTERS['float'],
    'format':       defaults.DEFAULT_FILTERS['format'],
    'i':            defaults.DEFAULT_FILTERS['int'],
    'indent':       defaults.DEFAULT_FILTERS['indent'],
    'int':          defaults.DEFAULT_FILTERS['int'],
    'join':         defaults.DEFAULT_FILTERS['join'],
    'json':         do_json,
    'len':          defaults.DEFAULT_FILTERS['length'],
    'list':         defaults.DEFAULT_FILTERS['list'],
    'lower':        defaults.DEFAULT_FILTERS['lower'],
    'lua':          do_lua,
    'max':          defaults.DEFAULT_FILTERS['max'],
    'min':          defaults.DEFAULT_FILTERS['min'],
    'next':         do_next,
    'reverse':      defaults.DEFAULT_FILTERS['reverse'],
    'round':        defaults.DEFAULT_FILTERS['round'],
    's':            defaults.DEFAULT_FILTERS['string'],
    'sort':         defaults.DEFAULT_FILTERS['sort'],
    'str':          defaults.DEFAULT_FILTERS['string'],
    'sum':          defaults.DEFAULT_FILTERS['sum'],
    'trim':         defaults.DEFAULT_FILTERS['trim'],
    'unique':       defaults.DEFAULT_FILTERS['unique'],
    'upper':        defaults.DEFAULT_FILTERS['upper'],
    'xgroupby':     x2pyxl.xgroupby,
}
GLOBALS = {
    'range':        defaults.DEFAULT_NAMESPACE['range'],
    'dict':         defaults.DEFAULT_NAMESPACE['dict'],
    'cycler':       defaults.DEFAULT_NAMESPACE['cycler'],
    'joiner':       defaults.DEFAULT_NAMESPACE['joiner'],
    'namespace':    defaults.DEFAULT_NAMESPACE['namespace'],
    'loadws':       x2pyxl.load_worksheet,
    'output':       ignore_return(utils.open_as_stdout),
}


def get_j2exc_lineno():
    exc_tb = sys.exc_info()[2]
    stacks = traceback.extract_tb(exc_tb)
    exc_tb = None
    for frame in reversed(stacks):
        if frame[0] == '<template>':
            return frame[1]
    return 1


def print_version(ctx, param, value):
    if not value or ctx.resilient_parsing:
        return
    try:
        from .. import __version__
    except ImportError:
        __version__ = '???'
    click.echo(__version__)
    ctx.exit()


@click.command()
@click.argument('template', type=click.File('rb'))
@click.option('--verbose', is_flag=True)
@click.option('--version', is_flag=True, callback=print_version, expose_value=False, is_eager=True)
def main(template, verbose):
    if verbose:
        @Ctx.set_abort_handler
        def _at_abort():
            Ctx.error(traceback.format_exc())
            sys.exit(1)

    os.chdir(os.path.dirname(os.path.abspath(template.name)))

    j2data = template.read()
    encoding = utils.detect_encoding(j2data)
    try:
        j2txt = j2data.decode(encoding, errors="ignore")
    except (LookupError, TypeError):
        Ctx.set_ctx(template.name, 1)
        Ctx.abort('无法识别模板文件的文件编码')

    j2env = jinja2.Environment(
        extensions=[
            'jinja2.ext.do',
            'jinja2.ext.loopcontrols',
        ],
    )
    j2env.filters.clear()
    j2env.filters.update(FILTERS)
    j2env.globals.clear()
    j2env.globals.update(GLOBALS)
    try:
        j2res = j2env.from_string(j2txt).render()
    except Exception:
        Ctx.set_ctx(template.name, get_j2exc_lineno())
        Ctx.abort('处理模板文件时发生错误：{0}', get_pyexc_msg())
    else:
        print(j2res, file=sys.stdout, flush=True)
