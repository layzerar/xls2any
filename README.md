## 如何打包

xls2any 基于 Python 开发，打包的主机需要安装 Python 的开发环境，目标主机无需安装 Python。

安装 Python 开发环境:

* 下载并安装 Python 3.6.5 以上版本
* 确保 Python 安装目录下的 Scripts 路径包含在系统的 PATH 变量里面。例如：C:\Python36\Scripts
* 在控制台中，使用命令 `pip install -r requirements.txt` 安装依赖包
* 运行批处理脚本 `.\scripts\pyinstaller_build.bat` 完成打包
* 把上一步打包生成的目录 xls2any-x.x.x 压缩成 Zip 包

## 安装说明

xls2any 采用压缩包方式发行软件。

xls2any 安装步骤：

* 解压 xls2any 压缩包，并存放到一个固定位置（最好不要放在桌面）
* 进入压缩包解压后的目录，以管理员身份运行批处理脚本 `assoc_j2ext.bat` 完成安装
* 进入 examples 目录，双击运行 datasource1_tolua.j2 文件。如果安装正确，当前目录下将会生成一个 datasource1.lua 文件

xls2any 工作原理：
* xls2any 通过读取后缀名为 .j2 的配置文件，从目标 Excel 文件中读取数据按配置要去生成目标 lua 文件
* .j2 配置文件的语法格式实际上一种叫 [Jinja2](http://jinja.pocoo.org/docs/2.10/templates/) 的模板文件，中文文档可见 [Jinja2](http://docs.jinkan.org/docs/jinja2/templates.html)

## 功能概述

以下我们以 datasource1_tolua.j2 为例简述如何编写 j2 文件

```jinja
{% set ws1 = loadws('./datasource1.xlsx', 'Sheet1', 2) -%}  --[0]读取 ./datasource1.xlsx 文件的 Sheet1 工作表作为变量 ws1，并指定第2行为字段头
{% set ws2 = loadws('./datasource1.xlsx', 'Sheet2', 2) -%}  --[0]读取 ./datasource1.xlsx 文件的 Sheet2 工作表作为变量 ws2，并指定第2行为字段头
{% do output('datasource1.lua') -%}                         --[0]将模板生成的结果保存到 datasource1.lua 文件
Test =
{
    {%- for row in ws1['3:'] %}     --[1]重复结构1开始，从第三行开始遍历工作表 ws1 的剩余行
    {
        mon_id = {{ row[1] | check('x < 1005') | lua }},        --[2]读取当前行的第1列数据，然后检查数据是否小于1005，最后再转为 lua 格式的数据
        mon_name = {{ row['B'] | lua }},                        --[2]读取当前行的 B 列数据，然后转为 lua 格式的数据
        mon_pos = {{ row.cut('@mon_pos_x', 3).asdict('pos_x', 'pos_y', 'pos_z') | lua }},   --[2]根据字段头，从 mon_pos_x 列开始顺序读取三列，分别用 pos_x, pos_y, pos_z 作为名称转为一个 lua 表
        drop_prob = {{ (row['@mon_drop_prob'] / 100) | lua }},  --[2]根据字段头，读取 mon_drop_prob 列的数据，除以100以后再转为 lua 格式的数据
        mon_drops = {
            {%- for grp in row.slc('@drop1_item', 2, 3) %}      --[2]重复结构2开始，从 drop1_item 列开始，按每两列一个分组，截取三个分组
            {{ grp.asdict('id', 'weight') | lua }},             --[3]将当前分组的两列，分别用 id，weight 作为名称转化为一个 lua 表
            {%- endfor %}                                       --[2]重复结构2结束
        },
        model_id = {{ ws2.vlookup(row['@mon_model'], 'B4:C13', 2) | lua }}, --[2]根据当前行 mon_model 列的数据去 ws2 工作表中查找(vlookup方式) B4:C13 区间的值
    },
    {%- endfor %}                   --[1]重复结构1结束
}
```

上面的示例中出现两种分隔符：

+ `{% ... %}` 用于执行代码语句
+ `{{ ... }}` 用于执行代码语句并把语句返回的结果输出到当前位置
+ 在 `{%` 符号后加上 `-` 表示，去除此符号之前的所有的空白字符
+ 在 `%}` 符号前加上 `-` 表示，去除此符号之后的所有的空白字符

上面的示例中出现的 `'3:'` 和 `'B4:C13'` 属于范围表达式，和 Excel 公式中的范围表达式类似。假设我们有个 Excel 工作表，总共有Z列和100行，那么：

+ `'3:'` 等价于 `'A3:Z100'`
+ `'3:9'` 等价于 `'A3:Z9'`
+ `'3:C'` 等价于 `'A3:C100'`
+ `'B:C'` 等价于 `'B1:C100'`

假如，我们在 `loadws` 函数中指定了某一行作为字段头，那么我们可以使用该行出现的字段名来定位列，如 `'@mon_drop_prob'` 。该风格的列名也可以用在范围表达式中。假设我们有个 Excel 工作表，`'@mon_pos_x'` 位于第K列，`'@mon_pos_z'` 位于第M列，那么：

+ `'3:@mon_pos_z$100'` 等价于 `'A3:M100'`
+ `'@mon_pos_x:@mon_pos_z$100'` 等价于 `'K1:M100'`
+ `'@mon_pos_x$3:@mon_pos_z$100'` 等价于 `'K3:M100'`

另外，我们可以使用过滤器来对值进行加工，例如 `{{ row['B'] | int | lua }}` 。该代码语句的意思是，先把 `row['B']` 的值传给 `int` 过滤器，再把该过滤器输出的结果传给 `lua` 过滤器。部分过滤器是带有可选参数的，比如我们可以给 `lua` 过滤器加上参数调整输出文本的缩进 `{{ row['B'] | int | lua(indent=4) }}` 。

注意，过滤器表达式 `{{ row['B'] | int | lua(indent=4) }}` 实际上等价于函数调用表达式 `{{ lua(int(row['B']), indent=4) }}`。

## 过滤器列举

#### 过滤器 `abs(number)`

<span style="margin-left: 1em;"/> 获取数值类型的绝对值。

#### 过滤器 `bool(value)`

<span style="margin-left: 1em;"/> 将数值类型转换为布尔类型。

#### 过滤器 `b(value)`

<span style="margin-left: 1em;"/> 同 `bool` 过滤器。

#### 过滤器 `check(expr, msg=None)`

<span style="margin-left: 1em;"/> 检查参数 `expr` 表达式是否未真，非真则打印参数 `msg` 指定的错误信息。

#### 过滤器 `choice(seq)`

<span style="margin-left: 1em;"/> 返回序列中随机挑选的一个元素。

#### 过滤器 `default(value, default_value=u'', boolean=False)`

<span style="margin-left: 1em;"/> 如果目标数值未定义，则返回参数 `default_value` 指定的值。

#### 过滤器 `d(value, default_value=u'', boolean=False)`

<span style="margin-left: 1em;"/> 同 `default` 过滤器。

#### 过滤器 `escape(s)`

<span style="margin-left: 1em;"/> 将字符串转换为 HTML 安全的字符串。

#### 过滤器 `e(s)`

<span style="margin-left: 1em;"/> 同 `escape` 过滤器。

#### 过滤器 `float(value, default=0.0)`

<span style="margin-left: 1em;"/> 将数值变换成浮点数，如果转换失败则返回参数 `default` 指定的值。

#### 过滤器 `f(value, default=0.0)`

<span style="margin-left: 1em;"/> 同 `float` 过滤器。

#### 过滤器 `format(value, *args, **kwargs)`

<span style="margin-left: 1em;"/> 文本格式化过滤器。`{{ "%s - %s" | format("Hello?", "Foo!") }}` 输出 `Hello? - Foo!`。

#### 过滤器 `indent(s, width=4, indentfirst=False)`

<span style="margin-left: 1em;"/> 文本对齐过滤器。

#### 过滤器 `int(value, default=0)`

<span style="margin-left: 1em;"/> 将数值变换成整型，如果转换失败则返回参数 `default` 指定的值。

#### 过滤器 `i(value, default=0)`

<span style="margin-left: 1em;"/> 同 `int` 过滤器。

#### 过滤器 `join(value, d=u'', attribute=None)`

<span style="margin-left: 1em;"/> 将序列用参数 `d` 指定的连接符连接起来。`{{ [1, 2, 3] | join('|') }}` 输出 `1|2|3`。

#### 过滤器 `json(value, indent=None, closed=True)`

<span style="margin-left: 1em;"/> 将对象转换成合法的 JSON 文本。

#### 过滤器 `len(object)`

<span style="margin-left: 1em;"/> 返回目标对象的长度。

#### 过滤器 `list(value)`

<span style="margin-left: 1em;"/> 将数值转换成列表。

#### 过滤器 `lower(s)`

<span style="margin-left: 1em;"/> 将字符串转换成小写字符串。

#### 过滤器 `lua(value, indent=None, closed=True)`

<span style="margin-left: 1em;"/> 将对象转换成合法的 Lua 文本。

#### 过滤器 `max(...)`

<span style="margin-left: 1em;"/> 输出序列中的最大值。`{{ 1 | max(2) }}` 输出 `2`。

#### 过滤器 `min(...)`

<span style="margin-left: 1em;"/> 输出序列中的最小值。`{{ 1 | min(2) }}` 输出 `1`。

#### 过滤器 `num(value, default=0)`

<span style="margin-left: 1em;"/> 将数值转换数值类型（整型或浮点），如果转换失败则返回参数 `default` 指定的值。

#### 过滤器 `n(value, default=0)`

<span style="margin-left: 1em;"/> 同 `num` 过滤器

#### 过滤器 `clamp(value, lower, upper)`

<span style="margin-left: 1em;"/> 确保数值不超过参数 `upper` 的值且不小于参数 `lower` 的值。

#### 过滤器 `next(value, num=1)`

<span style="margin-left: 1em;"/> 将迭代器向前跳参数 `num` 指定的个数。

#### 过滤器 `reverse(value)`

<span style="margin-left: 1em;"/> 反转目前序列的迭代顺序。

#### 过滤器 `round(value, precision=0, method='common')`

<span style="margin-left: 1em;"/> 将目标数值进行四舍五入。

#### 过滤器 `sort(value, reverse=False, case_sensitive=False, attribute=None)`

<span style="margin-left: 1em;"/> 排序目标序列。

#### 过滤器 `splitf(value, nth=1, fs=None)`

<span style="margin-left: 1em;"/> 文本段提取过滤器。`{{ "abc 123 xyz" | splitf(2, fs=' ') }}` 输出 `123`。

#### 过滤器 `str(object)`

<span style="margin-left: 1em;"/> 将目标对象转换为字符串。

#### 过滤器 `s(object)`

<span style="margin-left: 1em;"/> 同 `str` 过滤器

#### 过滤器 `sum(seq, attribute=None, start=0)`

<span style="margin-left: 1em;"/> 计算目标序列的数值之和。

#### 过滤器 `trim(value)`

<span style="margin-left: 1em;"/> 去除字符串两端的空白字符。

#### 过滤器 `unique(value, case_sensitive=False, attribute=None)`

<span style="margin-left: 1em;"/> 对目标序列做去重处理。

#### 过滤器 `upper(s)`

<span style="margin-left: 1em;"/> 将字符串转换成小写字符串。

#### 过滤器 `xgroupby(rows, *keys, asc=True, required=True)`

<span style="margin-left: 1em;"/> 按指定的字段对行对象集进行分组。

#### 过滤器 `xrequire(rows, *keys, over=0)`

<span style="margin-left: 1em;"/> 剔除掉行对象集中指定字段为空的行。

## 函数列举

#### 全局函数 `loadws(filepath, sheetname, head=0)`

<span style="margin-left: 1em;"/> 加载指定路径指定名称的 Excel 工作页对象。

#### 全局函数 `output(filename, encoding='utf-8')`

<span style="margin-left: 1em;"/> 指定输出的文件路径。

## 工作页对象函数列举

#### 对象函数 `__getitem__(key)`

<span style="margin-left: 1em;"/>

#### 对象函数 `hidx(key, multi=False, hoff=0, hmax=sys.maxsize)`

<span style="margin-left: 1em;"/>

#### 对象函数 `keys(*idxs, token=True)`

<span style="margin-left: 1em;"/>

#### 对象函数 `rowx(vidx=None)`

<span style="margin-left: 1em;"/>

#### 对象函数 `valx(expr)`

<span style="margin-left: 1em;"/>

#### 对象函数 `rehead(vidx, hbeg=None, vbeg=None, hend=None, vend=None)`

<span style="margin-left: 1em;"/>

#### 对象函数 `select(expr)`

<span style="margin-left: 1em;"/>

#### 对象函数 `locate(ltag, htag, loff=1, hoff=-1)`

<span style="margin-left: 1em;"/>

#### 对象函数 `findall(val1, tab, *keys)`

<span style="margin-left: 1em;"/>

#### 对象函数 `findone(val1, tab, *keys)`

<span style="margin-left: 1em;"/>

#### 对象函数 `vlookup(val1, tab, idx)`

<span style="margin-left: 1em;"/>

#### 对象函数 `chkuniq(tab)`

<span style="margin-left: 1em;"/>

## 行对象函数列举

#### 对象函数 `__getitem__(key)`

<span style="margin-left: 1em;"/>

#### 对象函数 `hidx(key, multi=False)`

<span style="margin-left: 1em;"/>

#### 对象函数 `expr(key)`

<span style="margin-left: 1em;"/>

#### 对象函数 `keys(*idxs, token=True)`

<span style="margin-left: 1em;"/>

#### 对象函数 `vals(*keys)`

<span style="margin-left: 1em;"/>

#### 对象函数 `valx(key)`

<span style="margin-left: 1em;"/>

#### 对象函数 `cut(key, size=1)`

<span style="margin-left: 1em;"/>

#### 对象函数 `slc(key, size=1, num=0)`

<span style="margin-left: 1em;"/>
