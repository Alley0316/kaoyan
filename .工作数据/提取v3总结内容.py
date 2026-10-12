# -*- coding: utf-8 -*-
"""从 v3 总结生成器中抽取出「结构化内容 + 文本字面量」，供 v4 总结生成器复用。

做法：用 ast 精确解析源文件，按出现顺序收集顶层调用（band/sub/p/rich/table/example/gap），
把三引号文本抽为模板占位符，生成：
  .工作数据/v4总结模板.py   —— 含 TPL（结构）与 T（文本字典）的模板文件
这样 v4 总结的叙述文字与 v3 完全一致，只有统计表与书目表被替换，不会手抄出错。
"""
import ast
import io
import json
import pprint
import sys
import tokenize
from pathlib import Path

SRC = Path(sys.argv[1] if len(sys.argv) > 1 else
           '/home/zyf/考研/.高频考点工作/生成难度总结文档.py')
OUT = Path(sys.argv[2] if len(sys.argv) > 2 else '/home/zyf/考研/.工作数据/v4总结模板.py')
WANT = {'band', 'sub', 'p', 'rich', 'table', 'example', 'gap'}

lines = SRC.read_text().splitlines()
src = '\n'.join(lines)
tree = ast.parse(src)

# 收集三引号字符串区间（用 tokenize 判定引号类型）
triple = []
with open(SRC, 'rb') as fh:
    for tok in tokenize.tokenize(fh.readline):
        if tok.type == tokenize.STRING:
            s = tok.string.lstrip('rbuRBU')
            if s.startswith(('"""', "'''")):
                triple.append((tok.start[0], tok.end[0], tok.start[1], tok.end[1]))

texts = {}
offset = 0


def str_value(node):
    """把隐式拼接的多行字符串（BinOp/Add 链）还原为整段文本。"""
    if isinstance(node, ast.Constant) and isinstance(node.value, str):
        return node.value
    if isinstance(node, ast.BinOp) and isinstance(node.op, ast.Add):
        l, r = str_value(node.left), str_value(node.right)
        if l is not None and r is not None:
            return l + r
    return None


def hoist(node):
    """把多行或长字符串抽成占位符，返回新的值。

    源文件里的长段落是单引号字符串隐式拼接（AST 会合并成一个跨行 Constant），
    因此判定条件是「来源跨行」或「去空白后长度 ≥ 60」。
    """
    global offset
    v = str_value(node)
    if v is None:
        return None
    multi = getattr(node, 'end_lineno', 0) > getattr(node, 'lineno', 0)
    if multi or len(''.join(v.split())) >= 60:
        offset += 1
        key = f'TXT{offset:03d}'
        texts[key] = v
        return f'@@{key}@@'
    return v if not multi else v


def lit(x):
    """ast.unparse 得到的字面量字符串还原为真实值；占位符与非法表达式保留原样。"""
    if not isinstance(x, str) or x.startswith('@@'):
        return x
    try:
        v = ast.literal_eval(x)
    except Exception:
        return x
    return v


items = []
for stmt in tree.body:
    if not isinstance(stmt, ast.Expr) or not isinstance(stmt.value, ast.Call):
        continue
    call = stmt.value
    fn = getattr(call.func, 'id', None)
    if fn not in WANT:
        continue
    rec = {'fn': fn, 'line': stmt.lineno, 'text_keys': []}
    # 位置参数
    args = []
    for a in call.args:
        h = hoist(a)
        if h is not None:
            args.append(h)
            if isinstance(h, str) and h.startswith('@@'):
                rec['text_keys'].append(h.strip('@'))
        elif isinstance(a, ast.List):
            vals = []
            for el in a.elts:
                if isinstance(el, ast.List):
                    row = []
                    for x in el.elts:
                        hh = hoist(x)
                        row.append(hh if hh is not None else lit(ast.unparse(x)))
                        if isinstance(hh, str) and hh.startswith('@@'):
                            rec['text_keys'].append(hh.strip('@'))
                    vals.append(row)
                else:
                    hh = hoist(el)
                    vals.append(hh if hh is not None else lit(ast.unparse(el)))
                    if isinstance(hh, str) and hh.startswith('@@'):
                        rec['text_keys'].append(hh.strip('@'))
            args.append(vals)
        else:
            args.append(lit(ast.unparse(a)))
    rec['args'] = args
    rec['kwargs'] = {k.arg: lit(ast.unparse(k.value)) for k in call.keywords}
    items.append(rec)

# 头信息：TITLE 常量
title = None
for stmt in tree.body:
    if isinstance(stmt, ast.Assign) and getattr(stmt.targets[0], 'id', '') == 'TITLE':
        title = stmt.value.value
lines_out = [
    '# -*- coding: utf-8 -*-',
    '"""由 提取v3总结内容.py 自动生成：v3 总结的结构与文本字面量（勿手改）。"""',
    '',
    f'TITLE = {title!r}',
    'TPL = ' + pprint.pformat(items, width=200, sort_dicts=False),
    'T = ' + pprint.pformat(texts, width=200, sort_dicts=False),
    '',
]
OUT.write_text('\n'.join(lines_out))
print(f'抽出调用 {len(items)} 个（含文本 {len(texts)} 段），表 '
      f'{sum(1 for i in items if i["fn"] == "table")} 个，'
      f'例题 {sum(1 for i in items if i["fn"] == "example")} 个')
print('→', OUT)
