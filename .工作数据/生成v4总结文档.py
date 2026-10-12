# -*- coding: utf-8 -*-
"""生成 v4《四校真题考点与难度对比》总结。

叙述文字与判断沿用 v3（由 提取v3总结内容.py 精确抽取，避免手抄），只替换：
  ① 封面版本说明；② 编目统计表（按 v4 分册口径重算）；③ 新增「参考书单总表」一节；
  ④ 章节序号顺延；⑤ 结尾目录说明改为 v4 的册数与文件数。
输出：v4/四校真题考点与难度对比_v4.docx
"""
import importlib.util
import json
import re
from pathlib import Path

W = Path(__file__).resolve().parent
O = Path('/home/zyf/考研/v4')

# 复用考点册生成器的样式函数，保证五份文档视觉一致
spec = importlib.util.spec_from_file_location('bookgen', W / '生成v4考点文档.py')
bg = importlib.util.module_from_spec(spec)
spec.loader.exec_module(bg)

T = json.loads((W / 'v4考点数据.json').read_text())
MANIFEST = json.loads((W / 'v4册清单.json').read_text())
PAGE_CNT = json.loads((W / 'v4页数.json').read_text())
ORDER = ['安徽大学', '黑龙江大学', '云南民族大学', '西安工业大学']
DATE_CN = '2026 年 10 月 12 日'

tpl_spec = importlib.util.spec_from_file_location('v3tpl', W / 'v4总结模板.py')
tpl = importlib.util.module_from_spec(tpl_spec)
tpl_spec.loader.exec_module(tpl)


def resolve_key(value, keys):
    """把 '@@TXTxxx@@' 换成实际文本，并记录用到的键。"""
    if isinstance(value, str):
        m = re.fullmatch(r'@@(TXT\d+)@@', value)
        if m:
            keys.add(m.group(1))
            return tpl.T[m.group(1)]
        return value
    if isinstance(value, list):
        return [resolve_key(v, keys) for v in value]
    return value


def num(s):
    return float(s) if isinstance(s, str) else s


def run_node(node, keys):
    """执行 TPL 中的一个调用节点。"""
    fn = node['fn']

    args = resolve_key(node['args'], keys)          # 首项为 d
    body = args[1:]
    kw = dict(node['kwargs'])
    if fn == 'gap':
        bg.gap(bg.DOC, num(body[0]))
    elif fn == 'band':
        bg.band(bg.DOC, body[0])
    elif fn == 'sub':
        bg.sub(bg.DOC, body[0])
    elif fn == 'p':
        bg.p(bg.DOC, *body, **kw)
    elif fn == 'rich':
        bg.rich(bg.DOC, body[0], **kw)
    elif fn == 'table':
        bg.table(bg.DOC, body[0], body[1], [num(x) for x in body[2]],
                 size=num(body[3]) if len(body) > 3 else 8.5)
    elif fn == 'example':
        bg.example(bg.DOC, *body, **kw)
    else:
        raise KeyError(fn)


# ------------------------------------------------------------ v4 专属内容
STATS_HEADERS = ['学校', '科目（册）', '考点条目', '证据记录', '回忆版', '参考书', '册页数']
STATS_ROWS = []
for m in MANIFEST:
    STATS_ROWS.append([m['school'], m['paper'], str(m['topics']), str(m['evidence']),
                       str(m['recall']), str(len(m['books'])), str(PAGE_CNT.get(m['file'], '?'))])
STATS_ROWS.append(['合计', f'{len(MANIFEST)} 册', str(sum(m["topics"] for m in MANIFEST)),
                   str(sum(m["evidence"] for m in MANIFEST)),
                   str(sum(m["recall"] for m in MANIFEST)), '11 本（含 1 册无指定书）',
                   str(sum(PAGE_CNT.get(m['file'], 0) for m in MANIFEST))])
STATS_WIDTHS = [2.4, 4.4, 1.7, 1.7, 1.4, 2.0, 1.4]

BOOK_HEADERS = ['学校／科目（册）', '参考书目', '本册条目']
BOOK_ROWS = []
for school in ORDER:
    for paper in T[school]:
        label = f'{school}\n{paper["paper"]}'
        for i, b in enumerate(paper['books'], 1):
            pub = '' if b['pub'] in ('—', '-', '') else f'（{b["pub"]}）'
            BOOK_ROWS.append([label if i == 1 else '', b['name'] + pub, str(b['local_max'])])
        BOOK_ROWS.append(['', f'证据级别：{paper.get("ev_kind") or "见正文"}；'
                              f'真题范围 {paper["years"]}', f'小计 {paper["stats"]["topics"]}'])
BOOK_WIDTHS = [3.6, 10.6, 2.8]

TPL_STATS_TEXT = ('v4 把考点册改为「一个 PDF 册 ＝ 一个考试科目」，正文按参考书分块，'
                  '每本书内编号各自从 1 开始；并把此前单列的「单次关注」全部并入正文，'
                  '只确认一次的条目同样给出类型、掌握要求、近年情况与全部出处。'
                  '因此 v4 的条目数（{total} 条）与 v3 的「复现主题＋单次关注」（249＋62）不是同一口径，'
                  '不能直接相减比较。')
TPL_NOTE_TEXT = ('使用这 {n} 册时请注意三条：① 同一学科按不同年份计次，同年多题不重复加次数，'
                 '「年次」是已确认证据的下限，不是预测概率；② 条目包含考点、专题与题型，'
                 '不是互斥的原子考点，证据条数不等于独立题数；③ 回忆卷没有记录到的题，'
                 '不能直接视为「没有考」。')


def main():
    bg.DOC = bg.setup('四校真题考点与难度对比')
    d = bg.DOC
    keys = set()
    total_topics = sum(m['topics'] for m in MANIFEST)
    for node in tpl.TPL:
        line = node['line']
        if line == 146:      # 副标题
            bg.p(d, f'依据四校《历年真题打印复核版》与 v4《高频考点》{len(MANIFEST)} 册；'
                    f'先给结论，再给依据与范围来源', 9.5, False, after=2, color='555555')
            continue
        if line == 147:      # 版本行
            bg.p(d, f'{DATE_CN}　v4（新增：参考书目前置于各册；考点编号按参考书各自从 1 起；'
                    f'全部条目含仅确认一次者；补黑龙江 2020 年现汉段与 2019 年尾题）',
                  9.5, False, after=10, color='555555')
            continue
        if line == 210:      # 统计说明
            bg.p(d, TPL_STATS_TEXT.format(total=total_topics), 10, after=4)
            continue
        if line == 212:      # 统计表
            bg.table(d, STATS_HEADERS, STATS_ROWS, STATS_WIDTHS, 8.2)
            continue
        if line == 221:      # 使用提示
            bg.p(d, TPL_NOTE_TEXT.format(n=len(MANIFEST)), 10, True, keep=True, after=3)
            continue
        if line == 644:
            bg.p(d, '页码均为各校《历年真题打印复核版》PDF 文件页码，'
                    'v4 各册的出处行可直接按页码回到原题核对。', 9.5)
            continue
        if line == 645:
            bg.p(d, f'v4 目录含本总结与配套《高频考点》{len(MANIFEST)} 册（共 {len(MANIFEST) + 1} 个 PDF）：'
                    f'安徽大学 2 册、黑龙江大学 2 册、云南民族大学 2 册、西安工业大学 2 册；'
                    f'各册正文按参考书分块，编号在每本书内各自从 1 开始。', 9.5)
            continue
        run_node(node, keys)
        if line == 227:      # 三节之后插入「参考书单总表」，后续章节序号顺延
            bg.band(d, '四、参考书单总表（各册前置书目的合并视图）')
            bg.p(d, '下表把 v4 各册前置的参考书目合并列出，便于横向查看四校的指定范围与版本差异；'
                    '每册正文的编号即按本表顺序分块。', 10, after=5)
            bg.table(d, BOOK_HEADERS, BOOK_ROWS, BOOK_WIDTHS, 8.5)
            bg.gap(d, 4)
            bg.p(d, '版本要点：黑龙江 742 现汉 2026 起改为邵敬敏《现代汉语通论》第 3 版；'
                    '云南 815 语理 2026 起改用《语言学纲要》2025 年版；'
                    '西安古代汉语 2024 为商务印书馆 2018 年本、2025 与 2026 为 1999 年修订本；'
                    '安徽《阅读与写作》官方未列指定教材，本表如实留空。', 9.5, after=8)
    # 章节序号顺延：插入新节后，原四～十改为五～十一
    shift = {'四、安徽大学': '五、安徽大学', '五、黑龙江大学': '六、黑龙江大学',
             '六、云南民族大学': '七、云南民族大学', '七、西安工业大学': '八、西安工业大学',
             '八、四校之间': '九、四校之间', '九、参考来源与核查状态': '十、参考来源与核查状态',
             '十、限制与未核事项': '十一、限制与未核事项'}
    for para in d.paragraphs:
        t = para.text
        for a, b in shift.items():
            if t.startswith(a):
                for r in para.runs:
                    if a in r.text:
                        r.text = r.text.replace(a, b, 1)
                break
    out = O / '四校真题考点与难度对比_v4.docx'
    d.save(out)
    print('已生成', out.name, '段落', len(d.paragraphs), '表格', len(d.tables))
    used = sorted(keys)
    unused = [k for k in tpl.T if k not in keys]
    print('引用文本段', len(used), '未引用', unused)


if __name__ == '__main__':
    main()
