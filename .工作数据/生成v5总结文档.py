# -*- coding: utf-8 -*-
"""生成 v5《四校真题考点与难度对比》总结。

叙述文字与难度判断沿用 v3／v4（由 提取v3总结内容.py 精确抽取的模板文本，避免手抄），
只替换与 v5 分册口径相关的部分：
  ① 封面版本说明与副标题；② 编目统计表（改为一校一册口径）；
  ③ 新增「v5 分段与书目变更对照表」；④ 安徽一节的范围来源按官方目录逐年改写
  （v4 用的是 2021 年书目，已按官方目录更正为按年份分列）；
  ⑤ 参考书单总表（按册与年份段）；⑥ 结尾目录说明。
输出：v5/四校真题考点与难度对比_v5.docx
"""
import importlib.util
import json
import re
from pathlib import Path

W = Path(__file__).resolve().parent
O = Path('/home/zyf/考研/v5')
VERSION = 'v5'

# 复用考点册生成器的样式函数，保证五份文档视觉一致
spec = importlib.util.spec_from_file_location('bookgen', W / '生成v5考点文档.py')
bg = importlib.util.module_from_spec(spec)
spec.loader.exec_module(bg)

T = json.loads((W / 'v5考点数据.json').read_text())
MANIFEST = json.loads((W / 'v5册清单.json').read_text())
ORDER = ['安徽大学', '黑龙江大学', '云南民族大学', '西安工业大学']
DATE_CN = '2026 年 10 月 12 日'

tpl_spec = importlib.util.spec_from_file_location('v3tpl', W / 'v4总结模板.py')
tpl = importlib.util.module_from_spec(tpl_spec)
tpl_spec.loader.exec_module(tpl)


def resolve_key(value, keys):
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
    fn = node['fn']
    args = resolve_key(node['args'], keys)
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


# ------------------------------------------------------------ v5 专属内容
def seg_rows(school):
    return T[school]['segments']


STATS_HEADERS = ['学校（册）', '考试科目段', '唯一考点', '段内条目', '证据（去重）',
                 '证据（段内合计）', '回忆版', '页数']
STATS_ROWS = []
PAGE_CNT = json.loads((W / 'v5页数.json').read_text()) if (W / 'v5页数.json').exists() else {}
for m in MANIFEST:
    first = True
    for s in m['segments']:
        STATS_ROWS.append([m['school'] if first else '', s['label'] + f'（{s["years"]}）',
                           str(m['unique_topics']) if first else '',
                           str(s['topics']), '', str(s['evidence']), str(s['recall']),
                           str(PAGE_CNT.get(m['file'], '')) if first else ''])
        first = False
    STATS_ROWS.append(['', '小计（去重口径：唯一考点／段内条目）',
                       str(m['unique_topics']), str(m['instances']),
                       str(m.get('unique_evidence', '')), str(m['evidence']),
                       str(m['recall']), str(PAGE_CNT.get(m['file'], ''))])
STATS_ROWS.append(['合计', f'{len(MANIFEST)} 册',
                   str(sum(m['unique_topics'] for m in MANIFEST)),
                   str(sum(m['instances'] for m in MANIFEST)),
                   str(sum(m.get('unique_evidence', 0) for m in MANIFEST)),
                   str(sum(m['evidence'] for m in MANIFEST)),
                   str(sum(m['recall'] for m in MANIFEST)),
                   str(sum(PAGE_CNT.get(m['file'], 0) for m in MANIFEST))])
STATS_WIDTHS = [1.9, 5.0, 1.5, 1.5, 1.7, 1.9, 1.2, 1.0]

SEG_HEADERS = ['学校', '科目段（代码）', '年份', '本段参考书', '段内条目']
SEG_ROWS = []
for school in ORDER:
    for i, seg in enumerate(seg_rows(school), 1):
        books = '；'.join(b['name'].split('（')[0] for b in seg['books'])
        code = seg['code'].split('；')[0] if '；' in seg['code'] else seg['code']
        SEG_ROWS.append([school if i == 1 else '', f'{seg["label"]}｜代码 {code}',
                         seg['years'], books, str(seg['stats']['topics'])])
SEG_WIDTHS = [2.0, 4.4, 2.0, 6.4, 1.6]

BOOK_HEADERS = ['学校／年份段', '参考书目', '版本依据', '本段条目']
BOOK_ROWS = []
for school in ORDER:
    for seg in seg_rows(school):
        label = f'{school}\n{seg["label"]}\n{seg["years"]}'
        for i, b in enumerate(seg['books'], 1):
            pub = '' if b.get('pub') in ('—', '-', '', None) else f'（{b["pub"]}）'
            BOOK_ROWS.append([label if i == 1 else '', b['name'] + pub, b['edition'],
                              str(b['local_max'])])
    BOOK_ROWS.append(['', '证据级别与卷面代码：见该册「二、参考书目（按年份段分列）」',
                      f'该校唯一考点 {T[school]["stats"]["unique_topics"]} 条', ''])
BOOK_WIDTHS = [3.4, 5.6, 6.2, 1.6]

TPL_STATS_TEXT = (
    'v5 把交付形态改为一校一册（4 册），册内再按「考试科目＋科目代码＋参考书」分段；'
    '参考书目按年份分列，不再用某一年的书目覆盖全部年份。'
    '读这张表要分清两个口径：①「唯一考点」是按考点去重后的数量，四校合计 {uq} 条；'
    '②「段内条目」是各段各有多少条要背，合计 {inst} 条，比前者多，因为同一考点若在不同'
    '年份段对应的书里都出现，会在各段各列一次，各段只带本段证据。'
    '证据同理：按「卷名＋题号＋原文」去重后四校合计 {uev} 条，段内合计 {ev} 条，'
    '后者是各段各写一次的口径。')
TPL_NOTE_TEXT = (
    '使用这 {n} 册时请注意四条：① 同一学科按不同年份计次，同年多题不重复加次数，'
    '「年次」是已确认证据的下限，不是预测概率；② 条目包含考点、专题与题型，'
    '不是互斥的原子考点，证据条数不等于独立题数；③ 回忆卷没有记录到的题，'
    '不能直接视为「没有考」；④ 科目代码与书目都随年份变过，看某一年的题'
    '要回到该年对应那一段的书目，不要用最新书目去评旧卷。')


def main():
    bg.DOC = bg.setup('四校真题考点与难度对比')
    d = bg.DOC
    keys = set()
    st = {m['school']: m for m in MANIFEST}
    for node in tpl.TPL:
        line = node['line']
        if line == 146:
            bg.p(d, f'依据四校《历年真题打印复核版》与 v5《高频考点》{len(MANIFEST)} 册'
                    f'（一校一册，册内按科目与参考书分段）；先给结论，再给依据与范围来源',
                 9.5, False, after=2, color='555555')
            continue
        if line == 147:
            bg.p(d, f'{DATE_CN}　v5（一校一册；册内按「考试科目＋科目代码＋参考书」分段；'
                    f'参考书目按年份分列；补正 19 条补录证据的卷别与年份）',
                 9.5, False, after=10, color='555555')
            continue
        if line == 210:
            bg.p(d, TPL_STATS_TEXT.format(
                uq=sum(m['unique_topics'] for m in MANIFEST),
                inst=sum(m['instances'] for m in MANIFEST),
                uev=sum(m.get('unique_evidence', 0) for m in MANIFEST),
                ev=sum(m['evidence'] for m in MANIFEST)), 10, after=4)
            continue
        if line == 212:
            bg.table(d, STATS_HEADERS, STATS_ROWS, STATS_WIDTHS, 8.0)
            continue
        if line == 221:
            bg.p(d, TPL_NOTE_TEXT.format(n=len(MANIFEST)), 10, True, keep=True, after=3)
            continue
        if line == 233:      # 安徽范围来源：按官方目录逐年改写（v4 用的是 2021 年书目）
            bg.p(d, '安徽大学的分段是按官方专业目录逐年核出来的，三段书目都不同：'
                    '2012—2016 年为 605（配套阅读与写作 807），现代汉语指定胡裕树主编'
                    '《现代汉语》（上海教育出版社 1995 年版，2007—2015 年）、2016 年换成黄伯荣、'
                    '廖序东《现代汉语》增订五版，古代汉语这一段及此前一贯是王力主编'
                    '《古代汉语》（中华书局 1999 年第三版）；2017—2018 年科目代码改为 615、'
                    '阅读与写作改为 808，现代汉语仍为黄廖增订五版；2019 年起改为 616／809，'
                    '古代汉语换成黄德宽主编《古代汉语》（高等教育出版社 2015 年版），'
                    '现代汉语 2020 年起换成增订六版。'
                    '2018 年及以后有安徽大学官网文件（官方原件），2012—2017 年为权威转载。',
                 9.5, after=3)
            continue
        if line == 236:
            bg.p(d, '两条必须记住：① 2007—2018 年「古代汉语」一直是王力本，'
                    '2019 年才换成黄德宽本——旧教辅里出现的郭锡良《古代汉语》不是安徽大学'
                    '任何一年的指定书目，只出现在卷内解析的页码标注里，不能当范围依据；'
                    '② 2004—2011 年该科的初试参考书目原件没有查到，v5 册内这一段如实留空，'
                    '不用后一年的书目倒推。', 9.5, after=8, color='555555')
            continue
        if line == 296:
            bg.sub(d, '1. 范围来源（2021、2026 官方大纲原件）')
            continue
        if line == 300:      # 黑龙江 2021 当年书目表 → 保留原表结构但标注年份口径
            bg.table(d, ['科目（段）', '参考书目', '年份口径'],
                     [['742 现代汉语', '黄伯荣、廖序东《现代汉语》增订第五版（高等教育出版社）',
                       '2015—2024 年适用；2026 年官方大纲已改为邵敬敏《现代汉语通论》第 3 版'],
                      ['842·古代汉语', '王力主编《古代漢語》一、二册（中华书局校订重排本）',
                       '2011—2026 年一贯'],
                      ['842·语言学理论',
                       '《语言学纲要》（修订版），叶蜚声、徐通锵著，王洪君、李娟修订（北京大学）',
                       '2010—2026 年一贯']],
                     [3.0, 7.6, 6.4], 8.5)
            continue
        if line == 645:
            bg.p(d, f'v5 目录含本总结与配套《高频考点》{len(MANIFEST)} 册'
                    f'（共 {len(MANIFEST) + 1} 个 PDF）：一校一册，'
                    '安徽大学 1 册（册内 5 段：语言文学基础 2004—2011、605、615、616 与阅读与写作）、'
                    '黑龙江大学 1 册（3 段：现汉 742、古理 842 的古汉部分与语理部分）、'
                    '云南民族大学 1 册（2 段：现汉 719、语理 815）、'
                    '西安工业大学 1 册（3 段：语言综合 902、语言综合 816、文学评论写作 702）；'
                    '各段正文按参考书分块，编号在每本书内各自从 1 开始。', 9.5)
            continue
        run_node(node, keys)
        if line == 155:      # 三节之后插入分段总表
            bg.gap(d, 4)
            bg.band(d, '四、v5 分段与书目变更对照表')
            bg.p(d, '下表把 v5 每一册的分段一次列清：哪一段是哪个科目、哪个科目代码、哪些年份、'
                    '对应哪本书。四校的科目代码与参考书都随年份变过，这张表就是'
                    '「某年的题该用哪本书评」的依据。', 10, after=5)
            bg.table(d, SEG_HEADERS, SEG_ROWS, SEG_WIDTHS, 8.2)
            bg.gap(d, 4)
            bg.p(d, '书目变更要点：安徽 现代汉语 胡裕树本（2007—2015）→黄廖增订五版'
                    '（2016—2019）→增订六版（2020—2021），古代汉语 王力本（2007—2018）'
                    '→黄德宽本（2019—2021）；黑龙江 2026 年官方大纲把 742 现汉改为邵敬敏'
                    '《现代汉语通论》第 3 版（属 2025 年之后的范围变动）；'
                    '云南 815 语理 2026 年起改《语言学纲要》2025 年版；'
                    '西安 古代汉语 2024 为商务印书馆 2018 年本、2025 与 2026 为 1999 年修订本；'
                    '安徽《阅读与写作》官方从未列指定教材，表中如实留空。', 9.5, after=3)
            bg.band(d, '五、参考书单总表（按册与年份段）')
            bg.p(d, '下表把各册各段的参考书目合并列出，便于横向查看四校的指定范围与版本差异；'
                    '每段正文的编号即按本表顺序分块。', 10, after=5)
            bg.table(d, BOOK_HEADERS, BOOK_ROWS, BOOK_WIDTHS, 8.2)
            bg.gap(d, 6)
    # 章节序号顺延：新增两节后，原四～十改为六～十二
    shift = {'四、安徽大学': '六、安徽大学', '五、黑龙江大学': '七、黑龙江大学',
             '六、云南民族大学': '八、云南民族大学', '七、西安工业大学': '九、西安工业大学',
             '八、四校之间': '十、四校之间', '九、参考来源与核查状态': '十一、参考来源与核查状态',
             '十、限制与未核事项': '十二、限制与未核事项'}
    for para in d.paragraphs:
        t = para.text
        for a, b in shift.items():
            if t.startswith(a):
                for r in para.runs:
                    if a in r.text:
                        r.text = r.text.replace(a, b, 1)
                break
    out = O / f'四校真题考点与难度对比_{VERSION}.docx'
    d.save(out)
    print('已生成', out.name, '段落', len(d.paragraphs), '表格', len(d.tables))
    unused = [k for k in tpl.T if k not in keys]
    print('引用文本段', len(keys), '未引用', unused)


if __name__ == '__main__':
    main()
