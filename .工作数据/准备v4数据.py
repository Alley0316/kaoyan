# -*- coding: utf-8 -*-
"""v4 数据准备：按「考试科目 → 参考书 → 条目」重组复核考点，并补入 v3 曾排除的原卷证据。

输入：.工作数据/复核考点.json（v3 统计输入，勿重跑上游脚本覆盖）
      .高频考点工作/{学校}_历年真题_打印版.json（原卷题面，用于补证）
输出：.工作数据/v4考点数据.json

v4 与 v3 的差别：
1. 一个 PDF 册 = 一个考试科目（试卷）；正文按参考书分块，每块编号各自从 1 开始。
2. 全部条目（含只确认一次的）都进入正文，不再单列「单次关注」。
3. 补入两处 v3 判为「不计」的原卷证据：
   - 黑龙江 2020 年卷末「第二部分 现代汉语语音、文字、词汇（75 分）」（打印版 p42，pid847—866）
   - 黑龙江 2019 年古理卷末 pid774「结合实例说明主语的语义类型」（打印版 p38，15 分）
"""
import json
from pathlib import Path
from collections import OrderedDict

W = Path(__file__).resolve().parent
SRC = json.loads((W / '复核考点.json').read_text())
RAW = {s: json.loads((W.parent / '.高频考点工作' / f'{s}_历年真题_打印版.json').read_text())
       for s in ['安徽大学', '黑龙江大学', '云南民族大学', '西安工业大学']}

# ---------------------------------------------------------------- 参考书表
# 证据级别：官方＝学校研究生院原件或官方大纲原件；转载＝权威转载全文（非官网原件）
ANHUI_NOTE = ('转载。来源为安徽大学 2021 年硕士研究生招生专业目录的转载全文'
              '（启航考研 jixun.iqihang.com/schoolzyml/20226650.html），非官网原件。'
              '现有真题止于 2021 年，故以上述旧「语言文学基础」范围为准；'
              '该校现行科目设置已变更，不能用新科目反推这批旧卷。')
HLJU_NOTE = ('官方。842、742 两份 2026 年大纲为官方大纲原件（文件内附参考书目与题型结构）；'
             '742 已改为邵敬敏《现代汉语通论》第 3 版与 20／60／50／20 结构，'
             '与 2021 年（黄廖增订第五版，25／70／40／15）不是同一套范围。'
             '2021 年当年大纲为权威转载（原始来源为黑龙江大学研究生院 2021 年招生章程通知），官方站点直连返回 403。')
YNU_NOTE = ('官方。《2025 年硕士研究生招生初试参考书目》《2026 年硕士研究生招生初试参考书目》'
            '均取得研究生院官方 PDF 原件。2024 年书目为转载，其 719／815 与 2025 官方原件一致，可用于 2024、2025 两卷。')
XATU_NOTE = ('语言综合：2024、2026 为官方书目页（西安工业大学研究生院）；2025 为权威转载，'
             '转载页标注来源为西安工业大学研究生院；三处现代汉语一致，古代汉语 2024 指定商务印书馆 2018 年本、'
             '2025 与 2026 为 1999 年修订本，评卷须按各自年份版本理解。'
             '文学评论写作：2024、2026 为官方书目页，2025 为权威转载；三处文学理论书目一致。')

PAPERS = OrderedDict([
    ('安徽大学', [
        dict(paper='语言文学基础（616）', years='2004—2021（18 年）', pages='p1—p67',
             note='2021 年专业目录所列「语言文学基础」参照书目。现有真题止于 2021 年，'
                  '2004—2021 各年基础卷均归入本册；2017 名解 20 选 10、2018 简答 A／B 整组选一、'
                  '2019 语言／文学整组任选、2021 明确「不可选做」，选做制度逐年不同。',
             books=[
                 dict(key='xdh', name='黄伯荣、廖序东《现代汉语》（增订六版）',
                      pub='高等教育出版社', edition='2021 年专业目录标注「六版」',
                      disc='基础', sections=['现代汉语 绪论与语音', '现代汉语 词汇与语法']),
                 dict(key='gdh', name='黄德宽《古代汉语》',
                      pub='高等教育出版社', edition='2021 年专业目录标示，版本页未标年份',
                      disc='基础', sections=['古代汉语 文字音韵与训诂', '古代汉语 句法与阅读']),
                 dict(key='gwx', name='袁行霈《中国文学史》（第三版）',
                      pub='高等教育出版社', edition='2021 年专业目录标注「三版」',
                      disc='基础', sections=['古代文学']),
                 dict(key='xdwx', name='丁帆《中国新文学史》',
                      pub='高等教育出版社', edition='2021 年专业目录标示，版本页未标年份',
                      disc='基础', sections=['现当代文学']),
                 dict(key='wxll', name='童庆炳《文学理论教程》（第四版）',
                      pub='高等教育出版社', edition='2021 年专业目录标注「四版」',
                      disc='基础', sections=['文学理论']),
             ], source_note=ANHUI_NOTE, ev_kind='转载：2021 年专业目录转载全文，非官网原件'),
        dict(paper='阅读与写作（809）', years='2007—2021（15 年）', pages='p15—p68',
             note='该科目官方表述为考查「阅读分析与文学写作的实际能力」，专业目录未列指定教材，'
                  '故本册不编造参考书；相关文言基础（句读、翻译、反切、活用）与《古代汉语》重合，'
                  '标注为参照范围，不当作指定书目。',
             books=[
                 dict(key='ydxz', name='本册无指定参考书：阅读与写作（实际能力考查）',
                      pub='—', edition='官方未列指定教材',
                      note='参照范围：黄德宽《古代汉语》的文选、句读、翻译与古注知识，'
                           '以及文学鉴赏与命题写作的实际表达能力。',
                      disc=None, sections=['阅读与写作', '阅读与写作(基础卷)']),
             ], source_note=ANHUI_NOTE, ev_kind='转载：同上（本科目官方未列指定教材）'),
    ]),
    ('黑龙江大学', [
        dict(paper='现代汉语（742）', years='2011—2014 合卷＋2015—2023＋2025—2026 回忆版',
             pages='p2—p56',
             note='2011—2014 为「语言学理论与现代汉语语音文字词汇」合卷，2025、2026 为'
                  '「现代汉语和古代汉语与语言学理论」回忆版合卷，均只取其中的现代汉语部分；'
                  '同一学科同年最多计一次。2022 年缺卷，2024 年现汉缺卷。'
                  '2020 年卷末「第二部分 现代汉语语音、文字、词汇（75 分）」v3 曾按未注年份不计，'
                  'v4 回原卷核实后补入。',
             books=[
                 dict(key='hlj-xdh', name='黄伯荣、廖序东《现代汉语》（增订第五版）',
                      pub='高等教育出版社', edition='2021 年当年书目标注「增订第五版」；'
                           '2026 年官方大纲已改为邵敬敏《现代汉语通论》第 3 版',
                      disc='现汉', sections=['现代汉语 绪论与语音', '现代汉语 文字与词汇', '现代汉语 语法与修辞',
                                       '现代汉语 补充(2020古理卷附段)']),
             ], source_note=HLJU_NOTE, ev_kind='官方：2026 年 742 大纲原件；2021 年大纲为权威转载'),
        dict(paper='古代汉语与语言学理论（842）', years='2011—2014 合卷＋2015—2023＋2024—2026 回忆版',
             pages='p4—p56',
             note='2011—2014 为「古代汉语与现代汉语语法修辞」合卷，2024—2026 为回忆版合卷，'
                  '均只取古代汉语、语言学理论部分；同一学科同年最多计一次。'
                  '2019 年卷末 pid774「主语的语义类型」v3 曾按疑似现汉不计，v4 回原卷核实为语理部分。',
             books=[
                 dict(key='hlj-gdh', name='王力主编《古代漢語》（一、二册）',
                      pub='中华书局', edition='校订重排本；2021 年当年书单未标出版年份',
                      disc='古汉', sections=['古代汉语']),
                 dict(key='hlj-yyll', name='《语言学纲要》（修订版），叶蜚声、徐通锵著，王洪君、李娟修订',
                      pub='北京大学出版社', edition='修订版；2026 年官方大纲仍为此本',
                      disc='语理', sections=['语言学理论']),
             ], source_note=HLJU_NOTE, ev_kind='官方：2026 年 842 大纲原件；2021 年大纲为权威转载'),
    ]),
    ('云南民族大学', [
        dict(paper='现代汉语（719）', years='2013—2025（11 年）', pages='p3—p28',
             note='2013、2014 年现汉卷题干未标科目代码，按同一科目序列归入本册。',
             books=[
                 dict(key='yn-xdh', name='黄伯荣、廖序东《现代汉语》（增订第六版）上、下册',
                      pub='高等教育出版社', edition='2017 年版；2025、2026 官方书目一致未变',
                      disc='现汉', sections=['现代汉语 绪论与语音', '现代汉语 文字与词汇', '现代汉语 语法与修辞',
                                       '现代汉语 修辞', '现代汉语 综合概念']),
             ], source_note=YNU_NOTE, ev_kind='官方：2025、2026 年参考书目 PDF 原件'),
        dict(paper='语言学理论（815）', years='2011—2025（15 年）', pages='p1—p29',
             note='2026 年官方书目已把 815 改为《语言学纲要》2025 年版，属 2025 年之后的范围变动，'
                  '不能用来评价既有真题；本册真题止于 2025 年，按 2010 年版《语言学纲要》评价。',
             books=[
                 dict(key='yn-yyll', name='《语言学纲要》（修订本），叶蜚声、徐通锵著，王洪君、李娟修订',
                      pub='北京大学出版社', edition='2010 年版（2026 年起改为 2025 年版）',
                      disc='语理', sections=['语言学理论 语音词汇与语法', '语言学理论 语言本质与系统',
                                       '语言学理论 演变接触与文字', '语言学理论 语法与词义',
                                       '语言学理论 语言学知识']),
             ], source_note=YNU_NOTE, ev_kind='官方：2025、2026 年参考书目 PDF 原件'),
    ]),
    ('西安工业大学', [
        dict(paper='语言综合（902／816）', years='2020—2025（6 年）', pages='p1—p50',
             note='科目代码随年份变化（2024 官方为 902，2025、2026 为 816），书目连续；'
                  '古代汉语指定版本 2024 与 2025 不同，本册按年份分别标注。',
             books=[
                 dict(key='xat-xdh', name='黄伯荣、李炜《现代汉语》（第二版）',
                      pub='北京大学出版社', edition='2016 年；2024、2025、2026 三处书目一致。'
                           '注意：不是黄伯荣、廖序东的同名教材',
                      disc='语综', sections=['现代汉语 绪论与语音', '现代汉语 文字与词汇', '现代汉语 语法与修辞']),
                 dict(key='xat-gdh', name='郭锡良《古代汉语》',
                      pub='商务印书馆', edition='2024 指定 2018 年本；2025、2026 为 1999 年修订本',
                      disc='语综', sections=['古代汉语 文字词义与句法', '古代汉语 阅读与诗律']),
             ], source_note=XATU_NOTE, ev_kind='2024、2026 官方书目页；2025 权威转载'),
        dict(paper='文学评论写作（702）', years='2020—2026（7 年）', pages='p4—p60',
             note='2020—2026 三部分大题均有选项（第一大题古典诗与现代诗任选），'
                  '题型连续出现不等于某位作家或某首作品高频。',
             books=[
                 dict(key='xat-wxll', name='童庆炳《文学理论教程》（第五版）',
                      pub='高等教育出版社', edition='2015 年版；2024、2025、2026 三处书目一致',
                      disc='评论', sections=['文学评论写作']),
             ], source_note=XATU_NOTE, ev_kind='2024、2026 官方书目页；2025 权威转载'),
    ]),
])

# ------------------------------------------------- 原卷补证（v3 曾排除，v4 补入）
def raw_item(school, pid):
    for r in RAW[school]:
        if r['pid'] == pid:
            return r
    raise KeyError((school, pid))


def add_topic(topics, discipline, section, title, points, kind, ev):
    for t in topics:
        if t['title'] == title and t['discipline'] == discipline:
            t['evidence'].extend(ev)
            t['evidence'].sort(key=lambda x: (x['year'], x['pid']))
            t['frequency'] = len({x['year'] for x in t['evidence']})
            t['years'] = sorted({x['year'] for x in t['evidence']})
            t['recent_hits'] = [y for y in t.get('recent_years', []) if y in set(t['years'])]
            return 'merged'
    topics.append(dict(section=section, title=title, points=points, kind=kind,
                       evidence=list(ev), discipline=discipline,
                       frequency=len({x['year'] for x in ev}),
                       years=sorted({x['year'] for x in ev}),
                       recent_years=[], recent_hits=[]))
    return 'added'


def ev_of(school, pid, title, section, discipline, kind, excerpt, pages, recall=False, precision='题面文字定位'):
    r = raw_item(school, pid)
    return dict(exam=r['exam'], eid=r['eid'], pid=pid, text=r['text'], type=kind, year=2020,
                subject='古代汉语与语言学理论', pages=pages, recall=recall, excerpt=excerpt,
                discipline=discipline, exam_pages=pages, page_precision=precision)


if __name__ == '__main__':
    # ===== 补证 1：黑龙江 2020 年卷末现代汉语段（打印版 p42，pid847—866）=====
    hl = SRC['黑龙江大学']
    add = [
        ('现代汉语 补充(2020古理卷附段)', '开口呼与非音质音位',
         '掌握开口呼、齐齿呼、合口呼、撮口呼的划分与非音质音位的类型（调位、重位、时位）。', 849,
         '开口呼 非音质音位（2020 年卷末现汉段术语题）'),
        ('现代汉语 补充(2020古理卷附段)', '语音四要素与声调轻声',
         '说明语音的物理四要素（音高、音强、音长、音色），并指出汉语声调主要与音高、轻声与音强（及音长）相关。', 851,
         '语音的物理四要素与声调、轻声的关系'),
        ('现代汉语 补充(2020古理卷附段)', '大写字母与隔音符号用法',
         '掌握《汉语拼音方案》中大写字母的使用场合与隔音符号（’）的书写规则。', 852,
         '大写字母和隔音符号的用法'),
        ('现代汉语 补充(2020古理卷附段)', '汉字的特点',
         '从表意体系、形体与音义关系、记录汉语的适应性等方面说明汉字特点。', 853,
         '简述汉字的特点'),
        ('现代汉语 补充(2020古理卷附段)', '多义词与同音同形词的区别',
         '区分多义词（义项间有联系）与同音同形词（意义无联系），能举例说明。', 854,
         '举例说明多义词与同音同形词的区别'),
        ('现代汉语 补充(2020古理卷附段)', '声母韵母发音描写与国际音标',
         '能描写给定声母、韵母的发音部位与发音方法，并写出对应国际音标。', 856,
         '描写声母或韵母的发音状况并注国际音标（10 分）'),
        ('现代汉语 补充(2020古理卷附段)', '音位变体及出现条件',
         '能分析/a/、/i/等音位的变体类型及其出现的语音条件。', 858,
         '分析下列音位的变体及出现条件（10 分）'),
        ('现代汉语 补充(2020古理卷附段)', '造字方式分析',
         '辨认给定汉字的造字方式（象形、指事、会意、形声），并说明判定依据。', 860,
         '分析下列汉字的造字方式（5 分）'),
        ('现代汉语 补充(2020古理卷附段)', '合成词构词方式分析',
         '分析合成词的构词方式；复合式还须指明联合、偏正、动宾、补充、主谓等结构类型。', 862,
         '分析合成词的构词方式及复合式结构类型（10 分）'),
        ('现代汉语 补充(2020古理卷附段)', '外来词类型',
         '按音译、音译兼意译、音译加意译、借形等类型判定外来词。', 865,
         '分析下列外来词的类型（5 分）'),
    ]
    for sec, title, points, pid, exc in add:
        e = ev_of('黑龙江大学', pid, title, sec, '现汉', '简答', exc, '42')
        print(' 补入现汉段:', title, add_topic(hl, '现汉', sec, title, points, '考点', [e]))

    # ===== 补证 2：黑龙江 2019 年古理卷末 pid774（打印版 p38，语理 15 分）=====
    r = raw_item('黑龙江大学', 774)
    e774 = dict(exam=r['exam'], eid=r['eid'], pid=774, text=r['text'], type='论述', year=2019,
                subject='古代汉语与语言学理论', pages='38', recall=False,
                excerpt='结合实例说明主语的语义类型（15 分）', discipline='语理',
                exam_pages='37-38', page_precision='题面文字定位')
    print(' 补入 2019 尾题:', add_topic(hl, '语理', '语言学理论', '主语的语义类型',
                                        '掌握主语与谓语的意义关系类型（施事、受事、工具、处所、时间等），能用实例说明。',
                                        '考点', [e774]))

    # ===== 补证 3：2020 年古汉「六书」论述（pid837，与现汉六书条目合并取证）=====
    r = raw_item('黑龙江大学', 837)
    e837 = dict(exam=r['exam'], eid=r['eid'], pid=837, text=r['text'], type='论述', year=2020,
                subject='古代汉语与语言学理论', pages='41', recall=False,
                excerpt='对文字学中「六书」的内容进行展开说明（12 分）', discipline='古汉',
                exam_pages='41', page_precision='题面文字定位')
    print(' 补入 2020 古汉六书:', add_topic(hl, '古汉', '古代汉语', '六书与造字法分析',
                                            '掌握六书名称与许慎定义、四体二用，并能辨析具体汉字的构造。',
                                            '考点', [e837]))

    # ===== 重算统计 =====
    for school, topics in SRC.items():
        for t in topics:
            t['evidence'].sort(key=lambda e: (e['year'], e['pid']))
            t['frequency'] = len({e['year'] for e in t['evidence']})
            t['years'] = sorted({e['year'] for e in t['evidence']})
            t['recent_hits'] = [y for y in t.get('recent_years', []) if y in set(t['years'])]

    # ===== 按科目、参考书重组 =====
    # 每个原始条目必须且只能归入一个科目册，防止同一主题在两册重复计算
    OWNER = {
        '安徽大学': {'基础': '语言文学基础（616）', '读写': '阅读与写作（809）'},
        # 例外：2004 年基础卷内的诗歌鉴赏写作题，按题型归入读写册
        '黑龙江大学': {'现汉': '现代汉语（742）', '古汉': '古代汉语与语言学理论（842）',
                   '语理': '古代汉语与语言学理论（842）'},
        '云南民族大学': {'现汉': '现代汉语（719）', '语理': '语言学理论（815）'},
        '西安工业大学': {'语综': '语言综合（902／816）', '评论': '文学评论写作（702）'},
    }
    out = OrderedDict()
    stats = OrderedDict()
    for school, papers in PAPERS.items():
        topics = SRC[school]
        out[school] = []
        assigned = {}
        for spec in papers:
            books = []
            for b in spec['books']:
                secs = b['sections']
                disc = b.get('disc')
                sel = [t for t in topics if t['section'] in secs
                       and (disc is None or t['discipline'] == disc)]
                for t in sel:
                    assigned[(t['discipline'], t['section'], t['title'])] = spec['paper']
                order = (lambda s: secs.index(s) if s in secs else 99)
                sel.sort(key=lambda t: (order(t['section']), -t['frequency'], t['title']))
                books.append(dict(key=b['key'], name=b['name'], pub=b['pub'], edition=b['edition'],
                                  note=b.get('note'), topics=sel, local_max=len(sel)))
            ev_all = [e for b in books for t in b['topics'] for e in t['evidence']]
            out[school].append(dict(paper=spec['paper'], years=spec['years'], pages=spec['pages'],
                                    note=spec['note'], source_note=spec['source_note'],
                                    ev_kind=spec.get('ev_kind'), books=books,
                                    stats=dict(topics=sum(len(b['topics']) for b in books),
                                               evidence=len(ev_all),
                                               recall=sum(1 for e in ev_all if e.get('recall')))))
            stats[f'{school}｜{spec["paper"]}'] = out[school][-1]['stats']
        # 校验：原始条目全部有归属，且归属与 OWNER 一致
        missing = [t for t in topics if (t['discipline'], t['section'], t['title']) not in assigned]
        EXC = {('安徽大学', '基础', '阅读与写作', '诗歌鉴赏写作')}
        wrong = [(t['title'], assigned[(t['discipline'], t['section'], t['title'])])
                 for t in topics
                 if (t['discipline'], t['section'], t['title']) in assigned
                 and (school, t['discipline'], t['section'], t['title']) not in EXC
                 and assigned[(t['discipline'], t['section'], t['title'])]
                 != OWNER[school].get(t['discipline'])]
        if missing:
            print(f'  !! {school} 未归册条目 {len(missing)}: '
                  f'{[(t["discipline"], t["section"], t["title"]) for t in missing]}')
        if wrong:
            print(f'  !! {school} 归册与 OWNER 不一致: {wrong}')
    (W / 'v4考点数据.json').write_text(json.dumps(out, ensure_ascii=False, indent=1))
    print('===== v4 分册统计 =====')
    for k, v in stats.items():
        print(f'  {k}: 条目 {v["topics"]}  证据 {v["evidence"]}  回忆版 {v["recall"]}')
    print(f'  合计条目 {sum(v["topics"] for v in stats.values())}'
          f'  证据 {sum(v["evidence"] for v in stats.values())}')
