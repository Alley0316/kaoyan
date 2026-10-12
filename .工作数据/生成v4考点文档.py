# -*- coding: utf-8 -*-
"""生成 v4 四校高频考点册（一个 PDF 册 = 一个考试科目；正文按参考书分块，编号各自从 1 开始）。

输入：.工作数据/v4考点数据.json（由 准备v4数据.py 生成）
输出：v4/<学校>_<科目>_高频考点_v4.docx → 再由 render_pdf.py 转 PDF

与 v3 的差别：参考书目置于正文最前；全部条目（含仅确认一次者）按参考书分块连续编号；
v3 曾排除的黑龙江 2020 年现汉附段、2019 年尾题已补入。
"""
import json
from pathlib import Path
from collections import defaultdict
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.enum.table import WD_TABLE_ALIGNMENT

W = Path(__file__).resolve().parent
O = Path('/home/zyf/考研/v4')
D = json.loads((W / 'v4考点数据.json').read_text())
VERSION = 'v4'
DATE = '2026 年 10 月 12 日'
FONT = 'Noto Serif CJK SC'

# 黑龙江各年试卷简称（沿用 v3 习惯写法）
ABBR = {'语言学理论与现代汉语语音文字词汇': '理现', '古代汉语与现代汉语语法修辞': '古现',
        '现代汉语': '现汉', '古代汉语与语言学理论': '古理', '语言学理论': '语理',
        '现代汉语和古代汉语与语言学理论 回忆版': '合卷', '古代汉语与语言学理论 回忆版': '古理'}
SHOW = set(ABBR) - {'现代汉语', '古代汉语与语言学理论', '语言学理论'}

NOTES = {
    '安徽大学': '基础卷 2005 与 2007 题目高度重复，本册仍按两年分别保留出处，不合并计数。'
                '原始资料含填答与解析，填空题只摘要列出考点，避免把资料中的填答当作标准答案。'
                '现有真题止于 2021 年，评的是这批卷子对应的旧「语言文学基础」范围。',
    '黑龙江大学': '2011—2014 分卷方式与后期不同，本册按知识部分拆入现代汉语、古代汉语、语言学理论三块，'
                  '同一学科同年最多计一次。2020 年卷末现代汉语段、2019 年卷末语理题在本版已按原卷补入。'
                  '2022 年缺卷，2024 年现汉缺卷。',
    '云南民族大学': '现代汉语与语言学理论分别计年次，共有知识点可在两科分别出现。'
                    '早期（2011、2013 等）及若干近年资料为回忆版，不完整；'
                    '数字只表示可确认的出现，不能据未见题目推断未考。',
    '西安工业大学': '语言综合现有 2020—2025 年，文学评论写作现有 2020—2026 年。'
                    '评论第一大题中古典诗与现代诗是任选关系，合记为一个鉴赏题型；第二、三大题也有材料选择。'
                    '题型连续出现不等于某位作家或某首作品高频。',
}


# ----------------------------------------------------------------- 基础样式
def font(r, size=10, bold=False, color='000000'):
    r.font.name = FONT; r.font.size = Pt(size); r.bold = bold
    r.font.color.rgb = RGBColor.from_string(color)
    rf = r._element.get_or_add_rPr().find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts'); r._element.get_or_add_rPr().insert(0, rf)
    for k in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
        rf.set(qn('w:' + k), FONT)
    return r


def p(d, text='', size=10, bold=False, keep=False, after=4, color='000000', style=None, indent=0):
    x = d.add_paragraph(style=style); f = x.paragraph_format
    f.line_spacing = 1.15; f.space_before = Pt(0); f.space_after = Pt(after)
    f.keep_with_next = keep; f.widow_control = True
    if indent:
        f.left_indent = Cm(indent)
    font(x.add_run(text), size, bold, color)
    return x


def rich(d, chunks, size=10, keep=False, after=4, indent=0):
    """chunks: [(text, bold, color, size)]，缺省项可省略。"""
    x = d.add_paragraph(); f = x.paragraph_format
    f.line_spacing = 1.15; f.space_before = Pt(0); f.space_after = Pt(after)
    f.keep_with_next = keep; f.widow_control = True
    if indent:
        f.left_indent = Cm(indent)
    for c in chunks:
        t = c[0]; b = c[1] if len(c) > 1 else False
        col = c[2] if len(c) > 2 else '000000'; sz = c[3] if len(c) > 3 else size
        font(x.add_run(t), sz, b, col)
    return x


def gap(d, after=4):
    x = d.add_paragraph(); f = x.paragraph_format
    f.line_spacing = 1.0; f.space_before = Pt(0); f.space_after = Pt(after)
    font(x.add_run(''), 1)
    return x


def setup(title):
    d = Document(); s = d.sections[0]
    s.page_width = Cm(21); s.page_height = Cm(29.7)
    s.top_margin = Cm(1.7); s.bottom_margin = Cm(1.7)
    s.left_margin = Cm(1.8); s.right_margin = Cm(1.8)
    s.header_distance = Cm(.65); s.footer_distance = Cm(.7)
    for name in ['Normal', 'Title', 'Heading 1', 'Heading 2']:
        st = d.styles[name]; st.font.name = FONT; st.font.color.rgb = RGBColor(0, 0, 0)
        for b in list(st.element.iter(qn('w:pBdr'))):
            b.getparent().remove(b)
    d.core_properties.title = title
    d.core_properties.author = ''
    d.core_properties.subject = f'{DATE} 四校真题高频考点 {VERSION}'
    h = s.header.paragraphs[0]; h.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    font(h.add_run(title), 8)
    f = s.footer.paragraphs[0]; f.alignment = WD_ALIGN_PARAGRAPH.CENTER
    font(f.add_run('第 '), 9)
    z = OxmlElement('w:fldSimple'); z.set(qn('w:instr'), 'PAGE'); f._p.append(z)
    font(f.add_run(' 页'), 9)
    p(d, title, 19, True, after=6, style='Title')
    return d


def band(d, title, fill='F4E8BF'):
    q = p(d, title, 12, True, True, 7, style='Heading 1')
    q.paragraph_format.space_before = Pt(8)
    q.alignment = WD_ALIGN_PARAGRAPH.CENTER
    z = OxmlElement('w:shd'); z.set(qn('w:fill'), fill)
    q._p.get_or_add_pPr().append(z)
    return q


def sub(d, text, size=11):
    q = p(d, text, size, True, True, 4, style='Heading 2')
    return q


def _tr_pr(row):
    trPr = row._tr.find(qn('w:trPr'))
    if trPr is None:
        trPr = OxmlElement('w:trPr'); row._tr.insert(0, trPr)
    return trPr


def table(d, headers, rows, widths, size=8.5):
    t = d.add_table(rows=1, cols=len(headers))
    t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.autofit = False
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]; c.text = ''
        pr = c.paragraphs[0]; pr.paragraph_format.space_after = Pt(1)
        pr.paragraph_format.line_spacing = 1.05
        font(pr.add_run(h), size, True)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ''
            pr = cells[i].paragraphs[0]
            pr.paragraph_format.space_after = Pt(1); pr.paragraph_format.line_spacing = 1.05
            font(pr.add_run(v), size)
    for r in t.rows:
        for i, w in enumerate(widths):
            r.cells[i].width = Cm(w)
    for idx, r in enumerate(t.rows):
        trPr = _tr_pr(r); trPr.append(OxmlElement('w:cantSplit'))
        if idx == 0:
            trPr.append(OxmlElement('w:tblHeader'))
    return t


def example(d, head, req, frame, extra, judge, after=6):
    """例题块：题目要求 → 参考课程框架 → 额外加工 → 难度判断（与 v3 总结样式一致）。"""
    q = p(d, '', keep=True, after=2)
    font(q.add_run(head), 10, True, 'A22F2F')
    for lab, txt in [('题目要求', req), ('参考课程框架', frame),
                     ('在教材基础上的额外加工', extra), ('难度判断', judge)]:
        rich(d, [('　' + lab + '：', True, '555555', 9), (txt, False, '000000', 9.5)],
             keep=(lab != '难度判断'), after=2, indent=0.4)
    d.paragraphs[-1].paragraph_format.space_after = Pt(after)


# --------------------------------------------------------------- 证据行工具
def evidence_line(t, school):
    """年份＋页码（页脚印刷页码）；合并卷／回忆卷另标试卷简称，并给回忆版打 *。"""
    by = defaultdict(list)
    for e in t['evidence']:
        by[(e['year'], str(e['pages']))].append(e)
    out = []
    for (y, pg), es in sorted(by.items()):
        s = f'{y} '
        if school == '黑龙江大学' and es[0].get('subject') in SHOW:
            s += ABBR.get(es[0]['subject'], es[0].get('discipline', '')) + ' '
        s += 'p' + pg
        if any(e.get('recall') for e in es):
            s += '*'
        out.append(s)
    return '；'.join(out)


def paper_years(spec):
    ys = set()
    for b in spec['books']:
        for t in b['topics']:
            for e in t['evidence']:
                if e.get('year'):
                    ys.add(e['year'])
    return sorted(ys)


def recent_years_for(topic, spec):
    """近五年＝本科目现有最新年份及此前四个自然年，仅列该校本科目现存年份。"""
    ys = paper_years(spec)
    if not ys:
        return []
    latest = ys[-1]
    return [y for y in ys if latest - 4 <= y <= latest]


def recall_count(books):
    ev = [e for b in books for t in b['topics'] for e in t['evidence']]
    return sum(1 for e in ev if e.get('recall')), len(ev)


# --------------------------------------------------------------------- 生成
def build(school, spec):
    title = f'{school} {spec["paper"].split("（")[0]} 高频考点'
    d = setup(title)
    rich(d, [(school, False, '000000', 11), (f'　{spec["paper"]}', True, '000000', 11),
             (f'　{v4_label(spec)}', False, '555555', 10)], after=2)
    rich(d, [('依据《' + school + '历年真题打印版》逐题归类；与《四校真题考点与难度对比》配套使用',
              False, '555555', 9)], after=2)
    rich(d, [(f'{DATE}　{v4_label(spec)}', False, '555555', 9)], after=8)

    # ---------- 一、参考书目（置于正文最前）----------
    band(d, '一、参考书目（本科目）')
    p(d, '以下书目为对应招生年份官方（或权威转载）公布的初试参考书目。'
         '本科目所有考点均按这些书分块编排，编号在每本书内各自从 1 开始。', 9.5, after=5)
    rows = []
    for i, b in enumerate(spec['books'], 1):
        rows.append([str(i), b['name'], b['pub'], b['edition'],
                     f'{b["local_max"]} 条'])
    table(d, ['序', '书名与作者', '出版社', '版本依据', '本册条目数'], rows,
          [1.0, 6.4, 2.6, 4.6, 2.4], 8.5)
    gap(d, 4)
    for i, b in enumerate(spec['books'], 1):
        if b.get('note'):
            rich(d, [(f'{i}. ', True, '555555', 9), (b['note'], False, '000000', 9)], after=2, indent=0.4)
    rich(d, [('证据级别：', True, '555555', 9), (spec['source_note'], False, '000000', 9)], after=8)

    # ---------- 二、真题范围与统计 ----------
    band(d, '二、真题范围与统计口径')
    table(d, ['项目', '内容'],
          [['本科目真题范围', f'{spec["years"]}'],
           ['对应页码', f'{spec["pages"]}（本校《历年真题打印版》页脚印刷页码）'],
           ['本册条目', f'{spec["stats"]["topics"]} 条（含只确认一次的条目）'],
           ['证据记录', f'{spec["stats"]["evidence"]} 条，其中回忆版 '
                    f'{spec["stats"]["recall"]} 条'],
           ['计数单位', '同一学科的不同年份；同年多题只计一次，故「年次」为已确认证据的下限'],
           ['编号规则', '按参考书分块，每本书内编号从 1 连续编号；全书不连续累计']],
          [3.4, 13.6], 9)
    gap(d, 4)
    rich(d, [('使用说明：', True),
             ('米色为参考书分块标题；红字为题型；黄色为真题题目摘录。每条含掌握要求、'
              '近年现存年份与出现情况、全部已确认出处。', False, '000000', 9)], after=3)
    rich(d, [(NOTES[school], False, '555555', 9)], after=8)

    # ---------- 三、考点正文（按参考书分块）----------
    band(d, '三、高频考点（按参考书分块编号）')
    for bi, b in enumerate(spec['books'], 1):
        sub(d, f'参考书 {bi}：{b["name"]}', 11.5)
        if b.get('note'):
            rich(d, [(b['note'], False, '555555', 8.5)], after=3)
        rich(d, [(f'出版社：{b["pub"]}　版本：{b["edition"]}　本块条目：{b["local_max"]} 条',
                  False, '555555', 8.5)], after=5)
        cur_sec = None
        for idx, t in enumerate(b['topics'], 1):
            if t['section'] != cur_sec:
                cur_sec = t['section']
                q = p(d, cur_sec, 9.5, True, True, 3)
                q.paragraph_format.space_before = Pt(4)
                z = OxmlElement('w:shd'); z.set(qn('w:fill'), 'EDEDED')
                q._p.get_or_add_pPr().append(z)
            q = p(d, '', keep=True, after=3, style='Heading 2')
            font(q.add_run(f'{idx}. {t["title"]}'), 11, True)
            font(q.add_run('　' + t['kind'] + '已确认 ' + str(t['frequency']) + ' 年次'), 9, True)
            types = '／'.join(sorted({e['type'] for e in t['evidence'] if e.get('type')}))
            rich(d, [(f'【{types}】', True, 'A22F2F', 9.5), ('掌握：' + t['points'], False, '000000', 9.5)],
                 keep=True, after=3)
            ev = t['evidence']
            samples = [ev[0], ev[-1]] if len(ev) > 1 else [ev[0]]
            q = p(d, '', keep=True, after=3)
            for i, e in enumerate(samples):
                if i:
                    font(q.add_run('　'), 9)
                ex = (e.get('excerpt') or '').strip()
                tx = ex[:70] + ('…' if len(ex) > 70 else '')
                r = font(q.add_run(f'【{e["year"]} {tx}】'), 9, True)
                r.font.highlight_color = WD_COLOR_INDEX.YELLOW
            rec = {e['year'] for e in ev if e.get('recall')}
            rys = recent_years_for(t, spec)
            hits = [y for y in t['years'] if y in rys]
            if not rys:
                line = '近年：本科目无现存年份记录。'
            elif not hits:
                line = '近年现存 ' + '、'.join(str(y) for y in rys) + '；其中暂无确认记录。'
            else:
                line = ('近年现存 ' + '、'.join(str(y) for y in rys) + '；已见 ' +
                        '、'.join(str(y) + ('*' if y in rec else '') for y in hits) + '。')
            p(d, line, 8.5, color='555555', after=2, keep=True)
            last = p(d, '全部出处：' + evidence_line(t, school) + '。', 8.5, color='555555', after=7)
            last.paragraph_format.keep_together = True
    # ---------- 四、资料边界 ----------
    band(d, '四、资料边界与未核事项')
    rc, tt = recall_count(spec['books'])
    for txt in [
        '条目来自目前可读题面、原图补证与人工归类，未对缺失题作补齐；页码与原题异常文字以打印复核版为准。',
        f'本册 {rc}／{tt} 条证据取自回忆版（年份后带 *），可靠度低于完整卷；回忆版未记录到的题不等于没有考。',
        '「年次」是已确认证据的下限，不换算命中率，也不作为命题预测概率。',
        '主题包含考点、专题与题型，不是互斥的原子考点；同一道题可支撑不同主题，证据条数不等于独立题数。',
        '参考书目为对应年份官方或权威转载公布的初试范围，完整教材未逐页对照，'
        '因此本册不下「全部题目均在书内」的结论。',
    ]:
        rich(d, [('· ', True), (txt,)], 9, after=2, indent=0.4)
    return d, title


def v4_label(spec):
    return f'{VERSION}（全部条目按参考书分块编号，编号各自从 1 起）'


if __name__ == '__main__':
    O.mkdir(parents=True, exist_ok=True)
    manifest = []
    for school, specs in D.items():
        for spec in specs:
            d, title = build(school, spec)
            fn = f'{school}_{spec["paper"].split("（")[0]}_高频考点_v4.docx'
            path = O / fn
            d.save(path)
            manifest.append(dict(file=fn, school=school, paper=spec['paper'],
                                  topics=spec['stats']['topics'], evidence=spec['stats']['evidence'],
                                  recall=spec['stats']['recall'],
                                  books=[b['name'] for b in spec['books']],
                                  local_max=[b['local_max'] for b in spec['books']]))
            print(f'{fn}  条目 {spec["stats"]["topics"]}  证据 {spec["stats"]["evidence"]}')
    (W / 'v4册清单.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=1))
    print('册数', len(manifest), '条目合计', sum(m['topics'] for m in manifest))
