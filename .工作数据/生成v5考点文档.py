# -*- coding: utf-8 -*-
"""生成 v5 四校高频考点册（一校一册；册内按「考试科目＋科目代码＋参考书」分段）。

输入：.工作数据/v5考点数据.json（准备v5数据.py → 补正v5证据.py）
输出：v5/<学校>_高频考点_v5.docx → 再由 render_pdf.py 转 PDF

与 v4 的差别：
1. 一个学校一个 PDF 册；科目与参考书在册内分段，不再一科目一个文件。
2. 每段标注科目名称、科目代码（卷面印刷口径）、年份区间、卷面页码与本书目。
3. 参考书按年份分列（如安徽古代汉语：2007—2018 王力本、2019—2021 黄德宽本），
   不再用一年的书目覆盖全部年份。
4. 同一考点若在不同分段都有证据，则在各段分别列出，只带本段证据，并注明该考点的
   全部年次，避免读者把分段计数当成缩水。
"""
import json
from collections import defaultdict
from pathlib import Path
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.enum.table import WD_TABLE_ALIGNMENT

W = Path(__file__).resolve().parent
O = Path('/home/zyf/考研/v5')
D = json.loads((W / 'v5考点数据.json').read_text())
VERSION = 'v5'
DATE = '2026 年 10 月 12 日'
FONT = 'Noto Serif CJK SC'

# 黑龙江各年试卷简称（沿用 v3／v4 习惯写法）
ABBR = {'语言学理论与现代汉语语音文字词汇': '理现', '古代汉语与现代汉语语法修辞': '古现',
        '现代汉语': '现汉', '古代汉语与语言学理论': '古理', '语言学理论': '语理',
        '现代汉语和古代汉语与语言学理论 回忆版': '合卷', '古代汉语与语言学理论 回忆版': '古理'}
SHOW = set(ABBR) - {'现代汉语', '古代汉语与语言学理论', '语言学理论'}

NOTES = {
    '安徽大学': '本册按年份把「语言文学基础」分成四段，是因为科目代码与参考书都换过：'
                '2004—2011 年该科书目未查到，只呈现真题；2012—2016 年现代汉语为胡裕树本、'
                '古代汉语为王力本；2017—2018 年科目代码改为 615，现代汉语换为黄廖增订五版；'
                '2019—2021 年改为 616／809，古代汉语换为黄德宽本、现代汉语 2020 年起为增订六版。'
                '同一考点若在不同年份段的书里都出现，会在对应段各列一次，各段只带本段证据。',
    '黑龙江大学': '842 是一份卷子里的两部分，按参考书分成「古代汉语」与「语言学理论」两段：'
                  '前者只用王力本，后者只用《语言学纲要》。同一学科同年最多计一次；'
                  '2022 年缺卷、2024 年现汉缺卷，2024 年古理卷打印复核版只有卷头、无题目内容。',
    '云南民族大学': '现代汉语与语言学理论分别计年次，同一知识点可在两科分别出现。'
                    '早期（2011、2013 等）及若干近年资料为回忆版，不完整；'
                    '数字只表示可确认的出现，不能据未见题目推断未考。',
    '西安工业大学': '语言综合按科目代码分为 902（2020—2024）与 816（2025）两段：'
                    '两段参考书相同、古代汉语版本不同。文学评论写作单独成段。'
                    '评论第一大题中古典诗与现代诗是任选关系，合记为一个鉴赏题型；'
                    '题型连续出现不等于某位作家或某首作品高频。',
}

# 段内条目较少、需要提醒「小段照做」的说明
SEG_HINT = {
    'xat-816': '本段只有 2025 一年。按整卷照做，不并入 902 段：'
               '该年科目代码已改为 816，古代汉语参考书版本也与 902 段不同。',
    'ah-jc-none': '本段各年的初试参考书目未取得，书目栏如实留空。'
                  '本段考点按知识块编排，出处给到卷面页码与题号，可直接回原卷核对。',
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
    return p(d, text, size, True, True, 4, style='Heading 2')


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


def seg_years(seg):
    return sorted({e['year'] for b in seg['books'] for t in b['topics'] for e in t['evidence']})


def recent_years_for(topic, seg):
    """近五年＝本段现有最新年份及此前四个自然年，仅列本段现存年份。"""
    ys = seg_years(seg)
    if not ys:
        return []
    latest = ys[-1]
    return [y for y in ys if latest - 4 <= y <= latest]


def recall_count(books):
    ev = [e for b in books for t in b['topics'] for e in t['evidence']]
    return sum(1 for e in ev if e.get('recall')), len(ev)


def example(d, head, req, frame, extra, judge, after=6):
    """例题块：题目要求 → 参考课程框架 → 额外加工 → 难度判断（与 v3／v4 总结样式一致）。"""
    q = p(d, '', keep=True, after=2)
    font(q.add_run(head), 10, True, 'A22F2F')
    for lab, txt in [('题目要求', req), ('参考课程框架', frame),
                     ('在教材基础上的额外加工', extra), ('难度判断', judge)]:
        rich(d, [('　' + lab + '：', True, '555555', 9), (txt, False, '000000', 9.5)],
             keep=(lab != '难度判断'), after=2, indent=0.4)
    d.paragraphs[-1].paragraph_format.space_after = Pt(after)



# --------------------------------------------------------------------- 生成
def build(school, blob):
    segs = blob['segments']
    title = f'{school} 高频考点'
    d = setup(title)
    uq = blob['stats']['unique_topics']
    pages_ = sum(s['stats']['evidence'] for s in segs)
    rich(d, [(school, False, '000000', 11),
             (f'　考试科目 {len(segs)} 段', True, '000000', 11),
             (f'　唯一考点 {uq} 条', False, '555555', 10)], after=2)
    rich(d, [('依据《' + school + '历年真题打印复核版》逐题归类；'
              '与《四校真题考点与难度对比》配套使用', False, '555555', 9)], after=2)
    rich(d, [(f'{DATE}　{VERSION}（一校一册；册内按考试科目＋科目代码＋参考书分段）',
              False, '555555', 9)], after=8)

    # ---------- 一、分册说明 ----------
    band(d, '一、本册分段与使用口径')
    rows = []
    for i, seg in enumerate(segs, 1):
        rows.append([str(i), seg['name'], seg['code'], seg['years'],
                     seg['page_real'] or seg['page'],
                     f'{seg["stats"]["topics"]} 条／{seg["stats"]["evidence"]} 证据'])
    rows.append(['合计', f'{len(segs)} 段', '—', '—', '—',
                 f'段内合计 {blob["stats"]["instances"]} 条／{blob["stats"]["evidence"]} 证据'])
    table(d, ['段', '考试科目', '科目代码（卷面）', '本段年份', '卷面页码', '条目／证据'],
          rows, [0.9, 3.3, 3.5, 2.3, 2.5, 4.5], 8.2)
    gap(d, 3)
    rich(d, [('口径说明：', True),
             (f'本册唯一考点 {uq} 条（按考点去重）；段内条目合计 '
              f'{blob["stats"]["instances"]} 条、证据 {blob["stats"]["evidence"]} 条，'
              '因为同一考点若在不同年份段的参考书里都出现，会在各段分别列出一次，'
              '各段只带本段证据。两套数字都对：前者是「考了多少个考点」，后者是'
              '「各段各有多少条要背」。', False, '000000', 9)], after=3)
    rich(d, [(NOTES[school], False, '555555', 9)], after=8)

    # ---------- 二、参考书目（按年份段分列）----------
    band(d, '二、参考书目（按年份段分列）')
    p(d, '下表逐段列出该段真题对应的官方参考书目。同一门课在不同年份换过教材的，'
         '按年份段分别列出，不合并成一条；未查到官方书目的年份段如实留空。', 9.5, after=5)
    rows = []
    for i, seg in enumerate(segs, 1):
        label = f'{seg["label"]}\n{seg["years"]}'
        for j, b in enumerate(seg['books'], 1):
            rows.append([label if j == 1 else '', b['name'],
                         b.get('pub') if b.get('pub') not in ('—', '-', '') else '—',
                         b['edition'], f'{b["local_max"]} 条'])
    table(d, ['年份段', '书名与作者', '出版社', '版本依据', '本段条目'], rows,
          [3.2, 5.4, 2.2, 4.2, 1.6], 8.2)
    gap(d, 4)
    seen_note = set()
    for seg in segs:
        for b in seg['books']:
            if b.get('note') and b['name'] not in seen_note:
                seen_note.add(b['name'])
                rich(d, [('· ', True, '555555', 9), (b['name'][:24] + '：', True, '555555', 9),
                         (b['note'], False, '000000', 9)], after=2, indent=0.4)
    # 证据级别按内容去重：同一级别覆盖多个段时只写一次，段名合并列出
    groups = []
    for seg in segs:
        for g in groups:
            if g['note'] == seg['source_note']:
                g['labels'].append(seg['label'])
                break
        else:
            groups.append(dict(note=seg['source_note'], labels=[seg['label']]))
    for g in groups:
        rich(d, [('证据级别｜' + '／'.join(g['labels']) + '：', True, '555555', 9),
                 (g['note'], False, '000000', 9)], after=3, indent=0.4)
    gap(d, 5)

    # ---------- 三、考点正文（按段落 → 参考书分块）----------
    band(d, '三、高频考点（按科目段 → 参考书分块编号）')
    for si, seg in enumerate(segs, 1):
        sub(d, f'第 {si} 段　{seg["label"]}　科目代码 {seg["code"]}　'
               f'{seg["years"]}（{seg["stats"]["topics"]} 条）', 12)
        rich(d, [(f'卷面页码：{seg["page_real"] or seg["page"]}（本校《历年真题打印复核版》'
                  f'页脚印刷页码；分段名义范围 {seg["page"]}）　'
                  f'证据：{seg["stats"]["evidence"]} 条，其中回忆版 {seg["stats"]["recall"]} 条',
                  False, '555555', 8.5)], after=3)
        rich(d, [(seg['note'], False, '000000', 9)], after=3)
        if SEG_HINT.get(seg['key']):
            rich(d, [('提醒：', True, 'A22F2F', 9),
                     (SEG_HINT[seg['key']], False, '000000', 9)], after=4)
        for bi, b in enumerate(seg['books'], 1):
            q = p(d, f'参考书 {si}-{bi}：{b["name"]}', 11, True, True, 3)
            q.paragraph_format.space_before = Pt(5)
            z = OxmlElement('w:shd'); z.set(qn('w:fill'), 'E8EFF7')
            q._p.get_or_add_pPr().append(z)
            if b.get('note'):
                rich(d, [(b['note'], False, '555555', 8.5)], after=3)
            rich(d, [(f'出版社：{b.get("pub") or "—"}　版本：{b["edition"]}　'
                      f'本块条目：{b["local_max"]} 条', False, '555555', 8.5)], after=5)
            cur_sec = None
            for idx, t in enumerate(b['topics'], 1):
                if t['section'] != cur_sec:
                    cur_sec = t['section']
                    qq = p(d, cur_sec, 9.5, True, True, 3)
                    qq.paragraph_format.space_before = Pt(4)
                    z = OxmlElement('w:shd'); z.set(qn('w:fill'), 'EDEDED')
                    qq._p.get_or_add_pPr().append(z)
                q = p(d, '', keep=True, after=3, style='Heading 2')
                font(q.add_run(f'{idx}. {t["title"]}'), 11, True)
                extra = ''
                if t.get('all_frequency') and t['all_frequency'] > t['frequency']:
                    extra = f'（本科目全部年次 {t["all_frequency"]}，本段 {t["frequency"]}）'
                font(q.add_run('　' + t['kind'] + '本段已确认 ' + str(t['frequency'])
                               + ' 年次' + extra), 9, True)
                types = '／'.join(sorted({e['type'] for e in t['evidence'] if e.get('type')}))
                rich(d, [(f'【{types}】', True, 'A22F2F', 9.5),
                         ('掌握：' + t['points'], False, '000000', 9.5)], keep=True, after=3)
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
                rys = recent_years_for(t, seg)
                hits = [y for y in t['years'] if y in rys]
                if not rys:
                    line = '近年：本段无现存年份记录。'
                elif not hits:
                    line = '近年现存 ' + '、'.join(str(y) for y in rys) + '；其中暂无确认记录。'
                else:
                    line = ('近年现存 ' + '、'.join(str(y) for y in rys) + '；已见 ' +
                            '、'.join(str(y) + ('*' if y in rec else '') for y in hits) + '。')
                p(d, line, 8.5, color='555555', after=2, keep=True)
                last = p(d, '本段全部出处：' + evidence_line(t, school) + '。',
                         8.5, color='555555', after=7)
                last.paragraph_format.keep_together = True
        gap(d, 6)

    # ---------- 四、资料边界 ----------
    band(d, '四、资料边界与未核事项')
    rc, tt = recall_count([b for s in segs for b in s['books']])
    for txt in [
        '条目来自目前可读题面、原图补证与人工归类，未对缺失题作补齐；'
        '页码与原题异常文字以打印复核版为准。',
        f'本册 {rc}／{tt} 条段内证据取自回忆版（年份后带 *），可靠度低于完整卷；'
        '回忆版未记录到的题不等于没有考。',
        '黑龙江 2024 年古代汉语与语言学理论卷在打印复核版里只有卷头（「2024 年／'
        '专业课（二）／古汉」），没有题目内容，故本册对该卷不列任何出处。',
        '「年次」是已确认证据的下限，不换算命中率，也不作为命题预测概率。',
        '主题包含考点、专题与题型，不是互斥的原子考点；同一道题可支撑不同主题，'
        '证据条数不等于独立题数。',
        '参考书目为对应年份官方或权威转载公布的初试范围，完整教材未逐页对照，'
        '因此本册不下「全部题目均在书内」的结论；2004—2011 年安徽该科书目未查到，'
        '已在本册如实留空。',
    ]:
        rich(d, [('· ', True), (txt,)], 9, after=2, indent=0.4)
    return d, title


if __name__ == '__main__':
    O.mkdir(parents=True, exist_ok=True)
    manifest = []
    for school, blob in D.items():
        d, title = build(school, blob)
        fn = f'{school}_高频考点_{VERSION}.docx'
        path = O / fn
        d.save(path)
        manifest.append(dict(file=fn, school=school,
                             segments=[dict(key=s['key'], label=s['label'], years=s['years'],
                                            code=s['code'], topics=s['stats']['topics'],
                                            evidence=s['stats']['evidence'],
                                            recall=s['stats']['recall'],
                                            books=[b['name'] for b in s['books']])
                                       for s in blob['segments']],
                             unique_topics=blob['stats']['unique_topics'],
                             instances=blob['stats']['instances'],
                             evidence=blob['stats']['evidence'],
                             unique_evidence=blob['stats']['unique_evidence'],
                             recall=blob['stats']['recall']))
        print(f'{fn}  段 {len(blob["segments"])}  唯一考点 {blob["stats"]["unique_topics"]}  '
              f'段内条目 {blob["stats"]["instances"]}  证据 {blob["stats"]["evidence"]}')
    (W / f'{VERSION}册清单.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=1))
    print('册数', len(manifest))
