# -*- coding: utf-8 -*-
"""补正 v5 人工补录证据的卷别与年份（只改卷别／年份／页码，不改条目本身）。

背景：v3／v4 阶段为补足回忆版缺段，往数据里补录了 19 条「题目要求」层级的证据，
pid 用 1201+／1301+／1401+ 的占位号，年份按回忆印象填写。这些占位证据没有逐行
题面可回溯，经与打印复核版逐卷比对，发现卷别或年份填错，需按「证据＝卷别＋题号＋原文」
的口径补正；无法在现有原卷中定位的，一律移出考点正文并记入补正清单。

判据：把补录摘要与该卷原文做最长公共子串比对；子串 ≥6 字视为该卷确有此题。
补正动作分三类：
  D 删除   该卷已有同句真实题面（补录属重复），或整卷无题面内容（回忆版只有卷头）。
  A 改卷别 摘要能在另一年卷中找到同句，改到该卷并把页码／年份改为该卷口径。
  F 存疑   多年份都有同句（无法判定是某一年），移出正文并记入补正清单待人工复核。

输入：.工作数据/v5考点数据.json（prepare 脚本产物）
输出：.工作数据/v5考点数据.json（原地补正）、.工作数据/v5证据补正清单.json
"""
import json
import re
import sys
from collections import OrderedDict
from pathlib import Path

W = Path(__file__).resolve().parent
DATA_P = W / 'v5考点数据.json'
DATA = json.loads(DATA_P.read_text())
RAW = {s: json.loads((W.parent / '.高频考点工作' / f'{s}_历年真题_打印版.json').read_text())
       for s in DATA}
SYNTH = set(range(1201, 1210)) | set(range(1301, 1310)) | set(range(1401, 1410))

# 补正决定：pid → (动作, 目标卷, 目标页码, 说明)
FIX = {
    1401: ('A', '云南民族大学 2016年 现代汉语', '8', '2016 年卷二、分析题第 2 题的原题干'),
    1402: ('D', None, None, '2016 年卷已有同句（分析题第 3 题），补录属跨卷重复'),
    1403: ('A', '云南民族大学 2016年 现代汉语', '8', '2016 年卷分析题中的韵母辨析小题，原卷该小题题干未逐字印出'),
    1404: ('A', '云南民族大学 2016年 现代汉语', '8',
           '2016 年卷分析题第 5 题「指出下列复合式合成词的类型」被误记为造字法；'
           '2016 年卷分析题确含造字法内容，按分析题整体归 2016 年'),
    1405: ('D', None, None, '2016 年卷已有同句（分析题第 3 题），补录属跨卷重复'),
    1406: ('D', None, None, '2016 年卷已有同句（分析题第 4 题），补录属跨卷重复'),
    1407: ('D', None, None, '2016 年卷已有同句（分析题第 5 题），补录属跨卷重复'),
    1408: ('D', None, None, '2016 年卷已有同句（分析题第 6 题），补录属跨卷重复'),
    1409: ('D', None, None, '2020 年卷该题（分析题第（一）小题）已有真实题面，补录与之重复'),
    1301: ('F', None, None, '「指出下列汉字的造字方法」在 2021—2025 各年卷均有同句，无法判定补录指向哪一年'),
    1302: ('F', None, None, '「指出下列合成词的结构类型」在 2021—2025 各年卷均有同句，无法判定年份'),
    1303: ('F', None, None, '同 1302，且 2024 年卷已有真实题面'),
    1304: ('F', None, None, '「指出下列句中活用词的用法及意义」2021、2024 年卷均有，无法判定年份'),
    1201: ('D', None, None, '2024 年卷打印复核版只有卷头（「2024 年／专业课（二）／古汉」），无题目内容可定位'),
    1202: ('D', None, None, '2024 年卷无题目内容可定位'),
    1203: ('D', None, None, '2024 年卷无题目内容可定位'),
    1204: ('D', None, None, '2024 年卷无题目内容可定位'),
    1205: ('D', None, None, '2024 年卷无题目内容可定位'),
    1206: ('D', None, None, '2024 年卷无题目内容可定位'),
}


def pool(exam):
    return ''.join(re.sub(r'[\s　]+', '', r['text']) for r in RAW_Y[exam])


def longest_common(t, p):
    for L in range(len(t), 5, -1):
        if any(t[i:i + L] in p for i in range(len(t) - L + 1)):
            return L
    return 0


RAW_Y = {}


def main():
    global RAW_Y
    log = OrderedDict()
    n_drop = n_fix = n_doubt = 0
    for school, blob in DATA.items():
        RAW_Y = {ex: [r for r in RAW[school] if r['exam'] == ex]
                 for ex in {r['exam'] for r in RAW[school]}}
        for seg in blob['segments']:
            for b in seg['books']:
                topics = []
                for t in b['topics']:
                    ev = []
                    for e in t['evidence']:
                        if e['pid'] not in SYNTH:
                            ev.append(e)
                            continue
                        act, exam, pages, why = FIX[e['pid']]
                        rec = dict(pid=e['pid'], school=school, seg=seg['key'], topic=t['title'],
                                   from_exam=e['exam'], from_year=e['year'], act=act, why=why,
                                   text=e['text'][:70])
                        if act == 'A':
                            yr = int(re.search(r'(\d{4})年', exam).group(1))
                            # 若目标卷同一条目下已有同句真实题面，补录即为重复，不再计入
                            dup = any(
                                x['exam'] == exam and x['text'].strip().rstrip('。') ==
                                e['text'].strip().rstrip('。')
                                for t2 in b['topics'] for x in t2['evidence']
                                if t2['title'] == t['title'] and x['pid'] not in SYNTH)
                            if dup:
                                rec.update(act='D', why=rec['why'] + '；目标卷同条目已有同句真实题面，'
                                                                   '补录属重复，不再计入')
                                n_drop += 1
                            else:
                                e2 = dict(e)
                                e2.update(exam=exam, year=yr, pages=pages,
                                          excerpt=e.get('excerpt') or e['text'],
                                          page_precision='原卷题目要求定位', recall=False)
                                ev.append(e2)
                                rec.update(to_exam=exam, to_year=yr)
                                n_fix += 1
                        elif act == 'F':
                            n_doubt += 1
                        else:
                            n_drop += 1
                        log[f'{school}-{e["pid"]}'] = rec
                    if not ev:
                        log[f'{school}-无证据条目-{t["title"]}'] = dict(
                            school=school, seg=seg['key'], topic=t['title'], act='T',
                            why='补录证据移出后该条已无任何证据，未进正文')
                        continue
                    ev.sort(key=lambda x: (x['year'], x['pid']))
                    t['evidence'] = ev
                    t['years'] = sorted({x['year'] for x in ev})
                    t['frequency'] = len(t['years'])
                    t['all_frequency'] = t['frequency']
                    t['recall_hits'] = sorted({x['year'] for x in ev if x.get('recall')})
                    topics.append(t)
                b['topics'] = topics
                b['local_max'] = len(topics)
                b['years_seen'] = sorted({e['year'] for t in topics for e in t['evidence']})
        # 重算分册统计
        for seg in blob['segments']:
            ev = [e for b in seg['books'] for t in b['topics'] for e in t['evidence']]
            seg['stats'] = dict(topics=sum(b['local_max'] for b in seg['books']),
                                evidence=len(ev),
                                recall=sum(1 for e in ev if e.get('recall')))
            ys = sorted({e['year'] for e in ev})
            seg['years'] = f'{ys[0]}—{ys[-1]}' if ys else ''
            seg['all_years'] = ys
            pages = sorted({int(e['pages']) for e in ev if str(e.get('pages', '')).isdigit()})
            seg['page_real'] = f'p{pages[0]}—p{pages[-1]}' if pages else ''
        uniq = {(t['discipline'], t['section'], t['title'])
                for s in blob['segments'] for b in s['books'] for t in b['topics']}
        # 重算时保留 prepare 阶段带过来的口径字段（unique_evidence／v4_evidence 等）
        uev = {(e['exam'], e['pid'], e['text'].strip())
               for sg in blob['segments'] for b in sg['books'] for t in b['topics']
               for e in t['evidence']}
        blob['stats'] = dict(blob['stats'],
                             unique_topics=len(uniq),
                             unique_evidence=len(uev),
                             instances=sum(s['stats']['topics'] for s in blob['segments']),
                             evidence=sum(s['stats']['evidence'] for s in blob['segments']),
                             recall=sum(s['stats']['recall'] for s in blob['segments']),
                             segments=len(blob['segments']),
                             books=len({b['key'] for s in blob['segments'] for b in s['books']}))

    DATA_P.write_text(json.dumps(DATA, ensure_ascii=False, indent=1))
    (W / 'v5证据补正清单.json').write_text(json.dumps(log, ensure_ascii=False, indent=1))
    print('===== 证据补正结果')
    for school, blob in DATA.items():
        st = blob['stats']
        print(f'  {school}: 唯一考点 {st["unique_topics"]}　段内条目 {st["instances"]}　'
              f'证据 {st["evidence"]}　回忆版 {st["recall"]}')
    drop = [k for k, v in log.items() if v.get('act') == 'D']
    fix = [k for k, v in log.items() if v.get('act') == 'A']
    doubt = [k for k, v in log.items() if v.get('act') == 'F']
    print(f'  移出正文（重复／无题面可定位）{n_drop} 条；改卷别 {n_fix} 条；存疑移出 {n_doubt} 条')
    for k in fix:
        print(f'    改：{k} → {log[k]["to_exam"]}')
    for k in doubt:
        print(f'    存疑：{k} {log[k]["text"]}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
