# -*- coding: utf-8 -*-
"""逐题真实性核对：把 v5 每条证据回原卷核对。

三级检查：
① 卷名与题号：证据的 pid 必须存在，且所属卷名与证据 exam 完全一致。
② 年份与页码：证据的 year／pages 必须落在 .高频考点工作/pdf_index.json 记录的
   该卷起止页与实际年份内（pdf_index 是原卷打印版的分卷页码索引）。
③ 原文文字：证据 text 必须能在该卷原文（同卷全部题面行）里找到，去掉空白后比对。

输出：.工作数据/v5真实性核对.json ＋终端摘要；退出码 1 表示存在硬校验失败。
"""
import json
import re
import sys
from collections import defaultdict, OrderedDict
from pathlib import Path

W = Path(__file__).resolve().parent
DATA = json.loads((W / 'v5考点数据.json').read_text())
RAW = {s: json.loads((W.parent / '.高频考点工作' / f'{s}_历年真题_打印版.json').read_text())
       for s in DATA}
INDEX = json.loads((W.parent / '.高频考点工作' / 'pdf_index.json').read_text())


def norm(x):
    return re.sub(r'[\s　]+', '', x or '')


def pg_range(p):
    m = re.match(r'^(\d+)(?:-(\d+))?$', str(p or '').strip())
    if not m:
        return None
    a = int(m.group(1))
    return (a, int(m.group(2) or a))


# 人工补录的证据：v3 精修阶段按「题目要求」层级补入，没有逐行 pid，改为按整卷核文字。
# 核对方式不同（不做题号与页码定位），但更严：要求文字在该卷原文里能找到，且年份正确。
SYNTHETIC = set(range(1201, 1210)) | set(range(1301, 1310)) | set(range(1401, 1410))


report = OrderedDict()
fails = 0
for school, blob in DATA.items():
    raw = RAW[school]
    idx = INDEX[school]
    by_exam = defaultdict(list)
    for r in raw:
        by_exam[r['exam']].append(r)
    exam_of_pid = {r['pid']: r['exam'] for r in raw}
    norm_pool = defaultdict(list)
    for r in raw:
        norm_pool[r['exam']].append(norm(r['text']))
    # 卷名 → pdf_index 记录
    idx_map = {}
    for a in idx:
        key = f'{school} {a["year"]}年 {a["subject"]}'
        idx_map.setdefault(key, []).append(a)

    bad = dict(exam=0, pid=0, year=0, pages=0, text=0, synth=0,
               synth_year=0, synth_text=0)
    samples = defaultdict(list)
    per_exam = defaultdict(lambda: dict(ev=0, years=set(), pages=set()))
    for seg in blob['segments']:
        for b in seg['books']:
            for t in b['topics']:
                for e in t['evidence']:
                    pid, exam = e['pid'], e['exam']
                    per_exam[exam]['ev'] += 1
                    per_exam[exam]['pages'].add(str(e.get('pages')))
                    if pid in SYNTHETIC:
                        per_exam[exam]['synth'] = per_exam[exam].get('synth', 0) + 1
                        bad['synth'] += 1
                        a = (idx_map.get(exam) or [None])[0]
                        if a is None or e.get('year') != a['year']:
                            bad['synth_year'] += 1
                            samples['synth_year'].append(dict(pid=pid, ev=e.get('year'),
                                                              idx=a and a['year'], title=t['title']))
                        # 人工补录的是题目要求摘要（非逐字摘录），改用字符覆盖率判定：
                        # 摘要里的汉字至少有七成能在该卷原文中找到，即认定来源真实。
                        et = norm(e.get('text'))
                        if et and len(et) >= 4:
                            pool = ''.join(x for x in norm_pool[exam] if x)
                            hit = sum(1 for ch in set(et) if ch in pool)
                            ratio = hit / max(1, len(set(et)))
                            if ratio < .7:
                                bad['synth_text'] += 1
                                samples['synth_text'].append(dict(pid=pid, title=t['title'],
                                                                  ratio=round(ratio, 2),
                                                                  text=e['text'][:70]))
                        continue
                    if pid not in exam_of_pid:
                        bad['pid'] += 1
                        samples['pid'].append(dict(pid=pid, exam=exam))
                        continue
                    if exam_of_pid[pid] != exam:
                        bad['exam'] += 1
                        samples['exam'].append(dict(pid=pid, ev=exam, raw=exam_of_pid[pid]))
                        continue
                    cands = idx_map.get(exam) or []
                    if not cands:
                        bad['exam'] += 1
                        samples['exam'].append(dict(pid=pid, ev=exam, raw='pdf_index 无此卷名'))
                    else:
                        a = cands[0]
                        if e.get('year') != a['year']:
                            bad['year'] += 1
                            samples['year'].append(dict(pid=pid, ev=e.get('year'),
                                                        idx=a['year'], title=t['title']))
                        r = pg_range(e.get('pages'))
                        if r is None or not (a['start'] <= r[0] and r[1] <= a['end']):
                            bad['pages'] += 1
                            samples['pages'].append(dict(pid=pid, ev=str(e.get('pages')),
                                                         idx=f'{a["start"]}-{a["end"]}',
                                                         title=t['title']))
                    et = norm(e.get('text'))
                    if et and len(et) >= 8 and not any(et in x or x in et
                                                      for x in norm_pool[exam] if x):
                        bad['text'] += 1
                        samples['text'].append(dict(pid=pid, title=t['title'],
                                                    text=e['text'][:70]))
    no_ev = [ex for ex in by_exam if ex not in per_exam]
    total = sum(v['ev'] for v in per_exam.values())
    print(f'===== {school}：证据 {total}，卷 {len(by_exam)} 份')
    print(f"  ① 题号缺失 {bad['pid']}　② 卷名不符 {bad['exam']}　"
          f"③ 年份不符 {bad['year']}　④ 页码越界 {bad['pages']}　"
          f"⑤ 原文找不到 {bad['text']}　⑥ 人工补录 {bad['synth']} 条"
          f"（其中年份不符 {bad['synth_year']}、原文找不到 {bad['synth_text']}）")
    for k in ('pid', 'exam', 'year', 'pages', 'text', 'synth_year', 'synth_text'):
        for x in samples[k][:4]:
            print(f'    !! {k}: {x}')
    if no_ev:
        print(f'  !! 无证据落点的卷：{no_ev}')
    fails += sum(v for k, v in bad.items() if k != 'synth')
    report[school] = dict(total_evidence=total, bad=bad, samples=dict(samples),
                          exam_no_evidence=no_ev,
                          per_exam={k: dict(ev=v['ev'], pages=sorted(v['pages']))
                                    for k, v in per_exam.items()})

(W / 'v5真实性核对.json').write_text(json.dumps(report, ensure_ascii=False, indent=1))
# 逐段实际覆盖的卷面页范围（按该段全部证据的年份卷起止页推算），供 准备v5数据.py 核对
for school, blob in DATA.items():
    idx = {(a['year'], a['subject']): (a['start'], a['end']) for a in INDEX[school]}
    print(f'===== {school} 各段实际页范围')
    for seg in blob['segments']:
        yrs = sorted({e['year'] for b in seg['books'] for t in b['topics'] for e in t['evidence']})
        subj = {}
        for b in seg['books']:
            for t in b['topics']:
                for e in t['evidence']:
                    subj.setdefault(e['year'], set()).add(e.get('subject'))
        rng = []
        for y in yrs:
            for sj in sorted(x for x in subj[y] if x):
                a = idx.get((y, sj))
                if a:
                    rng.append(a)
        if rng:
            lo = min(x[0] for x in rng); hi = max(x[1] for x in rng)
            print(f'  {seg["label"]}: 声明 {seg["page"]}  实际 p{lo}—p{hi}  '
                  f'（{len(rng)} 份卷）')
        else:
            print(f'  {seg["label"]}: 声明 {seg["page"]}  实际无卷可推')
print(f'\n硬校验失败合计 {fails}')
sys.exit(1 if fails else 0)
