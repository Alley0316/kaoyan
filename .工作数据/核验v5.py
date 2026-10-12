# -*- coding: utf-8 -*-
"""v5 机械核验：分段完整性、编号连续性、页数与版面、与数据的逐条一致性。

用法：python3 核验v5.py [--pdf-dir v5]
输出：.工作数据/v5核验结果.json，并在终端打印每册结论；退出码 1 表示有册不通过。
"""
import json
import re
import subprocess
import sys
from pathlib import Path

W = Path(__file__).resolve().parent
ROOT = W.parent
PDFDIR = (Path(sys.argv[sys.argv.index('--pdf-dir') + 1]).resolve()
          if '--pdf-dir' in sys.argv else ROOT / 'v5')
DATA = json.loads((W / 'v5考点数据.json').read_text())
VERSION = 'v5'


def norm(x):
    return re.sub(r'\s+', '', x or '')


report = {}
fail = 0
for school, blob in DATA.items():
    fn = f'{school}_高频考点_{VERSION}.pdf'
    pdf = PDFDIR / fn
    txt = subprocess.run(['pdftotext', '-layout', str(pdf), '-'],
                         capture_output=True, text=True).stdout
    flat = norm(txt)
    info = subprocess.run(['pdfinfo', str(pdf)], capture_output=True, text=True).stdout
    pages = int(re.search(r'^Pages:\s+(\d+)', info, re.M).group(1))
    size = re.search(r'^Page size:\s+(.+)$', info, re.M).group(1).strip()
    checks = {}

    segs = blob['segments']
    # ① 分段标题与科目代码都在正文里
    missing_seg = [s['label'] for s in segs if norm(s['label']) not in flat]
    checks['缺少分段标题'] = missing_seg
    missing_code = [s['code'] for s in segs if norm(s['code']) not in flat]
    checks['缺少科目代码'] = missing_code

    # ② 每条考点的标题都在 PDF 文本中
    missing = [t['title'] for s in segs for b in s['books'] for t in b['topics']
               if norm(t['title']) not in flat]
    checks['条目缺失'] = missing

    # ③ 编号：每段每本书内从 1 连续编到 local_max
    nums_ok, nums_bad = [], []
    for si, seg in enumerate(segs, 1):
        for bi, b in enumerate(seg['books'], 1):
            marker = norm(f'参考书 {si}-{bi}：{b["name"]}')
            start = flat.find(marker)
            nxt = None
            if bi < len(seg['books']):
                nxt = flat.find(norm(f'参考书 {si}-{bi + 1}：{seg["books"][bi]["name"]}'))
            elif si < len(segs):
                nxt = flat.find(norm(f'第 {si + 1} 段'))
            block = flat[start:nxt] if start >= 0 and nxt and nxt > start else (
                flat[start:] if start >= 0 else '')
            found = [int(m) for m in re.findall(r'(?<![0-9])(\d{1,3})\.', block)]
            seq, seen = [], set()
            for x in found:
                if 1 <= x <= b['local_max'] and x not in seen:
                    seen.add(x)
                    seq.append(x)
            if start < 0:
                nums_bad.append(dict(book=b['name'][:20], err='未找到参考书标记'))
            elif seq == list(range(1, b['local_max'] + 1)):
                nums_ok.append(b['local_max'])
            else:
                nums_bad.append(dict(book=b['name'][:20], expect=b['local_max'],
                                     got=len(seq), head=seq[:10]))
    checks['编号连续的书'] = nums_ok
    checks['编号异常'] = nums_bad
    checks['编号合计'] = sum(nums_ok)
    checks['条目合计'] = sum(b['local_max'] for s in segs for b in s['books'])

    # ④ 页码与纸张
    checks['页数'] = pages
    checks['纸张'] = size
    over = [ln for ln in txt.splitlines() if len(ln) > 118]
    checks['超宽行数'] = len(over)
    checks['超宽示例'] = over[:2]

    # ⑤ 出处页码必须落在「该卷」在打印复核版里的起止页内（逐卷核对，不用段范围近似）
    INDEX = json.loads((ROOT / '.高频考点工作' / 'pdf_index.json').read_text())
    bad_page = []
    for seg in segs:
        idx = {(a['year'], a['subject']): (a['start'], a['end']) for a in INDEX[school]}
        for b in seg['books']:
            for t in b['topics']:
                for e in t['evidence']:
                    a = idx.get((e['year'], e.get('subject')))
                    r = re.findall(r'\d+', str(e.get('pages', '')))
                    if not a or not r:
                        continue
                    lo_, hi_ = int(r[0]), int(r[-1])
                    if not (a[0] <= lo_ and hi_ <= a[1]):
                        bad_page.append(dict(seg=seg['key'], title=t['title'], pid=e['pid'],
                                             exam=e['exam'], pages=str(e['pages']),
                                             range=f'{a[0]}-{a[1]}'))
    checks['出处页码越卷'] = bad_page[:5]
    checks['出处页码越卷数'] = len(bad_page)

    # ⑥ 年份显示：正文里每个证据年份都应能在这段正文找到（抽验首末两条）
    miss_year = []
    for seg in segs:
        for b in seg['books']:
            for t in b['topics']:
                for e in (t['evidence'][:1] + t['evidence'][-1:]):
                    if str(e['year']) not in flat:
                        miss_year.append(dict(seg=seg['key'], title=t['title'], year=e['year']))
    checks['年份未出现'] = miss_year[:5]
    checks['年份未出现数'] = len(miss_year)

    ok = (not missing_seg and not missing and not nums_bad and not bad_page
          and not miss_year and checks['超宽行数'] == 0)
    report[fn] = dict(ok=ok, **checks)
    fail += 0 if ok else 1
    print(f'{"PASS" if ok else "FAIL"}  {fn}  页 {pages}  段 {len(segs)}  '
          f'条目 {checks["条目合计"]}  编号合计 {checks["编号合计"]}'
          f'{"  缺条目 " + str(len(missing)) if missing else ""}'
          f'{"  编号异常 " + str(nums_bad) if nums_bad else ""}'
          f'{"  页码越段 " + str(len(bad_page)) if bad_page else ""}'
          f'{"  超宽 " + str(checks["超宽行数"]) if checks["超宽行数"] else ""}')

(W / f'{VERSION}核验结果.json').write_text(json.dumps(report, ensure_ascii=False, indent=1))
print(f'\n合计 {len(report)} 册，失败 {fail} 册')
sys.exit(1 if fail else 0)
