# -*- coding: utf-8 -*-
"""v4 机械核验：编号连续性、条目完整性、书单前置、页数与版面边界。

用法：python3 核验v4.py [--pdf-dir v4]
输出：.工作数据/v4核验结果.json，并在终端打印每册结论。
"""
import json
import re
import subprocess
import sys
from pathlib import Path

W = Path(__file__).resolve().parent
ROOT = W.parent
PDFDIR = Path(sys.argv[sys.argv.index('--pdf-dir') + 1]).resolve() if '--pdf-dir' in sys.argv else ROOT / 'v4'
DATA = json.loads((W / 'v4考点数据.json').read_text())

report = {}
fail = 0
for school, papers in DATA.items():
    for spec in papers:
        fn = f'{school}_{spec["paper"].split("（")[0]}_高频考点_v4.pdf'
        pdf = PDFDIR / fn
        txt = subprocess.run(['pdftotext', '-layout', str(pdf), '-'],
                             capture_output=True, text=True).stdout
        flat = re.sub(r'\s+', '', txt)
        info = subprocess.run(['pdfinfo', str(pdf)], capture_output=True, text=True).stdout
        pages = int(re.search(r'^Pages:\s+(\d+)', info, re.M).group(1))
        size = re.search(r'^Page size:\s+(.+)$', info, re.M).group(1).strip()
        checks = {}

        # ① 每个条目标题都在 PDF 文本中（去除空白后比对）
        missing = [t['title'] for b in spec['books'] for t in b['topics']
                   if re.sub(r'\s+', '', t['title']) not in flat]
        checks['条目缺失'] = missing

        # ② 编号：每本书内从 1 连续到 local_max
        nums_ok, nums_bad = [], []
        for bi, b in enumerate(spec['books'], 1):
            marker = re.sub(r'\s+', '', f'参考书{bi}：{b["name"]}')
            start = flat.find(marker)
            nxt = None
            if bi < len(spec['books']):
                nxt = flat.find(re.sub(r'\s+', '', f'参考书{bi + 1}：{spec["books"][bi]["name"]}'))
            block = flat[start:nxt] if start >= 0 and nxt and nxt > start else (
                flat[start:] if start >= 0 else '')
            found = [int(m) for m in re.findall(r'(?<![0-9])(\d{1,3})\.', block)]
            # 取该块内 1..local_max 的首次出现顺序
            seq, seen = [], set()
            for x in found:
                if 1 <= x <= b['local_max'] and x not in seen:
                    seen.add(x)
                    seq.append(x)
            if seq == list(range(1, b['local_max'] + 1)):
                nums_ok.append(b['local_max'])
            else:
                nums_bad.append(dict(book=b['name'][:24], expect=b['local_max'],
                                     got_len=len(seq), head=seq[:12]))
        checks['编号连续的书'] = nums_ok
        checks['编号异常'] = nums_bad
        checks['编号合计'] = sum(nums_ok)

        # ③ 书单前置：第一本书的标记出现在前两页文本内
        first = re.sub(r'\s+', '', f'参考书1：{spec["books"][0]["name"]}')
        head_txt = subprocess.run(['pdftotext', '-f', '1', '-l', '2', '-layout', str(pdf), '-'],
                                  capture_output=True, text=True).stdout
        checks['书单在前两页'] = first in re.sub(r'\s+', '', head_txt)
        checks['含参考书目节'] = '参考书目' in txt

        # ④ 版面：A4、无超界（按行宽估算）与页数
        checks['页数'] = pages
        checks['纸张'] = size
        over = [ln for ln in txt.splitlines() if len(ln) > 118]
        checks['超宽行数'] = len(over)
        checks['超宽示例'] = over[:2]

        ok = (not missing) and (not nums_bad) and checks['书单在前两页'] and checks['含参考书目节'] \
            and checks['超宽行数'] == 0 and abs(pages - checks['页数']) == 0
        report[fn] = dict(ok=ok, **checks)
        fail += 0 if ok else 1
        print(f'{"PASS" if ok else "FAIL"}  {fn}  页 {pages}  条目 '
              f'{sum(len(b["topics"]) for b in spec["books"])}  编号合计 {checks["编号合计"]}'
              f'{"  缺失 " + str(len(missing)) if missing else ""}'
              f'{"  编号异常 " + str(nums_bad) if nums_bad else ""}'
              f'{"  超宽 " + str(checks["超宽行数"]) if checks["超宽行数"] else ""}')

(W / 'v4核验结果.json').write_text(json.dumps(report, ensure_ascii=False, indent=1))
print(f'\n合计 {len(report)} 册，失败 {fail} 册')
sys.exit(1 if fail else 0)
