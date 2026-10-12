# -*- coding: utf-8 -*-
"""v5 总结册机械核验：统计数字、分段标签、书目版本要点与版面。

用法：python3 核验v5总结.py [--pdf v5/四校真题考点与难度对比_v5.pdf]
输出：.工作数据/v5总结核验结果.json ＋终端结论；退出码 1 表示不通过。
"""
import json
import re
import subprocess
import sys
from pathlib import Path

W = Path(__file__).resolve().parent
ROOT = W.parent
PDF = (Path(sys.argv[sys.argv.index('--pdf') + 1]).resolve()
       if '--pdf' in sys.argv else ROOT / 'v5' / '四校真题考点与难度对比_v5.pdf')
DATA = json.loads((W / 'v5考点数据.json').read_text())
MANIFEST = json.loads((W / 'v5册清单.json').read_text())
PAGES = json.loads((W / 'v5页数.json').read_text())


def norm(x):
    return re.sub(r'\s+', '', x or '')


txt = subprocess.run(['pdftotext', '-layout', str(PDF), '-'],
                     capture_output=True, text=True).stdout
flat = norm(txt)
info = subprocess.run(['pdfinfo', str(PDF)], capture_output=True, text=True).stdout
pages = int(re.search(r'^Pages:\s+(\d+)', info, re.M).group(1))
size = re.search(r'^Page size:\s+(.+)$', info, re.M).group(1).strip()
checks = {}

# ① 四校名与全部段落标签都在
checks['缺校名'] = [s for s in DATA if norm(s) not in flat]
missing_seg = [seg['label'] for blob in DATA.values() for seg in blob['segments']
               if norm(seg['label']) not in flat]
checks['缺分段标签'] = missing_seg
# 表格里的长标签会被换行拆开，故再按「·」与括号切片段核对一次
frag_bad = []
for blob in DATA.values():
    for seg in blob['segments']:
        for frag in re.split(r'[·（）()／/]', seg['label']):
            frag = norm(frag)
            if len(frag) >= 3 and frag not in flat:
                frag_bad.append(dict(label=seg['label'], frag=frag))
checks['缺分段标签片段'] = frag_bad

# ② 统计数字：每校的唯一考点、段内条目、证据三组数字都要出现
bad_num = []
for m in MANIFEST:
    for key, val in [('unique_topics', m['unique_topics']), ('instances', m['instances']),
                     ('evidence', m['evidence'])]:
        if str(val) not in flat:
            bad_num.append(dict(school=m['school'], key=key, value=val))
checks['统计数字缺失'] = bad_num

# ③ 总数：四校合计三类数字都要出现
tot = dict(unique=sum(m['unique_topics'] for m in MANIFEST),
           inst=sum(m['instances'] for m in MANIFEST),
           ev=sum(m['evidence'] for m in MANIFEST))
checks['合计缺失'] = [k for k, v in tot.items() if str(v) not in flat]
checks['合计'] = tot

# ④ 册页数与合计页数
checks['各册页数缺失'] = [f for f, p in PAGES.items() if str(p) not in flat]
checks['各册页数'] = PAGES
checks['页数合计'] = sum(PAGES.values())
checks['页数合计出现'] = str(sum(PAGES.values())) in flat

# ⑤ 书目版本要点（v5 纠正过的关键事实）必须在正文里
KEY_FACTS = [
    ('胡裕树', '安徽现代汉语 2007—2015 指定胡裕树本'),
    ('增订五版', '安徽 2016—2019 换黄廖五版'),
    ('王力', '安徽古代汉语 2007—2018 为王力本'),
    ('黄德宽', '安徽 2019 起换黄德宽本'),
    ('郭锡良', '西安古代汉语指定本；同时用于说明安徽未曾指定'),
    ('邵敬敏', '黑龙江 2026 起改邵敬敏本'),
    ('2025年版', '云南 815 于 2026 起改 2025 年版'),
    ('马学良', '云南 2024 转载另列书目，已说明不作依据'),
]
checks['关键事实缺失'] = [why for word, why in KEY_FACTS if norm(word) not in flat]

# ⑥ 版面
over = [ln for ln in txt.splitlines() if len(ln) > 118]
checks['超宽行数'] = len(over)
checks['超宽示例'] = over[:2]
checks['页数'] = pages
checks['纸张'] = size

ok = (not checks['缺校名'] and not checks['缺分段标签片段'] and not bad_num and not checks['合计缺失']
      and not checks['各册页数缺失'] and checks['页数合计出现']
      and not checks['关键事实缺失'] and checks['超宽行数'] == 0)
print(('PASS' if ok else 'FAIL') + f'  {PDF.name}  页 {pages}  合计 {tot}  页数合计 {checks["页数合计"]}')
for k in ['缺校名', '缺分段标签片段', '统计数字缺失', '合计缺失', '各册页数缺失', '关键事实缺失']:
    if checks[k]:
        print(f'  !! {k}: {checks[k]}')
if checks['超宽行数']:
    print(f'  !! 超宽行 {checks["超宽行数"]}: {checks["超宽示例"]}')
(W / 'v5总结核验结果.json').write_text(json.dumps(dict(ok=ok, **checks), ensure_ascii=False, indent=1))
sys.exit(0 if ok else 1)
