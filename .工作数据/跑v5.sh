#!/usr/bin/env bash
# v5 固化流水线：按固定顺序重跑数据层、考点册、渲染、页数、总结、核验。
# 用法：bash .工作数据/跑v5.sh
set -euo pipefail
cd /home/zyf/考研
P=/home/zyf/.dsh/shared-runtime/dependencies/python/bin/python3

echo "① 数据层（分段）"; $P .工作数据/准备v5数据.py | tail -4
echo "② 证据补正（占位证据的卷别年份）"; $P .工作数据/补正v5证据.py | head -6
echo "③ 真实性硬校验"; $P .工作数据/核对v5真实性.py > .tmpv5check/auth.log 2>&1 && tail -2 .tmpv5check/auth.log
echo "④ 生成四册 docx"; $P .工作数据/生成v5考点文档.py | tail -5
echo "⑤ 渲染 PDF"; rm -f v5/*.pdf; $P .工作数据/render_pdf.py /home/zyf/考研/v5 | tail -1
echo "⑥ 记录页数"; $P - <<'PY'
import json, subprocess, re
from pathlib import Path
v5 = Path('v5'); m = json.loads(Path('.工作数据/v5册清单.json').read_text())
cnt = {}
for x in m:
    pdf = v5 / x['file'].replace('.docx', '.pdf')
    info = subprocess.run(['pdfinfo', str(pdf)], capture_output=True, text=True).stdout
    cnt[x['file']] = int(re.search(r'^Pages:\s+(\d+)', info, re.M).group(1))
Path('.工作数据/v5页数.json').write_text(json.dumps(cnt, ensure_ascii=False, indent=1))
print(' 页数', cnt, '合计', sum(cnt.values()))
PY
echo "⑦ 生成总结"; $P .工作数据/生成v5总结文档.py | head -1
echo "⑧ 渲染总结"; $P .工作数据/render_pdf.py /home/zyf/考研/v5 | tail -1
echo "⑨ 机械核验"; $P .工作数据/核验v5.py | tail -2; $P .工作数据/核验v5总结.py | head -1
echo "完成"
