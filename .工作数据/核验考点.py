import json,re,hashlib
from pathlib import Path
from docx import Document
import pdfplumber
W=Path(__file__).resolve().parent;O=Path('/home/zyf/考研/四校复核新版');D=json.loads((W/'复核考点.json').read_text());report={}
def norm(x):return re.sub(r'\s+','',x)
for s,ts in D.items():
 docx=O/(s+'_高频考点复核版.docx');pdf=W/'render'/docx.stem/(docx.stem+'.pdf')
 with pdfplumber.open(pdf) as pp:
  tx=norm('\n'.join(p.extract_text() or '' for p in pp.pages));paras=[p.text for p in Document(docx).paragraphs if p.text]
  missing=[p for p in paras if norm(p) not in tx];assert not missing,(s,missing)
  for t in ts:
   assert t['title'] in ''.join(paras)
   assert t['frequency']==len({2005 if s=='安徽大学' and t['discipline']=='基础' and y==2007 else y for y in t['years']})
   assert set(t['recent_hits'])==set(t['years'])&set(t['recent_years'])
   assert not any(s=='黑龙江大学' and (847<=e['pid']<=866 or e['pid']==774) for e in t['evidence'])
  for i,p in enumerate(pp.pages,1):
   assert abs(p.width-595.28)<1 and abs(p.height-841.89)<1
   cc=[c for c in p.chars if 40<c['top']<790]
   assert all(c['x0']>=48 and c['x1']<=548 for c in cc),(s,i,'边界')
   assert '\ufffd' not in (p.extract_text() or '')
   assert (pdf.parent/f'page-{i}.png').exists()
  report[s]={'pages':len(pp.pages),'repeat':sum(t['frequency']>=2 for t in ts),'single':sum(t['frequency']==1 for t in ts),'evidence':sum(len(t['evidence']) for t in ts),'docx_pdf_text_match':True,'visual_review':'pending'}
(W/'考点机械核验.json').write_text(json.dumps(report,ensure_ascii=False,indent=2));print(json.dumps(report,ensure_ascii=False))
