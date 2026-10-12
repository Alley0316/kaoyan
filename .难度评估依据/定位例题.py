from pathlib import Path
import json,re,pdfplumber
w=Path('/home/zyf/考研/.高频考点工作');root=Path('/home/zyf/考研/四校真题换电脑交接/四校真题打印复核版');out={}
selected={'云南民族大学':[491,512,505,457,463],'黑龙江大学':[903,898,1024,1129,1158,1164],'安徽大学':[870,884,890,858,853],'西安工业大学':[595,598,627,730,773,774]}
for school,ids in selected.items():
 rows={x['pid']:x for x in json.loads((w/(school+'_历年真题_打印版.json')).read_text())};ix=json.loads((w/'pdf_index.json').read_text())[school]
 with pdfplumber.open(root/(school+'_历年真题_打印版.pdf')) as d:
  texts=[re.sub(r'\s+','',p.extract_text() or '') for p in d.pages]
  for pid in ids:
   x=rows[pid];area=ix[x['eid']-1];candidate=x['text']
   if pid==858:candidate='结合莫言的文学作品被改编成电影'
   n=re.sub(r'\s+','',candidate);key=n[:min(len(n),18)]
   pp=[i+1 for i in range(area['start']-1,area['end']) if key in texts[i]]
   info={'school':school,'pid':pid,'exam':x['exam'],'text':candidate,'pages':pp or list(range(area['start'],area['end']+1))}
   out[f'{school}:{pid}']=info
   print(school,pid,info['pages'],candidate[:70])
Path('/home/zyf/考研/.难度评估依据/例题定位.json').write_text(json.dumps(out,ensure_ascii=False,indent=2))
