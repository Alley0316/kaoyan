import json,re
from pathlib import Path
from collections import defaultdict
W=Path(__file__).resolve().parent
schools=['安徽大学','黑龙江大学','云南民族大学','西安工业大学']
rows={s:{x['pid']:x for x in json.loads((W/(s+'_历年真题_打印版.json')).read_text())} for s in schools}
index=json.loads((W/'pdf_index.json').read_text())
# Source page 5 is a handwritten 2024 recall sheet; retain only clearly legible prompts.
for pid,text in [(1201,'举例说明组合关系'),(1202,'举例说明词义引申方式中的隐喻和转喻'),(1203,'地域方言'),(1204,'文言翻译 大道之行也及韩非子五蠹节选'),(1205,'两个以的用法分析'),(1206,'名词作状语的类型 请举例说明')]:
 rows['黑龙江大学'][pid]={'pid':pid,'eid':26,'exam':'黑龙江大学 2024年 古代汉语与语言学理论 回忆版','text':text,'visual_source':'原题PDF第5页 手写回忆版'}
for school,extra in json.loads((W/'visual_evidence.json').read_text()).items():
 for row in extra:rows[school][row['pid']]=row
# Correct locators before classifying; these point to question instructions, not material bodies.
p=W/'topics.txt';t=p.read_text().replace('345,442,496','345,417,496').replace('190,266,363,447','190,266,364,446').replace('555,689,760','555,678,760').replace('556,693,774','556,697,774').replace(',882\n', '\n')
p.write_text(t)
# Keep pedagogically distinct topics separate instead of inflating their counts by union.
def split_line(school,title,newlines):
 global t
 active=False;out=[]
 for line in t.splitlines():
  if line in schools:active=line==school
  if active and line.split('|')[1:2]==[title]:out.extend(newlines)
  else:out.append(line)
 t='\n'.join(out)+'\n'
split_line('安徽大学','玄言诗与永明体',[
 '古代文学|玄言诗|掌握时代、诗风及其与山水诗的关系，联系代表作家。|44,567',
 '古代文学|永明体|掌握声律要求、代表作家及其对近体诗发展的意义。|515,629'])
split_line('安徽大学','先锋小说与新写实小说',[
 '现当代文学|先锋小说|准备叙事、语言实验和代表作品，结合文学史说明特点。|321,566,677,886',
 '现当代文学|新写实小说|掌握生活书写、叙事特点及与既有现实主义的关系。|89,144,247,739'])
split_line('安徽大学','连绵词与离合词',[
 '现代汉语 词汇与语法|连绵词|定义、语素数量与分类联系具体词例。|285,511,778,840@连绵词',
 '现代汉语 词汇与语法|离合词|从扩展形式说明词与短语的区别，举例解释。|776,840@离合词'])
split_line('云南民族大学','外来词与词缀构词',[
 '现代汉语 文字与词汇|外来词|掌握外来词的定义、主要类型及与借词的关系。|101@借词,353@外来词,415',
 '现代汉语 文字与词汇|词根 词缀与派生词|辨别词根词缀，解释派生方式并给出实例。|22@派生词,282,291@派生词'])
split_line('黑龙江大学','本义 古今词义与通假',[
 '古代汉语|汉字本义的辨别|利用字形、古书用例与语境说明辨别本义的方法。|145,314',
 '古代汉语|古今词义与古音通假|古今异义、引申与通假分别说明，按词义与字际关系辨别。|233,660,758,830'])
p.write_text(t)
# Infer the type from the enclosing question heading, with explicit exceptions for embedded groups.
def typename(s):
 if any(x in s for x in ['名词解释','词语解释','解释下列术语','解释术语','概念题']):return '名词'
 if any(x in s for x in ['简答','简述','问答','问答题']):return '问答'
 if '填空' in s:return '填空'
 if '论述' in s:return '论述'
 if '分析' in s or '辨析' in s or '改正' in s or '修改病句' in s:return '分析'
 if '赏析' in s or '鉴赏' in s:return '鉴赏'
 if '译' in s or '标点' in s or '句读' in s:return '阅读'
 if '写作' in s or '作文' in s:return '写作'
 if any(x in s for x in ['指出','写出','分析','辨析','请将']):return '分析'
 if '运用' in s:return '运用'
 return None
for school,rr in rows.items():
 context='题目';last=None
 for x in rr.values():
  if x['exam']!=last:context='题目';last=x['exam']
  text=x['text'].strip();m=re.match(r'^[一二三四五六七八九十]+(?:[、.．，。\s]|(?=概念|名词|简答|论述|分析|填空|问答|阅读|指出|写出))',text) or re.match(r'^\d+[、.．]\s*(名词解释|解释术语|简答|简述题|问答|分析题|论述|填空)',text) or re.match(r'^(名词解释|概念题|论述题|简答题|分析题)',text)
  if m:context=typename(text[:100]) or context
  x['type']=context
  if x['pid']>=1200:x['type']='名词' if x['pid']==1203 else ('阅读' if x['pid'] in (1204,1205) else '问答')
  if 'type_override' in x:x['type']=x['type_override']
  x['year']=int(re.search(r'(20\d{2})年',x['exam'])[1])
  x['subject']=x['exam'].split('年 ',1)[1]
  a=index[school][x['eid']-1]
  assert (a['year'],a['subject'])==(x['year'],x['subject'])
  x['pages']=str(a['start']) if a['start']==a['end'] else f"{a['start']}-{a['end']}"
  x['recall']='回忆' in x['subject'] or (school=='黑龙江大学' and x['year']>=2024) or (school=='云南民族大学' and x['year'] in (2020,2021,2023,2024)) or (school=='安徽大学' and x['year']>=2017)
# The supplemental modern-Chinese section at the end of the 2020 GuLi grouping has no separate year label.
# Exclude its prompts from frequency evidence rather than infer their exam year.
alltopics=defaultdict(list);school=None
for line in t.splitlines():
 if line in schools:school=line;continue
 if not line:continue
 sec,title,points,refs=line.split('|');ev=[]
 for ref in refs.split(','):
  a,*b=ref.split('@',1);pid=int(a)
  if school=='黑龙江大学' and 847<=pid<=866:continue
  row=rows[school][pid].copy();key=b[0] if b else None
  if key:
   assert re.sub(r'\s+','',key) in re.sub(r'\s+','',row['text']),(school,pid,key,row['text'])
   row['excerpt']=key
  else:
   tx=row['text'].strip()
   # Remove supplied solution and coaching annotations; examples use a prompt excerpt.
   tx=re.split(r'[：:]',tx,1)[0] if row['type']=='名词' else tx
   tx=re.sub(r'（(?:教辅|郭锡良|黄廖本|教材|平常|看是否).*?）','',tx)
   tx=re.sub(r'^[一二三四五六七八九十\d]+\s*[、.．]\s*','',tx)
   tx=re.sub(r'(?:郭锡良|黄廖本|教辅|教材)[上下本]?[Pp]\d+(?:-\d+)?','',tx)
   tx=re.sub(r'\s+',' ',tx)
   row['excerpt']=tx[:58]+('…' if len(tx)>58 else '')
  if school=='安徽大学' and pid in (856,857,858,840,831):
   row['type']='问答' if key and '问答' in row['text'][:row['text'].find(key)] else '名词'
  if school=='黑龙江大学' and pid in (48,116,207,291,413,506,588,679,746,820,968,1043,1109,1149):row['type']='阅读'
  if school=='黑龙江大学' and pid in (1017,1020,1023,1026,1029,1093,1095,1098,1137,1139):row['type']='分析'
  if school=='黑龙江大学' and pid in (1033,1034,1036,1037):row['type']='问答'
  if school=='黑龙江大学' and pid in (1063,1064,1075,1076):row['type']='论述'
  if school=='西安工业大学' and pid in (59,393,612,614,616):row['type']='分析'
  if school=='西安工业大学' and pid in (640,642):row['type']='阅读'
  if school=='安徽大学' and pid in (35,133,162,262,341,412,491,541,594,647,695,759,764,793,824,850,861,893):row['type']='阅读'
  if school=='安徽大学' and pid in (72,167,256,345,417,496,546,598,650,698):row['type']='鉴赏'
  if school=='安徽大学' and pid in (190,266,364,446,504,556,619,667,717,767,797,825,852,865,897):row['type']='写作'
  if school=='云南民族大学' and pid in (258,259,261):row['type']='未明'
  if school=='云南民族大学' and pid in (250,396,397,398):row['type']='论述'
  if school=='云南民族大学' and pid in (390,391,392,393,394,395):row['type']='分析'
  if school=='云南民族大学' and pid in (157,185,186,244,291,447):row['type']='名词'
  if school=='西安工业大学' and sec=='文学评论写作':row['type']='鉴赏' if '鉴赏' in title else ('论述' if '文论' in title else '评论')
  ev.append(row)
 distinct={x['eid'] for x in ev};years={x['year'] for x in ev}
 assert len(distinct)>=2,(school,title,'不足两卷')
 types=sorted(set(x['type'] for x in ev))
 alltopics[school].append({'section':sec,'title':title,'points':points,'evidence':ev,'frequency':len(distinct),'years':len(years),'types':types})
(W/'curated_topics.json').write_text(json.dumps(alltopics,ensure_ascii=False,indent=2))
for s,topics in alltopics.items():
 print(s,len(topics),'条；最高频',[(x['title'],x['frequency']) for x in sorted(topics,key=lambda x:-x['frequency'])[:6]])
print('全部引用存在；指定词匹配；页码与试卷索引匹配；每项至少两份不同卷次。')
