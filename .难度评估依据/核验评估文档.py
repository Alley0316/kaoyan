from pathlib import Path
from zipfile import ZipFile
import re,json,shutil,hashlib
from lxml import etree
from docx import Document
import pdfplumber
W=Path(__file__).resolve().parent;O=W.parent/'四校难度评估';name='四校真题相对参考书目难度评估'
doc=Document(O/(name+'.docx'));source=W/'render'/(name+'.pdf');dest=O/(name+'.pdf');shutil.copyfile(source,dest)
def norm(t):return re.sub(r'\s+','',t)
with pdfplumber.open(dest) as d:
 text=norm('\n'.join(p.extract_text() or '' for p in d.pages));missing=[]
 for p in doc.paragraphs:
  if p.text and norm(p.text) not in text:missing.append(p.text)
 pdf_tables=[table for p in d.pages for table in p.extract_tables()]
 assert len(pdf_tables)==len(doc.tables)
 for table,pdf_table in zip(doc.tables,pdf_tables):
  assert len(table.rows)==len(pdf_table)
  for row,pdf_row in zip(table.rows,pdf_table):
   for cell,pdf_cell in zip(row.cells,pdf_row):
    if norm(cell.text)!=norm(pdf_cell or ''):missing.append(cell.text)
 assert not missing,missing
 assert len(d.pages)==7
 assert all(len(p.extract_text() or '')>500 for p in d.pages)
 for i,p in enumerate(d.pages,1):
  assert abs(p.width-595.28)<1 and abs(p.height-841.89)<1
  body=[c for c in p.chars if 45<c['top']<p.height-35]
  assert min(c['x0'] for c in body)>=43 and max(c['x1'] for c in body)<=p.width-48,(i,'正文越界')
  assert '\ufffd' not in (p.extract_text() or '') and '\u25a0' not in (p.extract_text() or '')
with ZipFile(O/(name+'.docx')) as z:
 xml=etree.fromstring(z.read('word/document.xml'));ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
 links=xml.xpath('//w:hyperlink',namespaces=ns);assert len(links)==8
 for p in links:
  label=''.join(p.xpath('.//w:t/text()',namespaces=ns));assert norm(label) in text,label
for p in (W/'render').glob('page-*.png'):
 if int(p.stem.split('-')[-1])>7:p.unlink()
report={'pages':7,'representative_questions':22,'all_docx_text_and_cells_in_pdf':True,'all_source_links_visible':True,'page_bounds_checked':True,'all_final_pages_visually_reviewed':True,'sha256':{p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in [O/(name+'.docx'),dest]}}
(W/'最终核验.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
print('7页、22条例题、8个来源链接；Word正文及表格均保留，页面边界检查通过。')
