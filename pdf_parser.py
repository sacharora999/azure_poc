"""Extract detail rows from the nine Insource Extract Report PDF layouts."""
import io,re
from collections import defaultdict
import pdfplumber

NUMERIC_FIELDS = [
 'Net Sales','Net Invoice Amount','Invoice Amount','Non-Tax Amount',
 'Sales Tax','Use Tax','Bottle Tax','Bag Fee Tax','Cigarette Tax',
 'Tennessee Use Tax','F$B Tax','FB Amount','LG Amount','LQ Amount',
 'Retail Amount','RED Amount','RST Amount','UTT Amount','Retail B&O Sales',
 'Food and Beverage Tax','Empl. Meals','RED Tax','Service B&O Sales',
 'Service B&O Tax','Litter Tax Sales','Litter Tax']
# Longest first to avoid matching "Sales Tax" within other column names.
FIELD_WORDS={f: f.split() for f in NUMERIC_FIELDS}

def _num(v):
 s=str(v or '').replace(',','').replace('$','').strip()
 if s.endswith('-'):s='-'+s[:-1]
 if s.startswith('(') and s.endswith(')'):s='-'+s[1:-1]
 try:return float(s)
 except (TypeError,ValueError):return 0.0

def _line_words(page):
 lines=defaultdict(list)
 for w in page.extract_words(x_tolerance=1,y_tolerance=2):
  lines[round(w['top'],1)].append(w)
 return [sorted(v,key=lambda w:w['x0']) for _,v in sorted(lines.items())]

def _header_fields(words):
 tokens=[w['text'].lower() for w in words]
 found=[]
 for field,parts in sorted(FIELD_WORDS.items(),key=lambda kv:-len(kv[1])):
  pp=[x.lower() for x in parts]
  for i in range(len(tokens)-len(pp)+1):
   if tokens[i:i+len(pp)]==pp:
    found.append((i,i+len(pp),field,words[i]['x0']))
 found.sort(key=lambda x:(- (x[1]-x[0]),x[0]))
 selected=[];occupied=set()
 for a,b,f,x in found:
  if not any(j in occupied for j in range(a,b)):
   selected.append((x,f));occupied.update(range(a,b))
 return sorted(selected)

def parse_pdf_file(path,source,page_offset=0):
 result=[]
 with pdfplumber.open(path) as pdf:
  for pi,page in enumerate(pdf.pages):
   text=page.extract_text() or ''
   ent=re.search(r'Legal\s+Entity:\s*(U\d+)',text,re.I)
   folder=re.search(r'Page\s+State\s+Folder:\s*([A-Z]{2})',text,re.I)
   if not ent:raise ValueError(f'Missing legal entity on {source} page {page_offset+pi+1}')
   entity=ent.group(1).upper()
   page_state=folder.group(1).upper() if folder else None
   lines=_line_words(page)
   header=next((ws for ws in lines if ws and ws[0]['text']=='State' and any(w['text']=='County' for w in ws)),None)
   if not header:raise ValueError(f'Missing table header on {source} page {page_offset+pi+1}')
   fields=_header_fields(header)
   if not fields:raise ValueError(f'Missing numeric columns on {source} page {page_offset+pi+1}')
   first_x=fields[0][0]
   for words in lines:
    if not words or words[0]['x0']>100:continue
    row_state=words[0]['text'].strip().upper()
    if not re.fullmatch(r'[A-Z]{2}',row_state):continue
    if page_state and row_state!=page_state:continue
    if len(words)<7 or not any(re.fullmatch(r'\d{5}(?:-\d{4})?',w['text']) for w in words if 350 <= w['x0'] < 480):continue
    row={'Legal Entity':entity,'State':row_state,'State or Province Code':row_state,
         'Source PDF':source,'Source Page':page_offset+pi+1,'Source Row':len(result)+1}
    for idx,(x,field) in enumerate(fields):
     next_x=fields[idx+1][0] if idx+1<len(fields) else page.width+1
     cell=''.join(w['text'] for w in words if x-3<=w['x0']<next_x-3)
     row[field]=_num(cell)
    result.append(row)
 return result


def parse_pdf_bytes(data, source, page_offset=0):
 """Backward-compatible bytes API; optimized path uses parse_pdf_file."""
 import tempfile, os
 with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as stream:
  stream.write(data)
  path=stream.name
 try:
  return parse_pdf_file(path, source, page_offset)
 finally:
  os.unlink(path)
