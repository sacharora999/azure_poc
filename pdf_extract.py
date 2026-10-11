"""DI layout-table adapter with fail-closed handling for unknown report families.

Unrecognized layouts are rejected rather than silently returning zero rows.
"""
import re, json, os, io, time
from collections import defaultdict
from processor import PDF_AMOUNT_FIELDS
from pdf_parser import _num


ALIASES={'state or province code':'State','state':'State','legal entity':'Legal Entity',
         'retail b&o sales':'Retail B&O Sales','fb tax':'F$B Tax',
         'f&b tax':'F$B Tax','employee meals':'Empl. Meals',
         'empl. meals':'Empl. Meals'}
AMOUNTS={re.sub(r'[^a-z0-9]+','',x.lower()):x for x in PDF_AMOUNT_FIELDS}
STATES=set('AL AK AZ AR CA CO CT DE FL GA HI ID IL IN IA KS KY LA ME MD MA MI MN MS MO MT NE NV NH NJ NM NY NC ND OH OK OR PA RI SC SD TN TX UT VT VA WA WV WI WY DC'.split())

def norm_header(s):
    key=re.sub(r'\s+',' ',str(s or '').strip().lower())
    if key in ALIASES:return ALIASES[key]
    return AMOUNTS.get(re.sub(r'[^a-z0-9]+','',key))

def extract_rows(result,source,page_offset=0):
    """Extract rows from DI tables; requires explicit state and legal entity provenance.
    Page-level legal entity and state are read from page text, not assumed across files.
    """
    pages=result.get('pages',[])
    page_context={}
    for p in pages:
        page=int(p['pageNumber'])
        content='\n'.join(x.get('content','') for x in p.get('lines',[]))
        ent=re.search(r'Legal\s+Entity\s*:\s*([A-Za-z]\d+)',content,re.I)
        folder=re.search(r'Page\s+State\s+Folder\s*:\s*([A-Z]{2})',content,re.I)
        page_context[page]=(ent.group(1).upper() if ent else None,folder.group(1).upper() if folder else None)
    rows=[]; tables_seen=0; accepted_tables=0; rejected=[]
    for ti,table in enumerate(result.get('tables',[])):
        tables_seen+=1
        cells=table.get('cells',[])
        if not cells:continue
        grid=defaultdict(dict)
        for c in cells:
            grid[int(c['rowIndex'])][int(c['columnIndex'])]=str(c.get('content') or '').strip()
        header_idx=None; headers={}
        for ri in sorted(grid)[:4]:
            candidate={col:norm_header(val) for col,val in grid[ri].items()}
            if 'State' in candidate.values() and any(v in PDF_AMOUNT_FIELDS for v in candidate.values()):
                header_idx=ri;headers=candidate;break
        if header_idx is None:
            rejected.append(f'table {ti}: missing recognized State and financial headers')
            continue
        accepted_tables+=1
        page=int((table.get('boundingRegions') or [{}])[0].get('pageNumber') or 1)
        entity,folder=page_context.get(page,(None,None))
        for ri in sorted(grid):
            if ri<=header_idx:continue
            vals=grid[ri]
            row={field:vals.get(col,'') for col,field in headers.items() if field}
            state=str(row.get('State') or folder or '').strip().upper()
            ent=str(row.get('Legal Entity') or entity or '').strip().upper()
            if state not in STATES:continue  # excludes city/county/state subtotal rows
            if not re.fullmatch(r'[A-Za-z]\d+',ent):
                raise ValueError(f'{source} page {page_offset+page}: legal entity not found; review layout')
            if folder and folder!=state:
                raise ValueError(f'{source} page {page_offset+page}: state mismatch {folder}/{state}')
            amounts={f:_num(row.get(f)) for f in PDF_AMOUNT_FIELDS if f in row}
            if not amounts:continue
            rows.append({'Legal Entity':ent,'State':state,'State or Province Code':state,
                         'Source PDF':source,'Source Page':page_offset+page,'Source Row':len(rows)+1,**amounts})
    if not accepted_tables or not rows:
        raise ValueError(f'{source}: unsupported DI table layout or zero detail rows; {rejected[:5]}')
    # Fail closed if DI identified additional tables that may contain unparsed transactions.
    if rejected and os.getenv('ALLOW_UNRECOGNIZED_TABLES','false').lower()!='true':
        raise ValueError(f'{source}: {len(rejected)} unrecognized tables require layout review: {rejected[:5]}')
    return rows, {'tables':tables_seen,'acceptedTables':accepted_tables,'rejectedTables':rejected}

def analyze(data):
    import requests
    endpoint=os.environ['DOCUMENT_INTELLIGENCE_ENDPOINT'].rstrip('/')
    api_version=os.getenv('DOCUMENT_INTELLIGENCE_API_VERSION','2024-11-30')
    model=os.getenv('DOCUMENT_INTELLIGENCE_MODEL','prebuilt-layout')
    url=f'{endpoint}/documentintelligence/documentModels/{model}:analyze?api-version={api_version}'
    api_key=os.environ['DOCUMENT_INTELLIGENCE_KEY']
    session=requests.Session()
    response=session.post(url,headers={'Ocp-Apim-Subscription-Key':api_key,'Content-Type':'application/pdf'},data=data,timeout=(20,120))
    response.raise_for_status()
    poll=response.headers.get('Operation-Location')
    if not poll:raise RuntimeError('Document Intelligence did not return Operation-Location')
    deadline=time.monotonic()+int(os.getenv('DI_POLL_TIMEOUT_SECONDS','1200'))
    while time.monotonic()<deadline:
        r=session.get(poll,headers={'Ocp-Apim-Subscription-Key':api_key},timeout=(20,60))
        r.raise_for_status(); obj=r.json()
        if obj.get('status')=='succeeded':return obj['analyzeResult']
        if obj.get('status') in ('failed','canceled'):raise RuntimeError(f'DI analysis {obj.get("status")}: {obj.get("error")}')
        time.sleep(2)
    raise TimeoutError('Document Intelligence polling exceeded configured timeout')
