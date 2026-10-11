import io, re
from collections import defaultdict
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

MAIN_REQUIRED = {
    'Account Description','Account #','State','LegEnt','Beginning Balance','Lag',
    'JEs & Payments','Add.Entity Activity','Extract','Lag Ledger Activity',
    'Ledger Activity','AE Tracking'
}
SMW_REQUIRED = {'Tax Code','Taxable Amount','Tax Amount','Extended Price','Division'}




def _num(v):
    if v is None or v == "": return 0.0
    if isinstance(v, (int,float)): return float(v)
    s=str(v).replace(",","").replace("$","").strip()
    if s.startswith("(") and s.endswith(")"): s="-"+s[1:-1]
    try: return float(s)
    except: return 0.0

def _norm(v): return str(v or "").strip().upper()

def _selected_states(user_inputs):
    """Return normalized selected states while keeping backward compatibility with single 'state'."""
    raw = user_inputs.get('states', user_inputs.get('state'))
    if isinstance(raw, str):
        # Accept legacy single state plus comma-separated input.
        values = raw.split(',')
    elif isinstance(raw, (list, tuple, set)):
        values = raw
    else:
        values = [raw] if raw is not None else []
    states = []
    for value in values:
        state = _norm(value)
        if state and state not in states:
            states.append(state)
    return states or ['ID']

def _rows(ws):
    if ws.title == 'Blackline Monthly Balance':
        r2=[c.value for c in ws[2]]; r3=[c.value for c in ws[3]]
        hdr=[]
        for a,b in zip(r2,r3):
            parts=[str(x).strip() for x in (a,b) if x is not None and str(x).strip()]
            hdr.append(" ".join(parts))
        data=list(ws.iter_rows(min_row=4,values_only=True))
    elif ws.title == 'Data':
        hdr=[str(c.value).strip() if c.value is not None else '' for c in ws[2]]
        data=list(ws.iter_rows(min_row=3,values_only=True))
    else:
        it=ws.iter_rows(values_only=True)
        hdr=[str(x).strip() if x is not None else '' for x in next(it)]
        data=list(it)
    data=[r for r in data if any(v is not None and str(v).strip() for v in r)]
    return hdr,[dict(zip(hdr,r)) for r in data]

def validate_sheet(ws, required):
    hdr,_=_rows(ws)
    missing=required-set(hdr)
    if missing: raise ValueError(f"{ws.title}: missing required columns: {sorted(missing)}")

def parse_two_excels(main_bytes, smw_bytes):
    """Load exactly two workbooks; PDF tax details are handled separately."""
    main_wb = load_workbook(io.BytesIO(main_bytes), data_only=True, read_only=True)
    smw_wb = load_workbook(io.BytesIO(smw_bytes), data_only=True, read_only=True)
    if 'Blackline Monthly Balance' not in main_wb.sheetnames:
        raise ValueError('Main Account file missing Blackline Monthly Balance tab')
    if 'SMW Data' not in smw_wb.sheetnames:
        raise ValueError('SMW file missing SMW Data tab')
    validate_sheet(main_wb['Blackline Monthly Balance'], MAIN_REQUIRED)
    validate_sheet(smw_wb['SMW Data'], SMW_REQUIRED)
    raw = {}
    for ws in main_wb.worksheets:
        raw[f'main::{ws.title}'] = _rows(ws)
    for ws in smw_wb.worksheets:
        raw[f'smw::{ws.title}'] = _rows(ws)
    return raw


# Financial amount fields from the PDF parser. Identifiers (site, cost center,
# customer number, ZIP code) are deliberately not included.
PDF_AMOUNT_FIELDS = (
    'Net Sales', 'Net Invoice Amount', 'Invoice Amount', 'Non-Tax Amount',
    'Sales Tax', 'Use Tax', 'Bottle Tax', 'Bag Fee Tax', 'Cigarette Tax',
    'Tennessee Use Tax', 'F$B Tax', 'FB Amount', 'LG Amount', 'LQ Amount',
    'Retail Amount', 'RED Amount', 'RST Amount', 'UTT Amount',
    'Retail B&O Sales', 'Food and Beverage Tax', 'Empl. Meals',
    'RED Tax', 'Service B&O Sales', 'Service B&O Tax',
    'Litter Tax Sales', 'Litter Tax',
)

def _pdf_tax_breakdown(rows):
    """Return all PDF financial columns with nonzero entity/state net totals.

    The legacy metrics key is retained for compatibility with the middle block.
    """
    # One scan over detail rows rather than one full scan per amount field.
    totals = defaultdict(float)
    allowed = set(PDF_AMOUNT_FIELDS)
    for row in rows:
        for field, value in row.items():
            if field in allowed:
                totals[field] += _num(value)
    return {field: amount for field in PDF_AMOUNT_FIELDS
            if field in totals and (amount := round(totals[field], 2)) != 0}


def reconcile_two_inputs(raw, pdf_rows, user_inputs):
    """Reconcile Main Account + SMW with PDF detail rows, without a third workbook.

    The PDF replaces the previous ID source. PDF tax metrics are keyed by
    (state, legal entity) to prevent assigning a whole-state total to each entity.
    Extract metrics are derived from entity-matched PDF detail rows.
    """
    _, black = raw['main::Blackline Monthly Balance']
    _, smw = raw['smw::SMW Data']
    user_inputs = user_inputs or {}
    states = _selected_states(user_inputs)
    selected = set(states)
    period = user_inputs.get('period') or 'July 2026'
    main = []
    accounts = defaultdict(list)
    for original in black:
        if _norm(original.get('State')) not in selected:
            continue
        r = dict(original)
        r['Open Item'] = round(
            _num(r.get('Open Items')) if r.get('Open Items') is not None else
            _num(r.get('Beginning Balance')) + _num(r.get('Lag')) +
            _num(r.get('Add.Entity Activity')) + _num(r.get('JEs & Payments')), 2)
        r['Actual Open Items'] = round(
            _num(r.get('Actual Open Items')) if r.get('Actual Open Items') is not None
            else r['Open Item'], 2)
        main.append(r)
        entity = _norm(r.get('LegEnt'))
        if entity:
            r['LegEnt'] = entity
            accounts[entity].append(r)

    smw_by_entity = defaultdict(list)
    for r in smw:
        division = _norm(r.get('Division'))
        if division:
            smw_by_entity[division].append(r)

    # Accept either legacy detail rows or pre-aggregated financial totals.
    # Aggregated mode never reconstructs PDF transaction lists.
    aggregated = isinstance(pdf_rows, dict) and pdf_rows.get('__aggregated__') is True
    pdf_by_key = defaultdict(list)
    pdf_sums = defaultdict(lambda: defaultdict(float))
    if aggregated:
        for item in pdf_rows['entities']:
            state, entity = item['state'], item['entity']
            if state in selected:
                for field, amount in item['amounts'].items():
                    pdf_sums[(state, entity)][field] += float(amount)
    else:
        for r in pdf_rows or []:
            state = _norm(r.get('State or Province Code') or r.get('State'))
            entity = _norm(r.get('Legal Entity'))
            if state in selected and entity:
                pdf_by_key[(state, entity)].append(r)

    def total(rows, field):
        return round(sum(_num(r.get(field)) for r in rows), 2)

    details = {}
    # A single entity worksheet consolidates ALL requested states, including
    # states found only in PDFs (without a matching Main Account row).
    entity_keys = set(accounts) | {entity for (state, entity) in (pdf_sums if aggregated else pdf_by_key)}
    for entity in sorted(entity_keys):
        acct = accounts.get(entity, [])
        entity_states = ({_norm(r.get('State')) for r in acct}
                         | {state for (state, ent) in (pdf_sums if aggregated else pdf_by_key) if ent == entity})
        pdf = [] if aggregated else [r for st in sorted(selected) for r in pdf_by_key.get((st, entity), [])]
        sums = defaultdict(float)
        if aggregated:
            for st in selected:
                for field, amount in pdf_sums.get((st, entity), {}).items():
                    sums[field] += amount
        def pdf_total(field):
            return round(sums[field], 2) if aggregated else total(pdf, field)
        pdf_breakdown = ({field: round(sums[field], 2) for field in PDF_AMOUNT_FIELDS
                          if round(sums[field], 2) != 0} if aggregated
                         else _pdf_tax_breakdown(pdf))
        sm = smw_by_entity.get(_norm(entity), [])
        sales = [r for r in acct if _norm(r.get('Account Description')) == 'SALES TAX-DIRECT POSTING']
        use = [r for r in acct if 'USE TAX' in _norm(r.get('Account Description'))]
        details[entity] = {
            'state': ', '.join(sorted(entity_states)), 'states': sorted(entity_states),
            'period': period, 'accounts': acct,
            'sales_accounts': sales, 'use_accounts': use,
            'id_rows': [], 'smw_rows': sm, 'pdf_rows': pdf,
            'metrics': {
                'Extract Net Sales': pdf_total('Net Sales'),
                'Extract Sales Tax': pdf_total('Sales Tax'),
                'Extract Use Tax': pdf_total('Use Tax'),
                'PDF Tax Breakdown': pdf_breakdown,
                'SMW Sales': total(sm, 'Extended Price'),
                'SMW Tax': total(sm, 'Tax Amount'),
                'Vertex Sales': pdf_total('Net Sales'),
                'Vertex Use Tax Sales': pdf_total('Use Tax'),
                'Vertex Tax': round(pdf_total('Sales Tax') + pdf_total('Use Tax'), 2)
            }
        }
    return main, details

def _period_text(period):
    if str(period).strip().lower()=='july 2026': return 'PERIOD 11, FY2026 JULY 2026'
    return str(period).upper()

def _sum(rows,key): return round(sum(_num(r.get(key)) for r in rows),2)
def _money(cell):
    cell.number_format='#,##0.00;(#,##0.00);-'
def _safe(name,used):
    b=re.sub(r'[:\\/?*\[\]]','_',str(name))[:31] or 'Entity'; n=b; i=1
    while n in used:
        suf=f'_{i}'; n=b[:31-len(suf)]+suf; i+=1
    used.add(n); return n

def _account(rows, fallback):
    vals=[str(r.get('Account #')).strip() for r in rows if r.get('Account #')]
    return vals[0] if vals else fallback

def _put(sh,cell,value,bold=False,color=None,italic=False,align=None,size=10):
    sh[cell]=value
    sh[cell].font=Font(name='Arial',size=size,bold=bold,italic=italic,color=color or '000000')
    if align: sh[cell].alignment=Alignment(horizontal=align,vertical='center',wrap_text=True)
    if isinstance(value,(int,float)): _money(sh[cell])

def _merge_title(sh,rng,value,size=11,color='000000'):
    sh.merge_cells(rng); c=sh[rng.split(':')[0]]
    c.value=value; c.font=Font(name='Arial',size=size,bold=True,color=color)
    c.alignment=Alignment(horizontal='center',vertical='center')


def _excel_num(v):
    """Excel-safe numeric literal."""
    return float(_num(v))

def _main_row_map(sh):
    """
    Map (Legal Entity, Account Description) -> output row(s) in Main Account Recon.
    Used to create traceable cross-sheet formulas in entity tabs.
    """
    hdr = {sh.cell(1,c).value: c for c in range(1, sh.max_column+1)}
    result = defaultdict(list)
    ent_col = hdr.get('LegEnt')
    desc_col = hdr.get('Account Description')
    if not ent_col or not desc_col:
        return result
    for r in range(2, sh.max_row+1):
        ent = str(sh.cell(r,ent_col).value or '').strip()
        desc = str(sh.cell(r,desc_col).value or '').strip()
        result[(ent, desc)].append(r)
    return result

def _set_formula(sh, cell, formula, bold=False, color=None, align='right'):
    sh[cell] = formula
    sh[cell].font = Font(name='Arial', size=10, bold=bold, color=color or '000000')
    sh[cell].alignment = Alignment(horizontal=align, vertical='center', wrap_text=True)
    _money(sh[cell])

def _sum_main_formula(main_ws, rows, header, headers=None):
    """
    Formula referencing Main Account Recon for a specific set of rows/header.
    """
    headers = headers or {main_ws.cell(1,c).value:c for c in range(1,main_ws.max_column+1)}
    col = headers.get(header)
    if not col or not rows:
        return "=0"
    letter = get_column_letter(col)
    refs = [f"'Main Account Recon'!{letter}{r}" for r in rows]
    return "=" + "+".join(refs)

def _left_block(sh,d,main_ws,entity,row_cache,headers):
    sales=d['sales_accounts']; use=d['use_accounts']; m=d['metrics']

    _merge_title(sh,'A1:C1','FINAL',10)
    _merge_title(sh,'A2:C2',d['state'],12,'1F4E78')
    _merge_title(sh,'A4:C4',_period_text(d['period']),11,'1F4E78')
    _merge_title(sh,'A5:C5','Monthly',10,'1F4E78')

    sales_rows=row_cache.get(entity,{}).get('sales',[])
    use_rows=row_cache.get(entity,{}).get('use',[])

    # ---------------- SALES TAX ----------------
    _merge_title(sh,'A6:C6','44570829 SALES TAX',11)

    _put(sh,'A8','BEGINNING BALANCE:',bold=True)
    _set_formula(sh,'C8',_sum_main_formula(main_ws,sales_rows,'Beginning Balance',headers),bold=True)

    _put(sh,'A9','Prior Month Lag',italic=True)
    _set_formula(sh,'C9',_sum_main_formula(main_ws,sales_rows,'Lag',headers),color='4472C4')

    _put(sh,'A10','Add. Entity Activity',italic=True)
    _set_formula(sh,'C10',_sum_main_formula(main_ws,sales_rows,'Add.Entity Activity',headers),color='4472C4')

    _put(sh,'A11','JEs & Payments',italic=True)
    _set_formula(sh,'C11',_sum_main_formula(main_ws,sales_rows,'JEs & Payments',headers),color='4472C4')

    _put(sh,'A12','OPEN ITEMS:',bold=True)
    _set_formula(sh,'C12','=SUM(C8:C11)',bold=True)

    _put(sh,'A14','Extract SAP',italic=True)
    _set_formula(sh,'C14',_sum_main_formula(main_ws,sales_rows,'Extract',headers),color='4472C4')

    _put(sh,'A15','SMW Report Taxes',italic=True)
    # SMW is external to Main Account Recon; put source-derived numeric value here.
    _put(sh,'C15',m['SMW Tax'],color='4472C4',align='right')

    _put(sh,'A17','CURRENT LAG:',bold=True)
    _set_formula(sh,'C17','=SUM(C14:C15)',bold=True)

    _put(sh,'A18','ENDING BALANCE:',bold=True)
    _set_formula(sh,'C18','=C12+C14+C15+C17',bold=True)

    _merge_title(sh,'A20:C20','OPEN ITEM SUMMARY',11,'C00000')
    _put(sh,'A21','Total Open Sales Tax Items-See Details',size=9)
    # Open-item detail total is calculated from the right-side detail amounts.
    _set_formula(sh,'C21','=SUM(L6:L16)',bold=True)
    _put(sh,'A22','Difference',italic=True,size=9)
    _set_formula(sh,'C22','=C21-C12',bold=True,color='C00000')

    # ---------------- USE TAX ----------------
    _merge_title(sh,'A24:C24','44570986 USE TAX',11)

    _put(sh,'A25','BEGINNING BALANCE:',bold=True)
    _set_formula(sh,'C25',_sum_main_formula(main_ws,use_rows,'Beginning Balance',headers),bold=True)

    _put(sh,'A26','Prior Month Lag',italic=True)
    _set_formula(sh,'C26',_sum_main_formula(main_ws,use_rows,'Lag',headers),color='4472C4')

    _put(sh,'A27','Add. Entity Activity',italic=True)
    _set_formula(sh,'C27',_sum_main_formula(main_ws,use_rows,'Add.Entity Activity',headers),color='4472C4')

    _put(sh,'A28','JEs & Payments',italic=True)
    _set_formula(sh,'C28',_sum_main_formula(main_ws,use_rows,'JEs & Payments',headers),color='4472C4')

    _put(sh,'A29','OPEN ITEMS:',bold=True)
    _set_formula(sh,'C29','=SUM(C25:C28)',bold=True)

    _put(sh,'A31','Extract SAP',italic=True)
    _set_formula(sh,'C31',_sum_main_formula(main_ws,use_rows,'Extract',headers),color='4472C4')

    _put(sh,'A32','Extract Overhead/AE Track',italic=True)
    _set_formula(sh,'C32',_sum_main_formula(main_ws,use_rows,'AE Tracking',headers),color='4472C4')

    _put(sh,'A34','CURRENT LAG:',bold=True)
    _set_formula(sh,'C34','=SUM(C31:C32)',bold=True)

    _put(sh,'A35','ENDING BALANCE:',bold=True)
    _set_formula(sh,'C35','=C29+C31+C32+C34',bold=True)

    _merge_title(sh,'A37:C37','OPEN ITEM SUMMARY',11,'C00000')
    _put(sh,'A38','Total Open Use Tax Items-See Details',size=9)
    # Right-side use-tax detail can start at a dynamic row. The total row in column L
    # is linked after _right_block writes it; temporary formula is replaced there.
    _put(sh,'C38',0.0,bold=True,align='right')
    _put(sh,'A39','Difference',italic=True,size=9)
    _set_formula(sh,'C39','=C38-C29',bold=True,color='C00000')

    # ---------------- CURRENT PAYMENT CALCULATION ----------------
    _merge_title(sh,'A41:C41','CURRENT PAYMENT CALCULATION',11)

    _put(sh,'A42','Sales Tax',bold=True)
    _set_formula(sh,'C42','=C14+C15',bold=True)

    _put(sh,'A43','Use Tax',bold=True)
    _set_formula(sh,'C43','=C31+C32')

    _put(sh,'A44','Use Tax U447',italic=True); _put(sh,'C44',0.0,align='right')
    _put(sh,'A45','Use Tax U101',italic=True); _put(sh,'C45',0.0,align='right')
    _put(sh,'A46','Hold Neg U100 CC 080484 Cajon Name',italic=True); _put(sh,'C46',0.0,align='right')

    _put(sh,'A47','Total Liability',bold=True)
    _set_formula(sh,'C47','=SUM(C42:C46)',bold=True)

    _put(sh,'A48','Discount',bold=True)
    _put(sh,'C48',0.0,align='right')

    _put(sh,'A49','Return',bold=True)
    _set_formula(sh,'C49','=C47-C48')

    _put(sh,'A50','Delta',bold=True)
    _set_formula(sh,'C50','=C47-C48',bold=True)

def _middle_block(sh,d,main_ws,entity,row_cache,headers):
    m=d['metrics']
    _merge_title(sh,'E1:G1',d['state'],12,'1F4E78')
    _merge_title(sh,'E3:G3',_period_text(d['period']),11,'1F4E78')
    _merge_title(sh,'E4:G4','Monthly',10,'1F4E78')
    _merge_title(sh,'E5:G5','Gross Sales Reconciliation',11)
    # Replace the five former sales rows with all nonzero PDF financial amount categories.
    # Only nonzero entity/state totals are displayed, retaining original PDF labels.
    tax_items = list(m.get('PDF Tax Breakdown', {}).items())
    for rr, (tax_name, amount) in enumerate(tax_items, start=6):
        _put(sh, f'E{rr}', tax_name, size=9)
        _put(sh, f'G{rr}', amount, align='right')
    # Expand only columns E:G; never insert worksheet rows globally because
    # that would move the independent left/right reconciliation blocks.
    shift = max(0, len(tax_items) - 5)
    def mr(original_row):
        return original_row + shift if original_row >= 11 else original_row
    total_row = 11 + shift
    _put(sh, f'E{total_row}', 'Total PDF Amounts', bold=True)
    _set_formula(sh, f'G{total_row}', f'=SUM(G6:G{total_row-1})', bold=True)
    _put(sh, f'E{mr(13)}', 'Vertex Sales')
    _put(sh, f'G{mr(13)}', m['Vertex Sales'], align='right')
    _put(sh, f'E{mr(14)}', 'Vertex Use Tax Sales')
    _put(sh, f'G{mr(14)}', m['Vertex Use Tax Sales'], align='right')
    _put(sh, f'E{mr(15)}', 'Total Vertex Sales', bold=True)
    _set_formula(sh, f'G{mr(15)}', f'=SUM(G{mr(13)}:G{mr(14)})', bold=True)
    _put(sh, f'E{mr(17)}', 'Difference', bold=True)
    _set_formula(sh, f'G{mr(17)}', f'=G{total_row}-G{mr(15)}', bold=True)
    _put(sh, f'E{mr(18)}', 'Explanation of difference:', size=9)
    sh.merge_cells(f'F{mr(18)}:G{mr(18)}')
    _put(sh, f'F{mr(18)}', '', italic=True, size=9)
    sh.merge_cells(f'E{mr(19)}:G{mr(19)}')
    _put(sh, f'E{mr(19)}', 'Special Note: Use Tax is back calculated for extract amounts', size=9)
    sh.merge_cells(f'E{mr(20)}:G{mr(20)}')
    _put(sh, f'E{mr(20)}', 'Difference due to back calculating', size=9)
    _merge_title(sh, f'E{mr(22)}:G{mr(22)}', 'CHECK-SIB 0.00', 10)
    _merge_title(sh, f'E{mr(23)}:G{mr(23)}', 'SALES TAX', 10)
    entity_rows=row_cache.get(entity,{})
    sales_rows=entity_rows.get('sales',[])
    use_auto_rows=entity_rows.get('use_auto',[])

    def refs(rows, header):
        col=headers.get(header)
        if not rows or not col:
            return "0"
        letter=get_column_letter(col)
        return "+".join(f"'Main Account Recon'!{letter}{rr}" for rr in rows)

    _set_formula(sh,f'E{mr(24)}',f'=C12-({refs(sales_rows,"Actual Open Items")})'); _put(sh,f'F{mr(24)}','OPEN ITEMS')
    _set_formula(sh,f'E{mr(25)}',f'=C17-({refs(sales_rows,"Extract")})'); _put(sh,f'F{mr(25)}','CURRENT LIABILITY')
    _set_formula(sh,f'E{mr(26)}',f'=C18-({refs(sales_rows,"Ending Balance")})'); _put(sh,f'F{mr(26)}','ENDING BALANCE')

    _merge_title(sh,f'E{mr(27)}:G{mr(27)}','USE TAX',10)
    _set_formula(sh,f'E{mr(28)}',f'=C29-({refs(use_auto_rows,"Actual Open Items")})'); _put(sh,f'F{mr(28)}','ENDING BALANCE')
    _set_formula(sh,f'E{mr(29)}',f'=C34-({refs(use_auto_rows,"Extract")})'); _put(sh,f'F{mr(29)}','CURRENT LIABILITY')
    _set_formula(sh,f'E{mr(30)}',f'=C35-({refs(use_auto_rows,"Ending Balance")})'); _put(sh,f'F{mr(30)}','ENDING BALANCE')

    for r in list(range(24,27))+list(range(28,31)):
        sh[f'E{mr(r)}'].fill=PatternFill('solid',fgColor='E2F0D9')

def _right_block(sh,d,main_ws,entity,row_cache,headers):
    _merge_title(sh,'J1:L1',d['state'],12,'1F4E78')
    _merge_title(sh,'J3:L3',_period_text(d['period']),11,'1F4E78')
    _merge_title(sh,'J4:L4','OPEN ITEM DETAIL',10,'1F4E78')
    _merge_title(sh,'J5:L5','44570829 SALES TAX',11)

    actual_col=headers['Actual Open Items']
    actual_letter=get_column_letter(actual_col)
    entity_rows=row_cache.get(entity,{})
    sales_source=entity_rows.get('sales_source',[])
    use_source=entity_rows.get('use_source',[])

    r=6
    for rr,desc in sales_source:
        source=d['accounts']
        match=next((x for x in source if _norm(x.get('Account Description'))==_norm(desc)),{})
        note=str(match.get('Notes') or '').strip()
        resp=str(match.get('Responsible Person') or '').strip()
        detail=note or (desc + (f" - {resp}" if resp else ""))
        sh.merge_cells(start_row=r,start_column=10,end_row=r,end_column=11)
        _put(sh,f'J{r}',detail,size=9)
        _set_formula(sh,f'L{r}',f"='Main Account Recon'!{actual_letter}{rr}")
        r+=1

    sales_total_row=max(r,17)
    _merge_title(sh,f'J{sales_total_row}:K{sales_total_row}','Total Open Sales Tax Items:',10)
    if r>6:
        _set_formula(sh,f'L{sales_total_row}',f'=SUM(L6:L{r-1})',bold=True)
        _set_formula(sh,'C21',f'=L{sales_total_row}',bold=True)
    else:
        _put(sh,f'L{sales_total_row}',0.0,bold=True,align='right')
        _set_formula(sh,'C21',f'=L{sales_total_row}',bold=True)

    ur=max(r+2,19)
    _merge_title(sh,f'J{ur}:L{ur}','44570986 USE TAX',11)
    use_start=ur+1
    ur=use_start
    for rr,desc in use_source:
        match=next((x for x in d['accounts'] if _norm(x.get('Account Description'))==_norm(desc)),{})
        note=str(match.get('Notes') or '').strip()
        resp=str(match.get('Responsible Person') or '').strip()
        detail=note or (desc + (f" - {resp}" if resp else ""))
        sh.merge_cells(start_row=ur,start_column=10,end_row=ur,end_column=11)
        _put(sh,f'J{ur}',detail,size=9)
        _set_formula(sh,f'L{ur}',f"='Main Account Recon'!{actual_letter}{rr}")
        ur+=1

    total_row=max(ur+8,34)
    _merge_title(sh,f'J{total_row}:K{total_row}','Total Open Use Tax Items:',10)
    if ur>use_start:
        _set_formula(sh,f'L{total_row}',f'=SUM(L{use_start}:L{ur-1})',bold=True)
    else:
        _put(sh,f'L{total_row}',0.0,bold=True,align='right')
    _set_formula(sh,'C38',f'=L{total_row}',bold=True)

def _style(sh):
    sh.sheet_view.showGridLines=False
    widths={'A':31,'B':14,'C':17,'D':4,'E':27,'F':23,'G':18,'H':4,'I':4,'J':54,'K':32,'L':18}
    for c,w in widths.items(): sh.column_dimensions[c].width=w
    for r in range(1,51): sh.row_dimensions[r].height=18
    sh.row_dimensions[1].height=20; sh.row_dimensions[4].height=20
    black=Side(style='medium',color='000000'); thin=Side(style='thin',color='808080')
    # Main boxes/separators visible in reference video.
    for r in range(1,51):
        sh[f'C{r}'].border=Border(right=black)
        sh[f'G{r}'].border=Border(right=black)
    for rng in [(1,1,3),(6,1,3),(18,1,3),(20,1,3),(22,1,3),(24,1,3),(35,1,3),(37,1,3),(39,1,3),(41,1,3),(47,1,3),(50,1,3),
                (1,5,7),(5,5,7),(17,5,7),(22,5,7),(30,5,7),(1,10,12),(4,10,12),(5,10,12)]:
        rr,c1,c2=rng
        for c in range(c1,c2+1):
            old=sh.cell(rr,c).border
            sh.cell(rr,c).border=Border(left=old.left,right=old.right,top=old.top,bottom=black)
    sh.page_setup.orientation='landscape'; sh.page_setup.fitToWidth=1; sh.page_setup.fitToHeight=1
    sh.sheet_properties.pageSetUpPr.fitToPage=True
    sh.print_area=f'A1:L{max(50, sh.max_row)}'
    sh.freeze_panes='A6'

def build_output(main,details):
    wb=Workbook()
    ws=wb.active; ws.title='Main Account Recon'
    cols=['Account Description','Account #','State','LegEnt','Beginning Balance','Lag','JEs & Payments',
          'Open Item','Add.Entity Activity','Actual Open Items','Extract','Lag Ledger Activity',
          'Ledger Activity','AE Tracking','Ending Balance','Balance From FS10','Difference',
          'Future Postings','Responsible Person','Notes']
    ws.append(cols)
    for r in main: ws.append([r.get(c) for c in cols])
    for c in ws[1]:
        c.fill=PatternFill('solid',fgColor='1F4E78'); c.font=Font(color='FFFFFF',bold=True)
        c.alignment=Alignment(horizontal='center',wrap_text=True)
    ws.freeze_panes='A2'
    for i in range(1,len(cols)+1): ws.column_dimensions[get_column_letter(i)].width=18
    ws.column_dimensions['A'].width=48

    # Financial columns in Main Account Recon:
    # positive -> 1,234.56, negative -> (1,234.56), zero -> -
    financial_headers = {
        'Beginning Balance','Lag','JEs & Payments','Open Item',
        'Add.Entity Activity','Actual Open Items','Extract',
        'Lag Ledger Activity','Ledger Activity','AE Tracking',
        'Ending Balance','Balance From FS10','Difference','Future Postings'
    }
    header_to_col = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}
    for header in financial_headers:
        col = header_to_col.get(header)
        if col:
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row, col)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00;(#,##0.00);-'

    # Cache Main Account Recon row locations once for all entity sheets.
    headers={ws.cell(1,c).value:c for c in range(1,ws.max_column+1)}
    ent_col=headers['LegEnt']; desc_col=headers['Account Description']
    row_cache=defaultdict(lambda: {'sales':[], 'use':[], 'use_auto':[], 'sales_source':[], 'use_source':[]})
    for rr in range(2,ws.max_row+1):
        ent=str(ws.cell(rr,ent_col).value or '').strip()
        desc=_norm(ws.cell(rr,desc_col).value)
        if desc=='SALES TAX-DIRECT POSTING':
            row_cache[ent]['sales'].append(rr)
            row_cache[ent]['sales_source'].append((rr, str(ws.cell(rr,desc_col).value or '').strip()))
        if 'USE TAX' in desc:
            row_cache[ent]['use'].append(rr)
            row_cache[ent]['use_source'].append((rr, str(ws.cell(rr,desc_col).value or '').strip()))
        if desc=='USE TAX-ACCRUAL-AUTOMATIC': row_cache[ent]['use_auto'].append(rr)

    used={'Main Account Recon'}
    for ent,d in details.items():
        sh=wb.create_sheet(_safe(ent,used))
        _left_block(sh,d,ws,ent,row_cache,headers); _middle_block(sh,d,ws,ent,row_cache,headers); _right_block(sh,d,ws,ent,row_cache,headers); _style(sh)
        sh['A3']=f'LEGAL ENTITY: {ent}'; sh.merge_cells('A3:C3')
        sh['A3'].font=Font(name='Arial',bold=True,color='7F6000',size=9)
        sh['A3'].alignment=Alignment(horizontal='center')
    bio=io.BytesIO(); wb.save(bio); return bio.getvalue()
