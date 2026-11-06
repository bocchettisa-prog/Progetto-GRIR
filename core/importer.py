"""SAP Excel import and matching. No source workbook is modified."""
from collections import Counter
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
import re
from openpyxl import load_workbook

STATUSES=('Still Open','Amount Changed - Still Unresolved','Resolved','To Be Reviewed')
ENTITY_RE=re.compile(r'(?<![A-Z0-9])((?:GB|IE|NL)\d{1,3})(?![A-Z0-9])',re.I)
PERIOD_RE=re.compile(r'(?<!\d)(0[1-9]|1[0-2])[.\-_](20\d\d|\d\d)(?!\d)')

def filename_metadata(filename):
    stem=Path(filename).stem.upper()
    entity=ENTITY_RE.search(stem)
    period=PERIOD_RE.search(stem)
    if not entity or not period: raise ValueError('Filename must include entity (GB12, NL11, IE11...) and period MM.YY, e.g. GB12 GRIR 09.26.xlsx.')
    year=int(period.group(2));year=2000+year if year<100 else year
    return entity.group(1).upper(),f'{year:04d}-{int(period.group(1)):02d}'

def minor(value):
    if value is None or str(value).strip()=='': return None
    if isinstance(value,str):
        v=value.strip().replace(' ','').replace('£','').replace('€','')
        if v.startswith('(') and v.endswith(')'): v='-'+v[1:-1]
        if ',' in v and '.' in v:
            v=v.replace('.','').replace(',','.') if v.rfind(',')>v.rfind('.') else v.replace(',','')
        elif ',' in v: v=v.replace(',','.')
        value=v
    return int((Decimal(str(value))*100).quantize(Decimal('1'),rounding=ROUND_HALF_UP))

def key_from_assignment(raw):
    token=re.sub(r'\D','',str(raw or ''))
    if len(token)<13 or not token.startswith(('45','47','49')): raise ValueError(f'Invalid SAP Assignment {raw!r}')
    return token[:10],token[-3:]

def classify(comment):
    """Legacy importer entrypoint: now returns the discrepancy REASON, not action."""
    from classification import classify_reason
    return classify_reason(comment)

def find_sap_sheet(wb):
    for ws in wb:
        h={str(ws.cell(1,j).value or '').strip().lower():j for j in range(1,ws.max_column+1)}
        if 'assignment' in h and 'amount in local currency' in h and 'local currency' in h:
            return ws,h
    raise ValueError('Cannot locate SAP sheet with Assignment, Amount in local currency and Local Currency columns.')

def import_sap(path):
    """Read all nonzero SAP outline subtotals and validate against SAP's grand total."""
    wb=load_workbook(path,read_only=False,data_only=True)
    ws,h=find_sap_sheet(wb)
    records=[];warnings=[];currencies=Counter()
    subtotal_count=zero=positive=negative=0
    subtotal_sum=0;invalid=[];grand_totals=[]
    amount_col=h['amount in local currency'];curr_col=h['local currency']
    assign_col=h['assignment'];doc_col=h.get('document number',4)
    for n in range(2,ws.max_row+1):
        level=ws.row_dimensions[n].outlineLevel
        amt=minor(ws.cell(n,amount_col).value)
        if amt is None: continue
        assignment=ws.cell(n,assign_col).value
        doc=ws.cell(n,doc_col).value
        currency=str(ws.cell(n,curr_col).value or '').strip().upper()
        if level==0 and n>2 and not str(assignment or '').strip() and not str(doc or '').strip() and currency:
            grand_totals.append((n,amt,currency))
            continue
        if level != 1 or str(doc or '').strip():continue
        subtotal_count+=1;subtotal_sum+=amt
        if amt==0:zero+=1;continue
        if amt>0:positive+=1
        else:negative+=1
        try: po,line=key_from_assignment(assignment)
        except ValueError as exc:
            invalid.append(f'SAP row {n}: {exc}');continue
        if not currency:
            invalid.append(f'SAP row {n}: missing local currency');continue
        currencies[currency]+=1
        records.append(dict(po=po,item_line=line,current_minor=amt,vendor_text='',source_row=n))
    if len(grand_totals)!=1:
        raise ValueError(f'Expected exactly one SAP grand total in final ungrouped rows; found {len(grand_totals)}. Import blocked for accounting safety.')
    total_row,grand_total,total_currency=grand_totals[0]
    if len(currencies)!=1 or next(iter(currencies),None)!=total_currency:
        raise ValueError(f'Currency mismatch: subtotals {dict(currencies)}, grand total {total_currency}.')
    duplicates=[(k,v) for k,v in Counter((r['po'],r['item_line']) for r in records).items() if v>1]
    if duplicates:raise ValueError(f'Duplicate PO + Item Line keys: {duplicates[:5]}. Import blocked.')
    if not records:raise ValueError('No nonzero SAP subtotal rows found.')
    imported_sum=sum(r['current_minor'] for r in records)
    difference=imported_sum-grand_total
    if invalid or subtotal_sum!=grand_total or difference!=0:
        sample='; '.join(invalid[:4]) if invalid else ''
        raise ValueError(f'SAP BALANCE VALIDATION FAILED. SAP total row {total_row}: {grand_total/100:,.2f} {total_currency}; all subtotals: {subtotal_sum/100:,.2f}; imported PO lines: {imported_sum/100:,.2f}; difference: {difference/100:,.2f}. Invalid rows: {len(invalid)}. {sample} Import blocked.')
    warnings.append(f'SAP balance VERIFIED at row {total_row}: {grand_total/100:,.2f} {total_currency}. {positive} positive and {negative} negative open positions included.')
    warnings.append('SAP subtotal rows have no Vendor/Text; known vendors are recovered from previous reconciliations, new items need review.')
    return dict(records=records,currency=total_currency,subtotal_count=subtotal_count,
       zero_count=zero,positive_count=positive,negative_count=negative,warnings=warnings,
       sheet=ws.title,grand_total_minor=grand_total,imported_total_minor=imported_sum,
       validation_difference_minor=difference,grand_total_row=total_row,validated=True)

def import_rec(path):
    wb=load_workbook(path,read_only=False,data_only=True)
    ws=None;headers={}
    for s in wb:
        for row in s.iter_rows(min_row=1,max_row=min(80,s.max_row)):
            items={str(c.value or '').strip().lower():c.column for c in row}
            if any('purchase order' in k for k in items) and any('items line' in k for k in items):
                ws=s;headers=items;start=row[0].row+1;break
        if ws:break
    if ws is None:raise ValueError('Cannot identify previous REC table')
    get=lambda name:next((v for k,v in headers.items() if name in k),None)
    cols={name:get(name) for name in ('purchase order','items line','amount in local currency','vendor','gl account','account assignment','posting date','comments','status')}
    if not all(cols[k] for k in ('purchase order','items line','amount in local currency')):raise ValueError('Incomplete REC columns')
    result={}; errors=[]
    for row in ws.iter_rows(min_row=start):
        vals={k:row[col-1].value if col else None for k,col in cols.items()}
        po=re.sub(r'\D','',str(vals['purchase order'] or ''))
        if len(po)!=10 or not po.startswith(('45','47','49')):continue
        line=re.sub(r'\D','',str(vals['items line'] or '')).zfill(3)[-3:]
        if not line.strip('0'):continue
        key=(po,line)
        if key in result:
            errors.append(f'Duplicate previous REC key {po}/{line}');continue
        value=minor(vals['amount in local currency'])
        result[key]=dict(po=po,item_line=line,current_minor=value or 0,vendor_text=str(vals['vendor'] or '').strip(),gl_account=str(vals['gl account'] or '').strip(),account_assignment=str(vals['account assignment'] or '').strip(),posting_date=str(vals['posting date'] or '').strip(),comment=str(vals['comments'] or '').strip(),match_status=('Resolved' if (value or 0)==0 else str(vals['status'] or '').strip()))
    if errors:raise ValueError('; '.join(errors[:5]))
    if not result:raise ValueError('No positions found in previous REC')
    return result

def reconcile(current,previous):
    """Previous is a dict[(PO,line)] to positions. Prior Resolved items never roll forward."""
    now={(r['po'],r['item_line']):r for r in current}
    if len(now)!=len(current):raise ValueError('Duplicate current keys')
    out=[]
    for key,r in now.items():
        old=previous.get(key)
        if old and old.get('match_status')!='Resolved':
            comment=str(old.get('comment') or '').strip()
            st=('To Be Reviewed' if not comment else
                'Still Open' if r['current_minor']==old['current_minor'] else
                'Amount Changed - Still Unresolved')
            details={k:old.get(k,'') for k in ('gl_account','account_assignment','posting_date','comment','action_override')}
            details['vendor_text']=r.get('vendor_text') or old.get('vendor_text','')
            out.append(dict(**details,po=key[0],item_line=key[1],current_minor=r['current_minor'],previous_minor=old['current_minor'],reason_category=old.get('reason_category') or classify(old.get('comment')),match_status=st))
        else:
            out.append(dict(po=key[0],item_line=key[1],current_minor=r['current_minor'],previous_minor=None,vendor_text=r.get('vendor_text',''),gl_account='',account_assignment='',posting_date='',comment='',reason_category='Unclassified',match_status='To Be Reviewed'))
    for key,old in previous.items():
        if key not in now and old.get('match_status')!='Resolved' and old.get('current_minor') != 0:
            out.append(dict(**{k:old.get(k,'') for k in ('po','item_line','vendor_text','gl_account','account_assignment','posting_date','comment','action_override')},current_minor=0,previous_minor=old['current_minor'],reason_category=old.get('reason_category') or classify(old.get('comment')),match_status='Resolved'))
    return out

def historical_snapshot_from_excel(path, expected_entity, expected_period, expected_currency):
    """Load full previous SAP subtotals and enrich with old REC notes."""
    entity, period=filename_metadata(path)
    if entity!=expected_entity or period>=expected_period:
        raise ValueError(f'Previous Excel must belong to {expected_entity} and precede {expected_period}; found {entity}/{period}.')
    previous_rec=import_rec(path)
    previous_sap=import_sap(path)
    if previous_sap['currency']!=expected_currency:
        raise ValueError(f'Previous Excel currency {previous_sap["currency"]} differs from current {expected_currency}.')
    previous_rows=reconcile(previous_sap['records'],{})
    enriched=[]
    for r in previous_rows:
        key=(r['po'],r['item_line'])
        old=previous_rec.get(key)
        if old:
            for col in ('vendor_text','gl_account','account_assignment','posting_date','comment'):
                r[col]=old.get(col,'') or r.get(col,'')
            r['reason_category']=classify(r['comment'])
            r['match_status']='Still Open' if r['comment'] else 'To Be Reviewed'
        enriched.append(r)
    return dict(entity=entity,period=period,currency=expected_currency,
                rows=enriched, rec_count=len(previous_rec),sap_count=len(previous_sap['records']),
                notes=previous_sap['warnings'], partial=len(previous_rec)<len(enriched),
                grand_total_minor=previous_sap['grand_total_minor'],imported_total_minor=previous_sap['imported_total_minor'])
