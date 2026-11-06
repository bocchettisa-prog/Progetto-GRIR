"""V12 read-only controller analytics. Monetary amounts are integer minor units.

Every calculation uses the saved current values and never updates accounting rows.
No inferred age or GL posting is substituted for a missing verified GR date.
"""
from collections import defaultdict
from calendar import monthrange
from datetime import date, datetime, timedelta
from classification import effective_reason, effective_action, is_potential_reversal

BUCKETS = ('0–30 days','31–60 days','61–90 days','91–180 days','Over 180 days','Unknown / invalid date')

def signed_summary(rows):
    active=[p for p in rows if int(p['current_minor']) != 0]
    return dict(count=len(active),net_minor=sum(int(p['current_minor']) for p in active),
                exposure_minor=sum(abs(int(p['current_minor'])) for p in active),
                reversal_minor=sum(abs(int(p['current_minor'])) for p in active if is_potential_reversal(p)),
                review_count=sum(p['match_status']=='To Be Reviewed' for p in active),
                positive_count=sum(int(p['current_minor'])>0 for p in active))

def month_end(period):
    year,month=map(int,period.split('-'))
    return date(year,month,monthrange(year,month)[1])

def parse_posting_date(value):
    if isinstance(value,datetime):return value.date()
    if isinstance(value,date):return value
    text=str(value or '').strip()
    if not text:return None
    for fmt in ('%Y-%m-%d','%d.%m.%Y','%d/%m/%Y','%d-%m-%Y','%Y/%m/%d','%Y.%m.%d'):
        try:return datetime.strptime(text.split(' ')[0],fmt).date()
        except ValueError:pass
    return None

def age_days(p,period):
    dt=parse_posting_date(p.get('posting_date'))
    if dt is None:return None
    age=(month_end(period)-dt).days
    return age if age>=0 else None

def aging(rows,period):
    bins={name:[] for name in BUCKETS}
    for p in rows:
        if not int(p['current_minor']):continue
        a=age_days(p,period)
        group=('Unknown / invalid date' if a is None else
               '0–30 days' if a<=30 else '31–60 days' if a<=60 else
               '61–90 days' if a<=90 else '91–180 days' if a<=180 else 'Over 180 days')
        bins[group].append(p)
    return [dict(bucket=name,**signed_summary(items)) for name,items in bins.items()]

def movement(current,previous):
    a={(p['po'],p['item_line']):int(p['current_minor']) for p in current if int(p['current_minor'])!=0}
    b={(p['po'],p['item_line']):int(p['current_minor']) for p in previous if int(p['current_minor'])!=0}
    new=sum(amount for key,amount in a.items() if key not in b)
    resolved=-sum(amount for key,amount in b.items() if key not in a)
    changed=sum(a[key]-b[key] for key in a.keys()&b.keys())
    old=sum(b.values());now=sum(a.values())
    assert old+new+resolved+changed==now
    return dict(previous_minor=old,new_minor=new,resolved_minor=resolved,
                changed_minor=changed,current_minor=now,delta_minor=now-old,
                new_count=len(a.keys()-b.keys()),resolved_count=len(b.keys()-a.keys()),
                changed_count=sum(a[k]!=b[k] for k in a.keys()&b.keys()))

def matrix(rows):
    groups=defaultdict(list)
    for p in rows:
        if int(p['current_minor']):
            groups[(effective_reason(p),effective_action(p))].append(p)
    out=[dict(reason=k[0],action=k[1],**signed_summary(v)) for k,v in groups.items()]
    return sorted(out,key=lambda d:(-d['exposure_minor'],d['reason'],d['action']))

def top_groups(filtered,dimension,limit=10):
    if dimension not in ('Vendor','Plant'):raise ValueError('Unsupported dimension')
    groups=defaultdict(list)
    for p,plant in filtered:
        if int(p['current_minor']):
            name=(p.get('vendor_text') if dimension=='Vendor' else plant) or '(Not assigned)'
            groups[str(name)].append(p)
    out=[dict(name=k,**signed_summary(v)) for k,v in groups.items()]
    return sorted(out,key=lambda x:(-x['exposure_minor'],x['name']))[:limit]

def priority(rows,period,limit=100):
    ranked=[]
    for p in rows:
        amount=int(p['current_minor'])
        if not amount:continue
        age=age_days(p,period);score=0;reasons=[]
        if abs(amount)>=1_000_000:score+=30;reasons.append('Amount ≥ 10,000')
        elif abs(amount)>=250_000:score+=20;reasons.append('Amount ≥ 2,500')
        elif abs(amount)>=50_000:score+=10;reasons.append('Amount ≥ 500')
        if age is not None:
            if age>180:score+=30;reasons.append('Older than 180 days')
            elif age>90:score+=20;reasons.append('Older than 90 days')
            elif age>30:score+=10;reasons.append('Older than 30 days')
        if p['match_status']=='To Be Reviewed':score+=25;reasons.append('Needs review')
        if is_potential_reversal(p):score+=15;reasons.append('Potential GR reversal')
        ranked.append(dict(position=p,score=score,age=age,why=', '.join(reasons) or 'No priority trigger'))
    return sorted(ranked,key=lambda r:(-r['score'],-abs(int(r['position']['current_minor'])),r['position']['po'],r['position']['item_line']))[:limit]

def multi_entity(archive,period):
    entities=[r[0] for r in archive.db.execute('SELECT DISTINCT entity FROM reconciliations ORDER BY entity')]
    report=[]
    for entity in entities:
        p=next((r for r in archive.all_periods(entity) if r['period']==period),None)
        if p is None:continue
        summary=archive.summary(p['id'])
        if not summary['is_verified']:continue
        report.append(dict(entity=entity,period=period,currency=summary['currency'],
                           net_minor=summary['total_minor'],
                           exposure_minor=sum(abs(int(r['current_minor'])) for r in archive.positions(p['id'])),
                           reversal_minor=abs(summary['reversal_minor']),
                           open_count=summary['open_items'],version=summary['version']))
    return report
