"""Export an unchanged SAP source sheet plus a corporate REC template layout.

The full source workbook is opened as the output base; no source inputs are written.
Existing REC styling, column positions, header/footer layout are copied from the
previous complete reconciliation workbook, not reconstructed from an image.
"""
from copy import copy
from datetime import date, datetime
from calendar import monthrange
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties

HEADER_ROW=27
DATA_START=28

def locate_rec(wb):
    for ws in wb:
        for row in ws.iter_rows(min_row=1,max_row=min(60,ws.max_row),min_col=1,max_col=min(15,ws.max_column)):
            labels={str(c.value or '').strip().lower():c.column for c in row}
            if any('purchase order' in name for name in labels) and any('items line' in name for name in labels):
                return ws,row[0].row
    raise ValueError('Template does not contain a recognizable reconciliation table (Purchase order / Items line).')

def template_bottom(ws):
    for row in ws.iter_rows(min_row=DATA_START,min_col=4,max_col=4):
        if str(row[0].value or '').strip().lower().startswith('total supporting/subsidiary detail'):
            return row[0].row
    raise ValueError('Template REC lacks the Total Supporting/Subsidiary Detail footer')

def transfer_row(src,dst,src_num,dst_num,max_col=9,*,values=True):
    for col in range(1,max_col+1):
        a=src.cell(src_num,col);b=dst.cell(dst_num,col)
        if a.has_style:b._style=copy(a._style)
        if a.number_format:b.number_format=a.number_format
        if a.alignment:b.alignment=copy(a.alignment)
        if a.protection:b.protection=copy(a.protection)
        if values and a.value is not None:b.value=a.value
        if a.hyperlink:b._hyperlink=copy(a.hyperlink)
    if src_num in src.row_dimensions:
        d=src.row_dimensions[src_num];t=dst.row_dimensions[dst_num]
        if d.height is not None:t.height=d.height
        t.hidden=d.hidden

def as_excel_date(value):
    if not value:return None
    if isinstance(value,(datetime,date)):return value
    raw=str(value).strip()
    for fmt in ('%Y-%m-%d','%d.%m.%Y','%d/%m/%Y'):
        try:return datetime.strptime(raw,fmt)
        except ValueError:pass
    return raw

def export_reconciliation(source_sap,template_path,rec,positions,destination):
    if rec.get('sap_total_minor') is None or rec.get('sap_total_minor')!=sum(p['current_minor'] for p in positions):
        raise ValueError('Official export blocked: imported positions do not match the verified SAP closing balance.')
    if not Path(source_sap).is_file():raise FileNotFoundError(f'Current SAP source not found: {source_sap}')
    if not Path(template_path).is_file():raise FileNotFoundError(f'Previous full reconciliation/template not found: {template_path}')
    if Path(destination).resolve() in (Path(source_sap).resolve(),Path(template_path).resolve()):
        raise ValueError('Never overwrite either input workbook. Choose a new output file.')
    wb=load_workbook(source_sap)
    for s in wb:
        if s.title.lower() in ('rec','reconciliation'):
            raise ValueError('The current SAP workbook already contains a REC/Reconciliation worksheet; choose a SAP-only source file.')
    old=load_workbook(template_path)
    src,header=locate_rec(old)
    foot=template_bottom(src)
    last=min(src.max_row,max(foot+28,src.max_row))
    dst=wb.create_sheet('REC',0)
    dst.sheet_format=copy(src.sheet_format)
    dst.sheet_properties=copy(src.sheet_properties)
    dst.page_setup=copy(src.page_setup)
    dst.page_margins=copy(src.page_margins)
    dst.print_options=copy(src.print_options)
    dst.oddHeader=copy(src.oddHeader)
    dst.oddFooter=copy(src.oddFooter)
    dst.evenHeader=copy(src.evenHeader)
    dst.evenFooter=copy(src.evenFooter)
    dst.firstHeader=copy(src.firstHeader)
    dst.firstFooter=copy(src.firstFooter)
    for letter,dim in src.column_dimensions.items():
        target=dst.column_dimensions[letter]
        target.width=dim.width;target.hidden=dim.hidden;target.bestFit=dim.bestFit
    for i in range(1,header+1):transfer_row(src,dst,i,i)
    for merge in src.merged_cells.ranges:
        if merge.max_row<=header:dst.merge_cells(str(merge))
    rec_positions=[p for p in positions if p['current_minor']!=0 and p.get('match_status')!='Resolved']
    if any(p['current_minor']!=0 and p.get('match_status')=='Resolved' for p in positions):
        raise ValueError('Cannot export a Resolved position with a nonzero current amount. Review the matching data.')
    new_rows=len(rec_positions)
    original_last_data=foot-4
    for i,p in enumerate(rec_positions):
        r=DATA_START+i
        style_row=min(r,original_last_data)
        transfer_row(src,dst,style_row,r,values=False)
        for col,field in [(2,'vendor_text'),(3,'po'),(4,'gl_account'),(5,'account_assignment'),(6,'posting_date'),(8,'comment'),(9,'item_line')]:
            val=p.get(field,'') or None
            if col==6:val=as_excel_date(val)
            dst.cell(r,col).value=val
        dst.cell(r,7).value=p['current_minor']/100
        dst.cell(r,7).number_format='#,##0.00;(#,##0.00);-'
        dst.cell(r,3).number_format='@';dst.cell(r,9).number_format='@'
    new_foot=DATA_START+new_rows+3
    offset=new_foot-foot
    for i in range(foot,last+1):transfer_row(src,dst,i,i+offset)
    period=str(rec['period'])
    yy,mm=map(int,period.split('-'))
    current_close=date(yy,mm,monthrange(yy,mm)[1])
    if mm==1:previous_close=date(yy-1,12,31)
    else:previous_close=date(yy,mm-1,monthrange(yy,mm-1)[1])
    dst['A9']=current_close
    dst['C14']=rec.get('prepared_by','')
    dst['C15']=rec.get('supervisor','')
    dst['F14']=as_excel_date(rec.get('preparation_date',''))
    dst['F15']=None
    dst['E18']=current_close
    dst['E18'].number_format='dd.mm.yyyy'
    previous_gl=src['F18'].value
    dst['D22']=previous_close if isinstance(previous_gl,(int,float)) else None
    dst['D22'].number_format='dd.mm.yyyy'
    dst['E22']=previous_gl if isinstance(previous_gl,(int,float)) else None
    dst['F18']=rec['sap_total_minor']/100
    total=new_foot
    dst.cell(total,6).value=(f'=SUM(G{DATA_START}:G{DATA_START+new_rows-1})' if new_rows else '=0')
    dst.cell(total+2,6).value=f'=F18-F{total}'
    recon_items=total+5
    dst.cell(total+10,6).value=f'=SUM(E{recon_items}:E{recon_items+4})'
    dst.cell(total+12,6).value=f'=F{total+2}+F{total+10}-F18+F{total}'
    dst.freeze_panes=src.freeze_panes
    if src.auto_filter and src.auto_filter.ref:
        original_filter=src.auto_filter.ref
        left=original_filter.split(':')[0]
        right=''.join(ch for ch in original_filter.split(':')[-1] if ch.isalpha())
        dst.auto_filter.ref=f'{left}:{right}{max(header,DATA_START+new_rows-1)}'
    else:
        dst.auto_filter.ref=f'B{header}:I{max(header,DATA_START+new_rows-1)}'
    dst.sheet_view.showGridLines=src.sheet_view.showGridLines
    if src.print_area:dst.print_area=f'A7:I{new_foot+28}'
    if len(wb.worksheets)>=2 and wb.worksheets[1].title=='Sheet1':
        wb.worksheets[1].title=f'Open item as of {current_close:%d.%m.%Y}'
    check=wb.create_sheet('Reconciliation Status')
    headers=['Purchase Order','Item Line','Previous Amount','Current Amount','Difference','Match Status','Reason Category','Previous Comment','Review Required']
    check.append(headers)
    for p in positions:
        prev=p.get('previous_minor');now=p['current_minor']
        check.append([p['po'],p['item_line'],prev/100 if prev is not None else None,now/100,(now-prev)/100 if prev is not None else None,p['match_status'],p.get('reason_category',''),p.get('comment',''), 'Yes' if p['current_minor']<0 and (p['match_status'] in ('To Be Reviewed','Amount Changed - Still Unresolved') or not str(p.get('comment') or '').strip()) else 'No'])
    from openpyxl.styles import Font,PatternFill,Alignment
    for cell in check[1]:
        cell.fill=PatternFill('solid',fgColor='213452');cell.font=Font(color='FFFFFF',bold=True)
        cell.alignment=Alignment(wrap_text=True)
    for col,width in {'A':19,'B':13,'C':20,'D':20,'E':19,'F':34,'G':26,'H':55,'I':18}.items():
        check.column_dimensions[col].width=width
    for row in range(2,check.max_row+1):
        check.cell(row,1).number_format='@';check.cell(row,2).number_format='@'
        for col in (3,4,5):check.cell(row,col).number_format='#,##0.00;(#,##0.00);-'
    check.freeze_panes='C2';check.auto_filter.ref=f'A1:I{check.max_row}'
    validation=wb.create_sheet('SAP Balance Validation')
    validation.append(['Control','GBP / Local currency'])
    validation.append(['SAP grand total',rec['sap_total_minor']/100])
    validation.append(['Sum of current REC open items',f'=SUM(REC!G{DATA_START}:G{DATA_START+new_rows-1})' if new_rows else '=0'])
    validation.append(['Difference (should be zero)','=B3-B2'])
    validation.append(['Validation status','=IF(ROUND(B4,2)=0,"BALANCED","CHECK REQUIRED")'])
    validation.column_dimensions['A'].width=40;validation.column_dimensions['B'].width=28
    for row in (2,3,4):validation.cell(row,2).number_format='#,##0.00;(#,##0.00);-'
    wb.calculation=CalcProperties(calcMode='auto',fullCalcOnLoad=True,forceFullCalc=True)
    wb.active=0
    Path(destination).parent.mkdir(parents=True,exist_ok=True)
    wb.save(destination)
    return {'path':str(destination),'records':new_rows,'status_records':len(positions),'excluded_zero_or_resolved':len(positions)-new_rows,'rec_sheet':'REC','status_sheet':'Reconciliation Status','footer_row':new_foot}
