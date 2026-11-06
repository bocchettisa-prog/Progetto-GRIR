"""GR/IR Control Center – public demo web edition.

Every session has its own data. No database, shared caches or permanent file storage.
Uploaded workbooks are processed in short-lived temporary directories.
"""
from __future__ import annotations

import io
import sys
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).parent / "core"))
from importer import import_sap, historical_snapshot_from_excel, filename_metadata, reconcile
from exporter import export_reconciliation
from classification import effective_action, effective_reason, action_source
from filter_logic import select_positions, filtered_totals
from analytics import breakdown
from insights_engine import aging, matrix, movement, priority, signed_summary

st.set_page_config(page_title="GR/IR Control Center · Web", page_icon="📊", layout="wide", initial_sidebar_state="expanded")
st.markdown('''<style>
.block-container{padding-top:1.5rem;max-width:1450px}
[data-testid="stMetric"]{background:rgba(125,136,170,.075);padding:15px;border:1px solid rgba(128,128,128,.13);border-radius:12px}
.footer{font-size:.78rem;color:#8791a6;text-align:center;padding:30px 0 10px}
</style>''', unsafe_allow_html=True)

def amt(v,cur='GBP'):
    symbol={'GBP':'£','EUR':'€','USD':'$'}.get(cur,cur+' ')
    return ('−' if v<0 else '')+symbol+f'{abs(v)/100:,.2f}'

def dataframe(items):
    return pd.DataFrame(items)

def read_upload(upload):
    if upload is None:return None
    return upload.getvalue()

st.sidebar.title('GR/IR Control Center')
st.sidebar.caption('Web Edition · Created and tested by Salvatore Bocchetti')
st.sidebar.info('Public prototype · Uploaded Excel files are processed on the hosting server. Do not upload confidential company data.')

with st.sidebar.expander('1 · Upload Excel files',expanded=True):
    current_file=st.file_uploader('Current SAP report (without REC)',type=['xlsx'],key='current')
    previous_file=st.file_uploader('Previous full report (SAP + REC)',type=['xlsx'],key='previous')
    st.caption('Both filenames must include entity and period, for example GB12 GRIR 10.26.xlsx.')
    trigger=st.button('Import & reconcile',type='primary',use_container_width=True)

if trigger:
    if not current_file or not previous_file:
        st.error('Please upload both Excel files.')
    else:
        try:
            c_entity,c_period=filename_metadata(current_file.name)
            p_entity,p_period=filename_metadata(previous_file.name)
            if c_entity!=p_entity or p_period>=c_period:
                raise ValueError('Previous report must have the same entity and an earlier period.')
            with st.spinner('Reading SAP subtotals, validating balances and reconciling...'):
                with tempfile.TemporaryDirectory(prefix='grir_web_') as td:
                    cur=Path(td)/current_file.name
                    prev=Path(td)/previous_file.name
                    cur.write_bytes(read_upload(current_file))
                    prev.write_bytes(read_upload(previous_file))
                    sap=import_sap(cur)
                    snap=historical_snapshot_from_excel(prev,c_entity,c_period,sap['currency'])
                    old={(p['po'],p['item_line']):p for p in snap['rows']}
                    rows=reconcile(sap['records'],old)
                    if sum(int(p['current_minor']) for p in rows)!=sap['grand_total_minor']:
                        raise ValueError('Internal balance mismatch. Import cancelled.')
                st.session_state.pop('download_bytes', None)
                st.session_state.pop('download_name', None)
                st.session_state['work']={
                    'rows':rows,'original_rows':[dict(p) for p in rows],
                    'old_rows':snap['rows'],'entity':c_entity,'period':c_period,'prev_period':p_period,
                    'currency':sap['currency'],'total':sap['grand_total_minor'],
                    'previous_total':snap['grand_total_minor'],
                    'sap_source_name':current_file.name,'previous_name':previous_file.name,
                    'source_bytes':read_upload(current_file),'previous_bytes':read_upload(previous_file),
                    'warnings':sap['warnings'],'partial_prev':snap['partial'],
                    'prepared_by':'','supervisor':'','preparation_date':'',
                }
                # Session-only history: no browser-independent persistence.
                st.session_state.setdefault('session_history',{})[(c_entity,p_period)]=snap['rows']
                st.session_state['session_history'][(c_entity,c_period)]=rows
                st.session_state.setdefault('session_currencies',{})[(c_entity,p_period)]=sap['currency']
                st.session_state['session_currencies'][(c_entity,c_period)]=sap['currency']
            st.success('Imported and balanced successfully. Original files were not changed.')
        except Exception as exc:
            st.error(f'Import stopped: {exc}')

work=st.session_state.get('work')
if not work:
    st.title('GR/IR Control Center')
    st.subheader('Reconciliation software · Web Edition')
    st.write('Upload a **current SAP open-item export** and the **previous complete reconciliation** using the panel on the left. The application checks the SAP closing balance before allowing the official Excel export.')
    st.info('Your files are not included in the application or in its public source code. This initial web version does not retain reconciliation history between visits.')
    st.markdown('<div class="footer">GR/IR Reconciliation Software | Created and tested by Salvatore Bocchetti</div>',unsafe_allow_html=True)
    st.stop()

rows=work['rows']; currency=work['currency']; period=work['period']
st.title('GR/IR Control Center')
st.caption(f"{work['entity']} · {period} · {currency} · SAP balance validated")
tabs=st.tabs(['Executive Dashboard','Reconciliation Items','Aging','Movement','Reason & Action Matrix','Top Vendors & Plants','Priority Score','Multi-Entity','Export Excel'])

with st.sidebar.expander('2 · Filters',expanded=True):
    po=st.text_input('Purchase Order',key='f_po')
    vendor=st.text_input('Vendor',key='f_vendor')
    plant=st.text_input('Plant',key='f_plant')
    gl=st.text_input('GL Account',key='f_gl')
    assignment=st.text_input('Account Assignment',key='f_assignment')
    comment_search=st.text_input('Full Reason / Comment text',key='f_reason_text')
    reasons=['All']+sorted({effective_reason(p) for p in rows})
    actions=['All']+sorted({effective_action(p) for p in rows})
    reason=st.selectbox('Reason / Discrepancy',reasons,key='f_reason')
    action=st.selectbox('Corrective Action',actions,key='f_action')
    include_resolved=st.checkbox('Include Resolved (zero balance)',value=False)
filtered=select_positions(rows,po=po,vendor=vendor,plant=plant,gl=gl,assignment=assignment,category=reason,reason_text=comment_search,action=action)
if not include_resolved:filtered=[(p,pt) for p,pt in filtered if int(p['current_minor'])!=0]
selected=[p for p,_ in filtered]
ft=filtered_totals(filtered)

def listing(items):
    out=[]
    for p in items:
        acc=str(p.get('account_assignment') or '')
        out.append({'PO':p['po'],'Item Line':p['item_line'],'Vendor':p.get('vendor_text',''),
             'GL':p.get('gl_account',''),'Account Assignment':acc,'Plant':acc[:4] if acc[:4].isdigit() and acc.startswith('3') else '',
             'Previous':(p['previous_minor']/100 if p.get('previous_minor') is not None else None),
             'Current':p['current_minor']/100,'Status':p['match_status'],
             'Reason':effective_reason(p),'Action':effective_action(p),'Full Comment':p.get('comment','')})
    return out

with tabs[0]:
    previous=work['previous_total']; current=work['total']; delta=current-previous
    a,b,c,d=st.columns(4)
    a.metric('Net GR/IR (absolute)',amt(abs(current),currency))
    b.metric('Previous net GR/IR (absolute)',amt(abs(previous),currency))
    c.metric('Balance movement (signed)',amt(delta,currency),delta=f'{(abs(current)-abs(previous))/abs(previous)*100:+.2f}% exposure' if previous else None,delta_color='inverse')
    d.metric('SAP check',amt(0,currency),'VERIFIED',delta_color='off')
    x,y,z=st.columns(3)
    x.metric('Filtered positions',len(selected))
    y.metric('Filtered net amount',amt(ft['current_minor'],currency))
    z.metric('Filtered potential GR reversal',amt(abs(ft['reversal_minor']),currency))
    st.info('All financial totals use the original signed SAP values. Filters affect analysis totals, not the verified SAP closing balance.')
    for message in work['warnings'][:3]:st.caption(message)
    if work['partial_prev']:st.warning('Previous REC is partial. Missing findings are flagged for review; previous period financial balance is calculated from SAP, not the partial REC.')
    st.subheader('Reason breakdown — filtered positions')
    grp=breakdown(filtered,'Reason / Discrepancy')
    if grp:
        chart=dataframe([{'Reason':r['name'],'Exposure':r['exposure_minor']/100} for r in grp]).set_index('Reason')
        st.bar_chart(chart,horizontal=True)
    st.dataframe(dataframe(listing(selected)),use_container_width=True,hide_index=True,height=380)

with tabs[1]:
    st.subheader('Reconciliation Items & Review')
    st.caption('Changes are held only in this browser session until you download the Excel. No shared database is used.')
    st.metric('Filtered net total',amt(ft['current_minor'],currency))
    st.dataframe(dataframe(listing(selected)),use_container_width=True,hide_index=True,height=450)
    labels=[f"{p['po']} / {p['item_line']} · {p['match_status']}" for p in selected]
    if labels:
        chosen=st.selectbox('Edit selected position',range(len(labels)),format_func=lambda i:labels[i])
        item=selected[chosen]
        with st.form('edit_position'):
            c1,c2=st.columns(2)
            vendor_new=c1.text_input('Vendor/Text',value=item.get('vendor_text') or '')
            gl_new=c2.text_input('GL Account',value=item.get('gl_account') or '')
            ass_new=c1.text_input('Account Assignment',value=item.get('account_assignment') or '')
            posting_new=c2.text_input('Posting Date',value=str(item.get('posting_date') or ''))
            comment_new=st.text_area('Full Comment / Reason',value=item.get('comment') or '')
            action_new=st.selectbox('Corrective action override', ['Automatic','Not Specified','GR Reversal Required','GR Reversed (Completed)','Needs Review'],index=(['Automatic','Not Specified','GR Reversal Required','GR Reversed (Completed)','Needs Review'].index(item.get('action_override') or 'Automatic') if (item.get('action_override') or 'Automatic') in ['Automatic','Not Specified','GR Reversal Required','GR Reversed (Completed)','Needs Review'] else 0))
            saved=st.form_submit_button('Save changes')
        if saved:
            item.update(vendor_text=vendor_new,gl_account=gl_new,account_assignment=ass_new,posting_date=posting_new,comment=comment_new,action_override='' if action_new=='Automatic' else action_new)
            if item['current_minor']!=0:
                item['match_status']='To Be Reviewed' if not comment_new.strip() else ('Amount Changed - Still Unresolved' if item.get('previous_minor') is not None and item['previous_minor']!=item['current_minor'] else 'Still Open')
            st.success('Saved in this session. Download the workbook to retain your edits.')
            st.rerun()

with tabs[2]:
    st.subheader('Aging Analysis')
    st.caption('Based on recorded Posting Date. Missing or invalid dates remain in Unknown; dates are never inferred.')
    ag=aging(selected,period)
    st.dataframe(dataframe([{'Bucket':a['bucket'],'Items':a['count'],'Net Amount':a['net_minor']/100,'Exposure':a['exposure_minor']/100} for a in ag]),use_container_width=True,hide_index=True)
    st.bar_chart(dataframe({'Bucket':[a['bucket'] for a in ag],'Exposure':[a['exposure_minor']/100 for a in ag]}).set_index('Bucket'))

with tabs[3]:
    st.subheader('Movement Analysis')
    st.caption('This bridge reconciles the full two SAP snapshots and is not narrowed by dashboard filters.')
    m=movement(rows,work['old_rows'])
    bridge=[('Previous closing',m['previous_minor']),('New positions',m['new_minor']),('Resolved positions',m['resolved_minor']),('Amount changes',m['changed_minor']),('Current closing',m['current_minor'])]
    st.dataframe(dataframe([{'Component':n,'Signed GBP / Local':v/100} for n,v in bridge]),hide_index=True,use_container_width=True)
    st.metric('Bridge validation',amt(m['previous_minor']+m['new_minor']+m['resolved_minor']+m['changed_minor']-m['current_minor'],currency))
    st.caption(f"New: {m['new_count']} · Resolved: {m['resolved_count']} · Changed: {m['changed_count']}")

with tabs[4]:
    st.subheader('Reason & Corrective Action Matrix')
    mx=matrix(selected)
    st.dataframe(dataframe([{'Reason':v['reason'],'Action':v['action'],'Items':v['count'],'Net':v['net_minor']/100,'Exposure':v['exposure_minor']/100,'Potential Reversal':v['reversal_minor']/100} for v in mx]),use_container_width=True,hide_index=True)

with tabs[5]:
    st.subheader('Top Vendors & Plants')
    for dim in ('Vendor','Plant'):
        data=breakdown(filtered,dim)[:10]
        st.markdown(f'**Top 10 — {dim}**')
        if data:
            st.bar_chart(dataframe([{'Name':d['name'],'Exposure':d['exposure_minor']/100} for d in data]).set_index('Name'),horizontal=True)
            st.dataframe(dataframe([{'Name':d['name'],'Items':d['open_count'],'Net':d['net_minor']/100,'Exposure':d['exposure_minor']/100} for d in data]),hide_index=True,use_container_width=True)

with tabs[6]:
    st.subheader('Priority Score')
    st.caption('Worklist suggestion based on amount, age and review status; it is not an accounting risk certification.')
    score=priority(selected,period)
    st.dataframe(dataframe([{'Score':v['score'],'PO':v['position']['po'],'Line':v['position']['item_line'],'Amount':v['position']['current_minor']/100,'Age (days)':v['age'],'Priority Factors':v['why']} for v in score]),use_container_width=True,hide_index=True)

with tabs[7]:
    st.subheader('Multi-Entity — Current Browser Session')
    st.caption('This free public prototype does not preserve or share user histories. Only reports loaded in this session are available.')
    history=st.session_state.get('session_history',{})
    entries=[]
    for (entity,hperiod),hrows in sorted(history.items()):
        summary=signed_summary(hrows)
        entries.append({'Entity':entity,'Period':hperiod,'Currency':st.session_state.get('session_currencies',{}).get((entity,hperiod),'?'),'Signed Net':summary['net_minor']/100,'Positions':summary['count'],'Potential Reversal':summary['reversal_minor']/100})
    st.dataframe(dataframe(entries),hide_index=True,use_container_width=True)

with tabs[8]:
    st.subheader('Export Official Excel')
    st.caption('The exported workbook contains the current SAP worksheet, the corporate REC based on your previous template, Reconciliation Status, and SAP Balance Validation.')
    with st.form('export_details'):
        prepared=st.text_input('Prepared By',value=work.get('prepared_by',''))
        supervisor=st.text_input('Supervisor',value=work.get('supervisor',''))
        prep_date=st.date_input('Preparation Date')
        go=st.form_submit_button('Generate Excel',type='primary')
    if go:
        try:
            rec=dict(period=period,entity=work['entity'],currency=currency,prepared_by=prepared,supervisor=supervisor,preparation_date=prep_date.isoformat(),sap_total_minor=work['total'])
            with tempfile.TemporaryDirectory(prefix='grir_export_') as td:
                current_path=Path(td)/work['sap_source_name']
                previous_path=Path(td)/work['previous_name']
                target=Path(td)/f"{work['entity']} GRIR {period[5:7]}.{period[2:4]} Reconciled.xlsx"
                current_path.write_bytes(work['source_bytes'])
                previous_path.write_bytes(work['previous_bytes'])
                export_reconciliation(current_path,previous_path,rec,rows,target)
                output=target.read_bytes()
            work.update(prepared_by=prepared,supervisor=supervisor,preparation_date=prep_date.isoformat())
            st.session_state['download_bytes']=output
            st.session_state['download_name']=target.name
            st.success('Excel generated. All Resolved / zero-balance positions are excluded from REC but included in Reconciliation Status.')
        except Exception as exc: st.error(f'Export failed: {exc}')
    if st.session_state.get('download_bytes'):
        st.download_button('Download Official Reconciliation',data=st.session_state['download_bytes'],file_name=st.session_state['download_name'],mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary')

st.markdown('<div class="footer">GR/IR Reconciliation Software | Created and tested by Salvatore Bocchetti · Web prototype</div>',unsafe_allow_html=True)
