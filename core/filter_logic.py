from classification import effective_reason, effective_action, is_potential_reversal
"""Pure filtering and signed accounting totals for the desktop detail grid."""


def select_positions(rows, *, po='', vendor='', plant='', gl='', assignment='',
                     category='All', reason_text='', action='All'):
    terms = [str(v or '').strip().casefold() for v in (po, vendor, plant, gl, assignment, reason_text)]
    results=[]
    for p in rows:
        account=str(p.get('account_assignment') or '')
        plant_code=account[:4] if len(account)>=4 and account[:4].isdigit() and account.startswith('3') else ''
        fields=(str(p.get('po') or ''),str(p.get('vendor_text') or ''),plant_code,
                str(p.get('gl_account') or ''),account,str(p.get('comment') or ''))
        if any(needle not in field.casefold() for needle,field in zip(terms,fields)):
            continue
        if category not in ('All','') and effective_reason(p) != category:
            continue
        if action not in ('All','') and effective_action(p) != action:
            continue
        results.append((p,plant_code))
    return results


def filtered_totals(results):
    """Keep accounting signs; do not double-count zero-value Resolved entries."""
    positions=[p for p,_ in results]
    current=sum(int(p['current_minor']) for p in positions)
    previous=sum(int(p['previous_minor']) for p in positions if p.get('previous_minor') is not None)
    reversal=sum(int(p['current_minor']) for p in positions
                 if is_potential_reversal(p))
    return {'count':len(positions),'open_count':sum(int(p['current_minor'])!=0 for p in positions),
            'current_minor':current,'previous_minor':previous,'reversal_minor':reversal}
