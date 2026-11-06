"""Read-only aggregations of the currently filtered reconciliation positions.

All amounts use integer minor units and retain SAP accounting signs.
"""
from collections import defaultdict
from classification import effective_reason, effective_action, is_potential_reversal

DIMENSIONS = {
    'Reason / Discrepancy': 'reason_category',
    'Corrective Action': 'corrective_action',
    'Vendor': 'vendor_text',
    'Plant': 'plant',
    'GL Account': 'gl_account',
}


def breakdown(filtered, dimension):
    """Aggregate (position, derived_plant) pairs. Zero/Resolved rows are not exposure.

    Net = signed sum, Exposure = sum of absolute nonzero entries, not abs(Net).
    Potential reversal = abs(negative items classified GR Reversal Required).
    """
    if dimension not in DIMENSIONS:
        raise ValueError('Unsupported analytics dimension')
    groups = defaultdict(lambda: dict(count=0, open_count=0, net_minor=0,
                                      exposure_minor=0, reversal_minor=0, positive_count=0))
    for p, plant in filtered:
        name = (plant if dimension == 'Plant' else
                effective_reason(p) if dimension == 'Reason / Discrepancy' else
                effective_action(p) if dimension == 'Corrective Action' else
                str(p.get(DIMENSIONS[dimension]) or '').strip())
        name = name or '(Not assigned)'
        row = groups[name]
        row['count'] += 1
        amount = int(p['current_minor'])
        row['net_minor'] += amount
        if amount:
            row['open_count'] += 1
            row['exposure_minor'] += abs(amount)
            if amount > 0:
                row['positive_count'] += 1
            if is_potential_reversal(p):
                row['reversal_minor'] += abs(amount)
    denominator = sum(v['exposure_minor'] for v in groups.values())
    output = []
    for name, values in groups.items():
        output.append(dict(name=name, **values,
                           share_pct=(100 * values['exposure_minor'] / denominator if denominator else 0)))
    return sorted(output, key=lambda row: (-row['exposure_minor'], row['name'].casefold()))


def overview(filtered):
    amounts = [int(p['current_minor']) for p, _ in filtered]
    return {
        'count': len(filtered),
        'open_count': sum(a != 0 for a in amounts),
        'net_minor': sum(amounts),
        'exposure_minor': sum(abs(a) for a in amounts),
        'reversal_minor': sum(abs(int(p['current_minor'])) for p, _ in filtered
                              if is_potential_reversal(p)),
        'unclassified_count': sum(1 for p, _ in filtered if int(p['current_minor']) and
                                  effective_reason(p) in ('', 'Unclassified')),
    }
