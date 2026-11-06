"""V11: distinguish a discrepancy's root reason from proposed/completed action.

The original comment is always authoritative and never rewritten by this module.
Auto-derived actions are suggestions, not approved accounting instructions.
"""
import re
UNCLASSIFIED = 'Unclassified'
ACTIONS = ('Not Specified', 'GR Reversal Required', 'GR Reversed (Completed)', 'Needs Review')
REASONS = ('Invoice Never Received', 'Invoice Received for Smaller Amount',
    'Invoice Received for Smaller Quantity', 'Partial Credit Note Received',
    'Full Credit Note Received', 'Invoice Parked', 'Invoice Awaiting Receipt',
    'GR Reversed', 'Other / Review', 'Unclassified')
LEGACY_REASONS = {'GR Reversal Required', 'Awaiting Invoice', 'Quantity Difference',
    'Amount Difference', 'Needs Classification', 'Other / Review'}

def _normalize(comment):
    return ' '.join(str(comment or '').strip().casefold().split())

def classify_reason(comment):
    t = _normalize(comment)
    if not t:return UNCLASSIFIED
    if re.search(r'partial\s+(?:credit\s+note|cn)(?:\s+(?:was\s+)?received)?\b',t):
        return 'Partial Credit Note Received'
    if re.search(r'full\s+(?:credit\s+note|cn)(?:\s+(?:was\s+)?received)?\b',t):
        return 'Full Credit Note Received'
    if re.search(r'(?:invoice\s+(?:was\s+)?received\s+(?:for\s+)?(?:a\s+)?smaller\s+amount|invoice\s+smaller\s+amount|smaller\s+amount)',t):
        return 'Invoice Received for Smaller Amount'
    if re.search(r'(?:invoice\s+(?:was\s+)?received\s+(?:for\s+)?(?:a\s+)?smaller\s+quantity|smaller\s+quantity)',t):
        return 'Invoice Received for Smaller Quantity'
    if re.search(r'invoice\s+(?:was\s+)?never\s+received',t):return 'Invoice Never Received'
    if re.search(r'invoice\s+(?:is\s+)?awaiting\s+(?:goods\s+)?receipt',t):return 'Invoice Awaiting Receipt'
    if re.search(r'invoice\s+(?:is\s+)?parked',t):return 'Invoice Parked'
    if re.search(r'\bgr\s+(?:already\s+)?reversed\b',t):return 'GR Reversed'
    if re.search(r'awaiting\s+invoice|invoice\s+not\s+received',t):return 'Invoice Never Received'
    return 'Other / Review'

def classify_action(comment):
    t = _normalize(comment)
    if not t:return 'Not Specified'
    if re.search(r'\b(gr\s+(?:already\s+)?reversed|gr\s+has\s+been\s+reversed|reversal\s+(?:already\s+)?completed)\b',t):
        return 'GR Reversed (Completed)'
    if re.search(r'\b(?:not\s+to\s+reverse|do\s+not\s+reverse|no\s+need\s+to\s+reverse|gr\s+should\s+not\s+be\s+reversed|reversal\s+not\s+required)\b',t):
        return 'Needs Review'
    if re.search(r'\b(?:gr\s+to\s+be\s+reversed|gr\s+needs\s+to\s+be\s+reversed|gr\s+reversal\s+required|reverse\s+(?:the\s+)?gr|to\s+reverse)\b',t):
        return 'GR Reversal Required'
    if classify_reason(t) in ('Invoice Received for Smaller Amount','Invoice Received for Smaller Quantity'):
        return 'GR Reversal Required'
    return 'Not Specified'

def action_source(comment):
    action = classify_action(comment)
    if action != 'GR Reversal Required':
        return 'Comment' if action in ('GR Reversed (Completed)', 'Needs Review') else 'Not Specified'
    t=_normalize(comment)
    return ('Explicit' if re.search(r'\b(?:gr\s+to\s+be\s+reversed|gr\s+needs\s+to\s+be\s+reversed|gr\s+reversal\s+required|reverse\s+(?:the\s+)?gr|to\s+reverse)\b',t)
            else 'Inferred — verify')

def effective_reason(position):
    stored=str(position.get('reason_category') or '').strip()
    if stored and stored not in LEGACY_REASONS and stored != UNCLASSIFIED:return stored
    return classify_reason(position.get('comment'))

def effective_action(position):
    return str(position.get('action_override') or '').strip() or classify_action(position.get('comment'))

def is_potential_reversal(position):
    return int(position.get('current_minor') or 0)<0 and effective_action(position)=='GR Reversal Required'
