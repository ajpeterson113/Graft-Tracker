#!/usr/bin/env python3
"""Clinic to-dos for the Chi West Patients app: stock alerts and what to order, per account.

The Patients app's schedule (By clinic view) has a "To do" button on each clinic. It lists that
clinic's pending IVRs and follow-ups to verify (from the app's own patients) plus the stock and
ordering items written here, which come from the Graft Tracker and the patient list.

    python3 -I tools/clinic_todos.py APP_OUT PATIENT_LIST.xlsx OUT.json [--today YYYY-MM-DD]

APP_OUT            the Graft Tracker app build folder (graft_app.py build ... OUT); reads OUT/meta/stock.json
PATIENT_LIST.xlsx  the patient list Excel; the Good Samaritan tab's "To Order" block (columns O-Q)
                   becomes Good Sam's ordering item
OUT.json           write it to the Patients app (https://claude.ai/artifact/AwS5M2wc9Wm7Pe6oTuJuwQ) with
                   ArtifactData set: collection "meta", doc "clinic", file_path=OUT.json
                   (read the doc first and pass its version as if_version once it exists)

Refresh it whenever the tracker's stock tab or the Good Sam To Order block changes.
"""
import argparse
import datetime as dt
import json
import os
import re

import openpyxl

# Graft Tracker stock locations -> Patients app account names
LOC_ACCT = {
    'alexian': 'Alexian Brothers Wound Center',
    'goodsamwcc': 'Advocate Good Samaritan Hospital',
    'goodsamor': 'Advocate Good Samaritan Hospital',
    'stalexius': 'Ascension Saint Alexius',
    'glenoaks': 'UChicago Medicine AdventHealth GlenOaks',
    'elmhurst': 'Elmhurst Hospital',
    'alexianamb': 'Alexian Brothers - Ambulatory',
}
GOOD_SAM = 'Advocate Good Samaritan Hospital'


def mdy(s):
    try:
        d = dt.date.fromisoformat(str(s)[:10])
        return f'{d.month}/{d.day}/{str(d.year)[2:]}'
    except ValueError:
        return str(s or '')


def alerts_with_incoming(L):
    """Stock alerts with what's already on the way for that size (AJ 11/6), same wording as Chi West."""
    out = []
    for a in L.get('alerts') or []:
        a = dict(a)
        for s in L.get('sizes') or []:
            n = s.get('incoming') or 0
            if not n:
                m = re.match(r'\s*On order:\s*(\d+)', str(s.get('note') or ''))
                n = int(m.group(1)) if m else 0
            if n and str(a.get('title', '')).startswith(f"{s.get('product')} {s.get('size')}:"):
                m = re.search(r'arriving ([A-Z][a-z]{2} \d{1,2}/\d{1,2})', str(s.get('note') or ''))
                a['text'] = (a.get('text') or '') + (' · ' if a.get('text') else '') + f"{int(n)} on order" + (f", arriving {m.group(1)}" if m else '')
        out.append(a)
    return out


def stock_items(app_out):
    path = os.path.join(app_out, 'meta', 'stock.json')
    out = {}
    if not os.path.exists(path):
        return out
    st = json.load(open(path))
    for L in st.get('locations') or []:
        acct = LOC_ACCT.get(L.get('key'))
        if not acct or not L.get('alerts'):
            continue
        a = out.setdefault(acct, {'items': [], 'src': []})
        for al in alerts_with_incoming(L):
            a['items'].append({'kind': 'stock', 'level': al.get('level') or 'warn',
                               'title': f"{L.get('name')}: {al.get('title', '')}" if acct == GOOD_SAM else al.get('title', ''),
                               'text': al.get('text') or ''})
        a['src'].append(f"{L.get('name')} shelf from the Graft Tracker Stock tab"
                        + (f" (counted {mdy(L['counted'])}, used grafts taken off since)" if L.get('counted') else ''))
    return out


def good_sam_order(xlsx):
    """The Good Samaritan tab's To Order block (column O): title, Product/Size/Qty header, rows, then notes."""
    if not os.path.exists(xlsx):
        return None
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    if 'Good Samaritan' not in wb.sheetnames:
        return None
    ws = wb['Good Samaritan']
    top = next((r for r in range(1, 30) if str(ws.cell(r, 15).value or '').startswith('To Order')), None)
    if not top:
        return None
    title = ws.cell(top, 15).value
    rows, notes = [], []
    r = top + 2  # skip the Product / Size / Qty header
    while r <= ws.max_row:
        p, s, q = ws.cell(r, 15).value, ws.cell(r, 16).value, ws.cell(r, 17).value
        if p is None and s is None and q is None:
            break
        if s is not None or q is not None:
            rows.append((str(p or '').strip(), str(s or '').strip(), str(q or '').strip()))
        else:
            notes.append(str(p).strip())
        r += 1
    title = str(title).strip()
    if rows:
        text = ' · '.join(f"{p} {s} ×{q}".strip() for p, s, q in rows)
        return {'kind': 'order', 'level': 'warn', 'title': title, 'text': text, 'notes': notes}
    return {'kind': 'order', 'level': 'info', 'title': title, 'text': 'Nothing to order', 'notes': notes}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('app_out')
    ap.add_argument('patient_list')
    ap.add_argument('out')
    ap.add_argument('--today', default=dt.date.today().isoformat())
    a = ap.parse_args()
    accts = stock_items(a.app_out)
    gs = good_sam_order(a.patient_list)
    if gs:
        g = accts.setdefault(GOOD_SAM, {'items': [], 'src': []})
        g['items'].insert(0, gs)
        g['src'].insert(0, "WCC order from the patient list's To Order block")
    doc = {'updated': dt.datetime.now().astimezone().isoformat(timespec='seconds'), 'today': a.today,
           'accounts': {k: {'items': v['items'], 'src': ' · '.join(v['src'])} for k, v in accts.items()}}
    json.dump(doc, open(a.out, 'w'), indent=1)
    n = sum(len(v['items']) for v in doc['accounts'].values())
    print(f"{a.out}: {n} item(s) for {len(doc['accounts'])} account(s): "
          + '; '.join(f"{k} {len(v['items'])}" for k, v in doc['accounts'].items()))


if __name__ == '__main__':
    main()
