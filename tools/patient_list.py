#!/usr/bin/env python3
"""Build the patient list Excel (Chi_West_Patient_List.xlsx) from the Chi West Patients app.

    python3 -I tools/patient_list.py PDB OUT.xlsx [--today YYYY-MM-DD]
                                     [--to-order TO_ORDER.json | --prev OLD.xlsx]

PDB      ArtifactData export of the Patients app (collection `patients`, out_dir=PDB),
         so the patients are PDB/patients/*.json.
OUT      The workbook to write (Chi_West_Patient_List_<date>.xlsx, then copied to the
         repo's fixed path Chi_West_Patient_List.xlsx).
--to-order  JSON for the Good Samaritan tab's "To Order" block (columns O-Q):
            {"title": "To Order — week of 10/19 (updated 10/14)",
             "rows": [["GRAFIX", "16mm", "3"], ...],
             "notes": ["...", ...]}
--prev      Without --to-order, the block is copied from this older workbook instead.

Rules (AJ): initials only, no full names; no sales or dollar figures; Smith+Nephew colors
(orange #E9963E titles/tabs, slate blue #3F7097 headers on the To Order block).
Tabs: Summary, All Patients, one tab per account (Alexian WCC, Alexian Ambulatory and Alexian Main
OR share "Alexian"), Follow-ups, Application Log, Day Notes. Summary counts are formulas over
All Patients, so they stay right if rows are edited by hand.
"""
import argparse
import datetime as dt
import glob
import json
import os
import sys

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ORANGE, SLATE, DARK, GREY_TXT = 'FFE9963E', 'FF3F7097', 'FF1F2A36', 'FF6B7480'
HDR_FILL = PatternFill('solid', fgColor='FFEEEEEE')
SEC_FILL = PatternFill('solid', fgColor='FFFCEBD9')
THIN = Side(style='thin', color='FFBFBFBF')
MED = Side(style='medium', color='FF7F7F7F')
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
F = lambda **k: Font(name='Arial', **{'size': 10, **k})

ACCOUNTS = [  # Summary / All Patients order
    'Advocate Good Samaritan Hospital', 'Alexian Brothers Wound Center', 'Alexian Ambulatory', 'Alexian Brothers Medical Center',
    'Northwest Suburban Foot and Ankle Clinic', 'Saint Joseph Hospital - Elgin', 'Sherman',
    'Ascension Saint Alexius', 'Central DuPage Foot and Ankle', 'Elmhurst Hospital',
    'Suburban Foot and Ankle Center', 'UChicago Medicine AdventHealth GlenOaks',
    'Romano Orthopaedics - Lombard',
]
TABS = [  # (tab name, title, accounts on the tab) in tab order
    ('Good Samaritan', 'Good Samaritan', ['Advocate Good Samaritan Hospital']),
    ('Alexian', None, ['Alexian Brothers Wound Center', 'Alexian Ambulatory', 'Alexian Brothers Medical Center']),
    ('NWSFA', 'NWSFA', ['Northwest Suburban Foot and Ankle Clinic']),
    ('St Joes Elgin', 'St Joes Elgin', ['Saint Joseph Hospital - Elgin']),
    ('Sherman', 'Sherman', ['Sherman']),
    ('Ascension Saint Alexius', 'Ascension Saint Alexius', ['Ascension Saint Alexius']),
    ('Central DuPage Foot and Ankle', 'Central DuPage Foot and Ankle', ['Central DuPage Foot and Ankle']),
    ('Elmhurst Hospital', 'Elmhurst Hospital', ['Elmhurst Hospital']),
    ('Suburban Foot and Ankle Center', 'Suburban Foot and Ankle Center', ['Suburban Foot and Ankle Center']),
    ('Romano Orthopaedics - Lombard', 'Romano Orthopaedics - Lombard', ['Romano Orthopaedics - Lombard']),
    ('UChicago GlenOaks', 'UChicago Medicine AdventHealth GlenOaks', ['UChicago Medicine AdventHealth GlenOaks']),
]
SECTIONS = [  # (app status, label, always shown)
    ('active', 'Current Patients', True),
    ('approved', 'Potential New Starts (size = suggested)', True),
    ('pending', 'Pending IVRs', True),
    ('potential', 'Potential IVRs', False),
    ('or', 'STRAVIX / OR Cases', False),
    ('hold', 'On Hold', False),
    ('done', 'Patients Done', False),
    ('denied', 'Denied IVRs', True),
]
LABEL = {k: v for k, v, _ in SECTIONS}
ACCT_COLS = ['Provider', 'Initials', 'Coverage Status', 'PA Required?', 'Product', 'Size', 'Frequency',
             'Application Count', 'Last Graft', 'Case Date', 'Wks Post-Op', 'Next Appointment', 'Notes']
ACCT_W = [18, 9, 22, 13, 20, 13, 14, 12, 11, 11, 10, 17, 48]


def d(s):
    try:
        return dt.date.fromisoformat(str(s)[:10]) if s else None
    except ValueError:
        return None


def ftime(t):
    if not t:
        return ''
    h, m = (int(x) for x in str(t)[:5].split(':'))
    return f'{(h % 12) or 12}:{m:02d} {"AM" if h < 12 else "PM"}'


def mdy(x):
    return f'{x.month}/{x.day}/{x.strftime("%y")}' if x else ''


def next_text(p):
    nd = d(p.get('next'))
    if not nd:
        return ''
    return mdy(nd) + (' ' + ftime(p['nextTime']) if p.get('nextTime') else '')


def clean(s):
    s = (s or '').strip()
    while s.endswith('|'):
        s = s[:-1].rstrip()
    return s


def num(v):
    try:
        return int(v) if v not in (None, '') else None
    except (TypeError, ValueError):
        return None


def load(pdb):
    pts = []
    for f in sorted(glob.glob(os.path.join(pdb, 'patients', '*.json'))):
        p = json.load(open(f, encoding='utf-8'))
        p.setdefault('id', os.path.basename(f)[:-5])
        p.pop('name', None)  # initials only in this workbook
        pts.append(p)
    if not pts:
        sys.exit(f'no patients under {pdb}/patients')
    return pts


def sort_key(p):
    nd = p.get('next') or ''
    return (0 if nd else 1, nd, p.get('nextTime') or '', num(p.get('order')) or 0, p['id'])


def header(ws, row, cols, height=23.25):
    for i, h in enumerate(cols, 1):
        c = ws.cell(row, i, h)
        c.font = F(bold=True, color=DARK)
        c.fill = HDR_FILL
        c.alignment = Alignment(vertical='center', wrap_text=True)
    ws.row_dimensions[row].height = height


def acct_row(ws, r, p):
    vals = [p.get('provider') or '', p.get('initials') or '', p.get('coverage') or '', p.get('pa') or '',
            p.get('product') or '', p.get('size') or '', p.get('freq') or '', num(p.get('apps')),
            d(p.get('last')), d(p.get('caseDate')), None, next_text(p), clean(p.get('notes'))]
    for i, v in enumerate(vals, 1):
        c = ws.cell(r, i, v)
        c.font = F()
        c.border = BOX
    ws.cell(r, 9).number_format = 'm/d/yy'
    ws.cell(r, 10).number_format = 'm/d/yy'
    ws.cell(r, 11, f'=IF(J{r}="","",IF(J{r}>TODAY(),"Upcoming",INT((TODAY()-J{r})/7)))')
    ws.cell(r, 13).alignment = Alignment(wrap_text=True, vertical='top')


def account_tab(wb, name, title, accounts, pts):
    ws = wb.create_sheet(name)
    ws.sheet_properties.tabColor = ORANGE[2:]
    for i, w in enumerate(ACCT_W, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    r = 1
    for n, acct in enumerate(accounts):
        rows = [p for p in pts if p.get('account') == acct]
        if n and not rows:
            continue
        t = ws.cell(r, 1, title if (title and n == 0) else acct)
        t.font = F(size=14 if n == 0 else 12, bold=True)
        r += 2
        for key, label, always in SECTIONS:
            sec = sorted((p for p in rows if p.get('status') == key), key=sort_key)
            if not sec and not always:
                continue
            c = ws.cell(r, 1, label)
            c.font = F(size=11, bold=True)
            for i in range(1, len(ACCT_COLS) + 1):
                ws.cell(r, i).fill = SEC_FILL
            header(ws, r + 1, ACCT_COLS)
            r += 2
            if not sec:
                ws.cell(r, 1, 'None').font = F(italic=True, color=GREY_TXT)
                r += 1
            for p in sec:
                acct_row(ws, r, p)
                r += 1
            r += 1
    return ws


def to_order_block(ws, spec):
    """Good Samaritan 'To Order' block in O3:Q.."""
    ws.column_dimensions['N'].width = 3
    ws.column_dimensions['O'].width = 20
    ws.column_dimensions['P'].width = 16
    ws.column_dimensions['Q'].width = 11
    r = 3
    ws.merge_cells(start_row=r, start_column=15, end_row=r, end_column=17)
    c = ws.cell(r, 15, spec.get('title') or 'To Order')
    c.font = F(size=11, bold=True, color='FFFFFFFF')
    for col in (15, 16, 17):
        ws.cell(r, col).fill = PatternFill('solid', fgColor=ORANGE)
    r += 1
    for col, h in zip((15, 16, 17), ('Product', 'Size', 'Qty')):
        c = ws.cell(r, col, h)
        c.font = F(bold=True, color='FFFFFFFF')
        c.fill = PatternFill('solid', fgColor=SLATE)
        c.alignment = Alignment(horizontal='center' if col == 17 else 'left', vertical='center')
    r += 1
    for prod, size, qty in spec.get('rows') or []:
        for col, v in zip((15, 16, 17), (prod, size, qty)):
            c = ws.cell(r, col, v)
            c.font = F(bold=(col == 17))
            c.border = BOX
            c.alignment = Alignment(horizontal='center' if col == 17 else 'left', vertical='center')
        r += 1
    for note in spec.get('notes') or []:
        ws.merge_cells(start_row=r, start_column=15, end_row=r, end_column=17)
        c = ws.cell(r, 15, note)
        c.font = F(size=9, italic=True, color='FF595959')
        c.alignment = Alignment(wrap_text=True, vertical='top')
        ws.row_dimensions[r].height = max(15, 12.5 * (len(note) // 45 + 1))
        r += 1
    # outline the block
    last = r - 1
    for rr in range(3, last + 1):
        for col in (15, 16, 17):
            c = ws.cell(rr, col)
            b = c.border
            c.border = Border(left=MED if col == 15 else b.left, right=MED if col == 17 else b.right,
                              top=MED if rr == 3 else b.top, bottom=MED if rr == last else b.bottom)


def prev_to_order(prev):
    """Read the To Order block (O3:Q..) from an older workbook."""
    try:
        ws = openpyxl.load_workbook(prev)['Good Samaritan']
    except (FileNotFoundError, KeyError):
        return None
    title = ws.cell(3, 15).value
    if not title:
        return None
    rows, notes, r = [], [], 5
    while r <= ws.max_row:
        o, p_, q = (ws.cell(r, c).value for c in (15, 16, 17))
        if o is None and p_ is None and q is None:
            break
        if p_ is None and q is None and len(str(o)) > 25:
            notes.append(str(o))
        else:
            rows.append([o or '', p_ or '', q or ''])
        r += 1
    return {'title': title, 'rows': rows, 'notes': notes}


def build(pts, out, today, to_order):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    accts = ACCOUNTS + sorted({p.get('account') for p in pts if p.get('account') not in ACCOUNTS})

    # Summary
    ws = wb.create_sheet('Summary')
    ws.sheet_properties.tabColor = ORANGE[2:]
    ws['A1'] = 'Patient List — BIOT Chicago West'
    ws['A1'].font = F(size=14, bold=True, color=ORANGE)
    ws['A2'] = f'Counts update from the All Patients tab. Updated {mdy(today)}'
    ws['A2'].font = F(italic=True, color=GREY_TXT)
    cols = ['Account'] + [lab for _, lab, _ in SECTIONS[:3]] + [LABEL['potential'], LABEL['hold'], LABEL['or'],
                                                                LABEL['done'], LABEL['denied'], 'Total',
                                                                'Applications (current)']
    header(ws, 4, cols, height=30)
    ws.column_dimensions['A'].width = 38
    for i in range(2, len(cols) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 13
    r = 5
    for a in accts:
        ws.cell(r, 1, a).font = F()
        for i, lab in enumerate(cols[1:9], 2):
            ws.cell(r, i, f"=COUNTIFS('All Patients'!$A:$A,$A{r},'All Patients'!$C:$C,\"{lab}\")").font = F()
        ws.cell(r, 10, f'=SUM(B{r}:I{r})').font = F(bold=True)
        ws.cell(r, 11, f"=SUMIFS('All Patients'!$K:$K,'All Patients'!$A:$A,$A{r},"
                       f"'All Patients'!$C:$C,\"Current Patients\")").font = F()
        r += 1
    ws.cell(r, 1, 'Total').font = F(bold=True)
    for i in range(2, len(cols) + 1):
        L = get_column_letter(i)
        ws.cell(r, i, f'=SUM({L}5:{L}{r - 1})').font = F(bold=True)

    # All Patients
    ws = wb.create_sheet('All Patients')
    ws.sheet_properties.tabColor = ORANGE[2:]
    allc = ['Account', 'ID', 'Status', 'Provider', 'Initials', 'Coverage Status', 'PA Required?', 'Product',
            'Size', 'Frequency', 'Application Count', 'Last Graft', 'Case Date', 'Wks Post-Op',
            'Next Appointment', 'Notes']
    header(ws, 1, allc)
    for i, w in enumerate([30, 10, 16, 18, 9, 22, 13, 20, 13, 14, 12, 11, 11, 11, 17, 48], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    order = {k: i for i, (k, _, _) in enumerate(SECTIONS)}
    rows = sorted(pts, key=lambda p: (accts.index(p.get('account')) if p.get('account') in accts else 99,
                                      order.get(p.get('status'), 99), sort_key(p)))
    for r, p in enumerate(rows, 2):
        vals = [p.get('account') or '', p['id'], LABEL.get(p.get('status'), p.get('status') or ''),
                p.get('provider') or '', p.get('initials') or '', p.get('coverage') or '', p.get('pa') or '',
                p.get('product') or '', p.get('size') or '', p.get('freq') or '', num(p.get('apps')),
                d(p.get('last')), d(p.get('caseDate')), None, next_text(p), clean(p.get('notes'))]
        for i, v in enumerate(vals, 1):
            ws.cell(r, i, v).font = F()
        ws.cell(r, 12).number_format = 'm/d/yy'
        ws.cell(r, 13).number_format = 'm/d/yy'
        ws.cell(r, 14, f'=IF(M{r}="","",IF(M{r}>TODAY(),"Upcoming",INT((TODAY()-M{r})/7)))')
    ws.freeze_panes = 'D2'
    ws.auto_filter.ref = f'A1:P{len(rows) + 1}'

    # One tab per account
    for name, title, accounts in TABS:
        if not any(p.get('account') in accounts for p in pts):
            continue
        tab = account_tab(wb, name, title, accounts, pts)
        if name == 'Good Samaritan':
            tab.page_setup.orientation = 'landscape'
            if to_order:
                to_order_block(tab, to_order)
    known = {a for _, _, accs in TABS for a in accs}
    for a in sorted({p.get('account') for p in pts} - known):
        account_tab(wb, a[:31], a, [a], pts)

    # Follow-ups: visits from last week's Monday on (logged result, else the booked appointment)
    start = today - dt.timedelta(days=today.weekday() + 7)
    fu = []
    for p in pts:
        logged = set()
        for e in p.get('log') or []:
            ed = d(e.get('date'))
            if not ed or ed < start:
                continue
            logged.add(ed)
            res = 'Grafted' if e.get('grafted') else 'Not grafted'
            if not e.get('grafted') and e.get('note'):
                res += ': ' + e['note'][:40]
            fu.append((ed, e.get('time') or '', e.get('loc') or '', p, e.get('size') or p.get('size') or '',
                       res, num(e.get('app')) if e.get('grafted') else None))
        nd = d(p.get('next'))
        if nd and nd >= start and nd not in logged and p.get('status') not in ('done', 'denied'):
            nxt = (num(p.get('apps')) or 0) + 1 if p.get('status') in ('active', 'approved') else None
            fu.append((nd, p.get('nextTime') or '', p.get('nextLoc') or '', p, p.get('size') or '',
                       LABEL.get(p.get('status'), ''), nxt))
    fu.sort(key=lambda x: (x[0], x[1], x[3].get('account') or '', x[3].get('initials') or ''))
    ws = wb.create_sheet('Follow-ups')
    ws.sheet_properties.tabColor = ORANGE[2:]
    header(ws, 1, ['Date', 'Time', 'Location', 'Account', 'Initials', 'Surgeon', 'Product', 'Size', 'Status',
                   'App #'], height=15)
    for i, w in enumerate([11, 10, 16, 30, 9, 16, 20, 13, 30, 7], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for r, (vd, tm, loc, p, size, st, app) in enumerate(fu, 2):
        vals = [vd, ftime(tm), loc, p.get('account') or '', p.get('initials') or '', p.get('provider') or '',
                p.get('product') or '', size, st, app]
        for i, v in enumerate(vals, 1):
            ws.cell(r, i, v).font = F()
        ws.cell(r, 1).number_format = 'm/d/yy'
    ws.freeze_panes = 'A2'

    # Application Log (newest first)
    log = [(d(e.get('date')), p, e) for p in pts for e in (p.get('log') or []) if d(e.get('date'))]
    log.sort(key=lambda x: (x[0], x[2].get('time') or ''), reverse=True)
    ws = wb.create_sheet('Application Log')
    ws.sheet_properties.tabColor = ORANGE[2:]
    header(ws, 1, ['Date', 'Account', 'Initials', 'Result', 'Size', 'App #', 'Covered by', 'Note'], height=15)
    for i, w in enumerate([10, 30, 9, 12, 14, 7, 11, 50], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for r, (ld, p, e) in enumerate(log, 2):
        vals = [ld, p.get('account') or '', p.get('initials') or '', 'Grafted' if e.get('grafted') else 'Not grafted',
                e.get('size') or '', num(e.get('app')), e.get('cover') or '', clean(e.get('note'))]
        for i, v in enumerate(vals, 1):
            ws.cell(r, i, v).font = F()
        ws.cell(r, 1).number_format = 'm/d/yy'

    # Day Notes (newest first)
    dn = [(d(n.get('date')), p, n) for p in pts for n in (p.get('dnotes') or []) if d(n.get('date'))]
    dn.sort(key=lambda x: x[0], reverse=True)
    ws = wb.create_sheet('Day Notes')
    ws.sheet_properties.tabColor = ORANGE[2:]
    header(ws, 1, ['Date', 'Account', 'Initials', 'Note'], height=15)
    for i, w in enumerate([10, 30, 9, 70], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for r, (nd, p, n) in enumerate(dn, 2):
        note = clean(n.get('note')) + (' [photo]' if n.get('asset') else '')
        vals = [nd, p.get('account') or '', p.get('initials') or '', note]
        for i, v in enumerate(vals, 1):
            c = ws.cell(r, i, v)
            c.font = F()
            c.alignment = Alignment(vertical='top', wrap_text=(i == 4))
        ws.cell(r, 1).number_format = 'm/d/yy'
        ws.row_dimensions[r].height = max(15, 12.75 * (len(note) // 95 + 1))

    wb.save(out)
    return len(pts), len(fu), len(log), len(dn)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('pdb')
    ap.add_argument('out')
    ap.add_argument('--today', default=None)
    ap.add_argument('--to-order', default=None)
    ap.add_argument('--prev', default=None)
    a = ap.parse_args()
    today = dt.date.fromisoformat(a.today) if a.today else dt.date.today()
    spec = json.load(open(a.to_order, encoding='utf-8')) if a.to_order else (prev_to_order(a.prev) if a.prev else None)
    n, nf, nl, nd = build(load(a.pdb), a.out, today, spec)
    print(f'{a.out}: {n} patients, {nf} follow-ups, {nl} log entries, {nd} day notes'
          + ('' if spec else ' (no To Order block)'))


if __name__ == '__main__':
    main()
