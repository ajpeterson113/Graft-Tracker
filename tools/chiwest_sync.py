#!/usr/bin/env python3
"""Build the Chi West app's data (db collection `data`) from the sources the other apps already keep.

    python3 -I tools/chiwest_sync.py --gt GTDB --patients Chi_West_Patient_List.xlsx \
        --targets Chi_West_Surgeon_Targets.xlsx --outreach TGTDB/outreach --cw CWDB OUT

GTDB      Graft Tracker app db download (or the build's OUT folder plus GTDB/todos): meta/snap.json,
          meta/forecast.json, meta/stock.json, months/*.json, todos/*.json
--patients  the patient list Excel the refresh just built (initials only, no names, no dollars)
--targets   Chi_West_Surgeon_Targets.xlsx from the repo
--outreach  Targets app `outreach` collection download (one json per target)
--cw        Chi West db download (data/*.json with their versions, actions/*, uploads/*); optional
OUT       writes OUT/data/<doc>.json and prints the ArtifactData batch(es) for the changed docs

Chi West (https://claude.ai/artifact/HzjrjSVDB4oF6oAfWUuwDN) only reads `data/*`; AJ's taps there go to
`actions` and `uploads`, which the refresh applies to the tracker / Patients app / Targets app, then marks
"applied" / "filed" (see chiwest/HOWTO.md).
"""
import argparse, datetime as dt, glob, json, os, re, sys
import openpyxl

SHORT = {
    'Alexian Brothers Wound Center': 'Alexian WCC', 'Alexian Ambulatory': 'Alexian Ambulatory',
    'Advocate Good Samaritan Hospital': 'Good Sam', 'Ascension Saint Alexius': 'St. Alexius',
    'Saint Joseph Hospital - Elgin': 'St. Joseph Elgin', 'Sherman': 'Sherman', 'Advocate Sherman Hospital': 'Sherman',
    'Central DuPage Foot and Ankle': 'Central DuPage F&A', 'Northwest Suburban Foot and Ankle Clinic': 'NW Suburban F&A',
    'Suburban Foot and Ankle Center': 'Suburban F&A', 'Romano Orthopaedics - Lombard': 'Romano Lombard',
    'Elmhurst Hospital': 'Elmhurst', 'UChicago Medicine AdventHealth GlenOaks': 'GlenOaks',
    'Ankle and Foot Surgery LLC': 'Ankle & Foot Surgery',
}
# target-list hospital (Contacts tab) -> account it belongs to in the tracker
HOSP_ACCT = {'Good Sam': 'Advocate Good Samaritan Hospital', 'St. Alexius / Alexian': 'Ascension Saint Alexius',
             'Elmhurst': 'Elmhurst Hospital', 'GlenOaks': 'UChicago Medicine AdventHealth GlenOaks',
             'Elgin': 'Saint Joseph Hospital - Elgin', 'Sherman': 'Sherman'}
STATUS = {'Current Patients': 'active', 'STRAVIX / OR Cases': 'or', 'Patients Done': 'done', 'On Hold': 'hold',
          'Pending IVRs': 'pending', 'Denied IVRs': 'denied', 'Potential New Starts (size = suggested)': 'approved',
          'Potential IVRs': 'potential'}


def slug(s):
    return re.sub(r'[^a-z0-9]+', '-', (s or '').lower()).strip('-')


def iso(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime('%Y-%m-%d')
    return ''


def cut(s, n):
    s = re.sub(r'\s+', ' ', str(s or '')).strip()
    return s if len(s) <= n else s[:n - 1].rstrip() + '…'


def hhmm(t):
    m = re.search(r'(\d{1,2}):(\d{2})\s*([AP]M)', str(t or ''), re.I)
    if not m:
        return ''
    h = int(m.group(1)) % 12 + (12 if m.group(3).upper() == 'PM' else 0)
    return f'{h:02d}:{m.group(2)}'


def next_appt(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return iso(v), ''
    m = re.search(r'(\d{1,2})/(\d{1,2})/(\d{2,4})', str(v or ''))
    if not m:
        return '', ''
    y = int(m.group(3)); y = y + 2000 if y < 100 else y
    return f'{y}-{int(m.group(1)):02d}-{int(m.group(2)):02d}', hhmm(v)


def load(path):
    return json.load(open(path)) if os.path.exists(path) else None


def loadall(d):
    out = {}
    for f in sorted(glob.glob(os.path.join(d, '*.json'))):
        out[os.path.basename(f)[:-5]] = json.load(open(f))
    return out


def rows(ws, start=2):
    hdr = [c.value for c in ws[start - 1]]
    for r in ws.iter_rows(min_row=start, values_only=True):
        if any(v not in (None, '') for v in r):
            yield dict(zip(hdr, r))


# ---------------------------------------------------------------- sources
def build_tracker(gt):
    snap = load(os.path.join(gt, 'meta', 'snap.json'))
    fc = load(os.path.join(gt, 'meta', 'forecast.json'))
    stock = load(os.path.join(gt, 'meta', 'stock.json'))
    months = sorted(loadall(os.path.join(gt, 'months')).values(), key=lambda m: m['ym'])
    todos = loadall(os.path.join(gt, 'todos'))
    if not snap or not months:
        sys.exit('STOP: --gt needs meta/snap.json and months/*.json')
    core = {k: snap.get(k) for k in ('updated', 'today', 'period', 'sales', 'forecast', 'accounts', 'year', 'months', 'periods', 'tabs')}
    core['todos'] = [dict(id=k, **{f: v.get(f) for f in ('text', 'detail', 'acct', 'ini', 'caseId', 'due', 'src', 'ref', 'created', 'status')})
                     for k, v in sorted(todos.items()) if v.get('status') != 'done']
    keep = ('id', 'date', 'ini', 'alx', 'acct', 'short', 'surgeon', 'prod', 'product', 'size', 'sz', 'qty', 'lots', 'price',
            'po', 'poDate', 'match', 'submitted', 'status', 'flags', 'outstanding', 'gh')
    cases = []
    for m in months[-3:]:
        for c in m['cases']:
            d = {k: c.get(k) for k in keep}
            d['notes'] = cut(c.get('notes'), 220)
            d['ym'], d['period'] = m['ym'], m.get('period')
            cases.append(d)
    items = []
    for i in (fc or {}).get('items', []):
        items.append({k: i.get(k) for k in ('acct', 'g', 'ini', 'surgeon', 'product', 'size', 'downSize', 'freq', 'apps', 'visits',
                                            'downVisits', 'up', 'down', 'start', 'ship', 'po')} | {'note': cut(i.get('note'), 180)})
    locs = []
    for l in (stock or {}).get('locations', []):
        locs.append({'name': l.get('name'), 'acct': l.get('acct'), 'kind': l.get('kind'), 'key': l.get('key'), 'counted': l.get('counted'),
                     'summary': l.get('summary'), 'about': cut(l.get('about'), 300), 'alerts': l.get('alerts', []),
                     'rules': cut(l.get('rules'), 300) if isinstance(l.get('rules'), str) else l.get('rules'),
                     'sizes': [{k: s.get(k) for k in ('product', 'size', 'onHand', 'min', 'max', 'usePerWeek', 'patients', 'runsOut',
                                                       'weeksLeft', 'order', 'status', 'pipeline', 'value', 'nextExp')} | {'note': cut(s.get('note'), 160)}
                               for s in l.get('sizes', [])],
                     'boxes': [{k: b.get(k) for k in ('product', 'size', 'lot', 'left', 'perBox', 'exp', 'flag')} for b in l.get('boxes', [])],
                     'incoming': l.get('incoming')})
    return core, {'cases': cases}, {'items': items, 'rules': (fc or {}).get('rules'), 'projectFrom': (fc or {}).get('projectFrom')}, {'locations': locs}


def build_patients(xlsx, today):
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    pts = []
    for r in rows(wb['All Patients']):
        nd, nt = next_appt(r.get('Next Appointment'))
        acct = r.get('Account') or ''
        pts.append({'id': r.get('ID') or slug(f"{acct}-{r.get('Initials')}"), 'acct': acct, 'short': SHORT.get(acct, acct),
                    'status': STATUS.get(r.get('Status'), 'active'), 'prov': r.get('Provider') or '', 'ini': r.get('Initials') or '',
                    'coverage': cut(r.get('Coverage Status'), 140), 'pa': r.get('PA Required?') or '', 'product': r.get('Product') or '',
                    'size': str(r.get('Size') or ''), 'freq': r.get('Frequency') or '', 'apps': r.get('Application Count') or 0,
                    'last': iso(r.get('Last Graft')), 'caseDate': iso(r.get('Case Date')), 'next': nd, 'time': nt,
                    'notes': cut(r.get('Notes'), 420)})
    t0 = (dt.date.fromisoformat(today) - dt.timedelta(days=45)).isoformat()
    visits = []
    for r in rows(wb['Follow-ups']):
        d = iso(r.get('Date'))
        if d >= t0:
            visits.append({'date': d, 'time': hhmm(r.get('Time')), 'loc': r.get('Location') or '', 'acct': r.get('Account') or '',
                           'ini': r.get('Initials') or '', 'prov': r.get('Surgeon') or '', 'product': r.get('Product') or '',
                           'size': str(r.get('Size') or ''), 'status': r.get('Status') or '', 'app': r.get('App #')})
    log = []
    for r in rows(wb['Application Log']):
        d = iso(r.get('Date'))
        if d >= t0:
            log.append({'date': d, 'acct': r.get('Account') or '', 'ini': r.get('Initials') or '', 'result': r.get('Result') or '',
                        'size': str(r.get('Size') or ''), 'app': r.get('App #'), 'cover': r.get('Covered by') or '', 'note': cut(r.get('Note'), 200)})
    notes = []
    for r in rows(wb['Day Notes']):
        d = iso(r.get('Date'))
        if d >= t0:
            notes.append({'date': d, 'acct': r.get('Account') or '', 'ini': r.get('Initials') or '', 'note': cut(r.get('Note'), 240)})
    return {'patients': pts, 'visits': visits, 'log': log, 'notes': notes}


def build_targets(xlsx, outreach_dir):
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    byh = list(rows(wb['By hospital']))
    con = list(rows(wb['Contacts & emails']))
    cmap = {(c.get('Name'), c.get('Credentials')): c for c in con}
    orx = loadall(outreach_dir) if outreach_dir and os.path.isdir(outreach_dir) else {}
    out, seen = [], set()
    for r in byh:
        c = cmap.get((r.get('Name'), r.get('Credentials'))) or {}
        hosp = c.get('Hospital') or r.get('Hospital') or ''
        tid = slug(f"{r.get('Name')}-{hosp}")
        seen.add(tid)
        o = orx.get(tid, {})
        hist = [{'d': (h.get('d') or '')[:10], 'a': h.get('a'), 'n': cut(h.get('n'), 200), 'who': h.get('who') or ''} for h in (o.get('history') or [])][-6:]
        out.append({'id': tid, 'name': r.get('Name') or '', 'cred': r.get('Credentials') or '', 'tier': str(r.get('Tier') or ''),
                    'hosp': hosp, 'hospFull': r.get('Hospital') or '', 'acct': HOSP_ACCT.get(hosp, ''), 'grp': r.get('Specialty group') or '',
                    'spec': r.get('Specialty') or '', 'practice': cut(r.get('Practice'), 120), 'office': cut(r.get('Office(s)'), 160),
                    'phone': r.get('Phone') or '', 'where': cut(r.get('Where to find / clinic days'), 160), 'fit': cut(r.get('Case types / biologics fit'), 200),
                    'role': cut(r.get('Role / program'), 160), 'angle': cut(r.get('Why / angle'), 320), 'notes': cut(r.get('Notes'), 200),
                    'customer': r.get('Existing customer') or '', 'score': r.get('Claims score'), 'wound': r.get('Wound claims'),
                    'orc': r.get('OR/surgical claims'), 'qsn': r.get('Q-codes S+N'), 'qcomp': r.get('Q-codes competitor'), 'npi': r.get('NPI') or '',
                    'email': c.get('Email (published)') or '', 'emailFmt': cut(c.get('Email format (system)'), 80),
                    'status': o.get('status') or 'none', 'followUp': o.get('followUp') or '', 'onote': cut(o.get('note'), 400),
                    'obj': cut(o.get('objection') or o.get('objections'), 200), 'hist': hist})
    for tid, o in orx.items():  # contacts added in the Targets app (e.g. cwl-*)
        if tid in seen:
            continue
        e = o.get('edits') or {}
        out.append({'id': tid, 'name': e.get('name') or tid.replace('cwl-', '').replace('-', ' ').title(), 'cred': '', 'tier': '',
                    'hosp': e.get('hospital') or 'NM CDH / Delnor expansion' if 'cdh' in tid or tid.startswith('cwl-') else '', 'acct': '',
                    'grp': 'Office / scheduling contact', 'spec': e.get('role') or '', 'office': cut(e.get('officeContact'), 160),
                    'status': o.get('status') or 'none', 'followUp': o.get('followUp') or '', 'onote': cut(o.get('note'), 400), 'obj': '',
                    'hist': [{'d': (h.get('d') or '')[:10], 'a': h.get('a'), 'n': cut(h.get('n'), 200), 'who': h.get('who') or ''} for h in (o.get('history') or [])][-6:]})
    return {'targets': out}


# ---------------------------------------------------------------- writes
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--gt', required=True); ap.add_argument('--patients', required=True)
    ap.add_argument('--targets', required=True); ap.add_argument('--outreach'); ap.add_argument('--cw')
    ap.add_argument('out')
    a = ap.parse_args()
    core, cases, fc, stock = build_tracker(a.gt)
    today = core.get('today') or dt.date.today().isoformat()
    docs = {'core': core, 'cases': cases, 'forecast': fc, 'stock': stock,
            'patients': build_patients(a.patients, today), 'targets': build_targets(a.targets, a.outreach)}
    stamp = dt.datetime.now(dt.timezone(dt.timedelta(hours=-6))).isoformat(timespec='seconds')
    os.makedirs(os.path.join(a.out, 'data'), exist_ok=True)
    prev = loadall(os.path.join(a.cw, 'data')) if a.cw else {}
    vers = load(os.path.join(a.cw, 'versions.json')) if a.cw else None
    writes = []
    for k, d in docs.items():
        old = prev.get(k)
        same = old is not None and {x: y for x, y in old.items() if x != 'synced'} == json.loads(json.dumps(d))
        d['synced'] = stamp
        p = os.path.join(a.out, 'data', k + '.json')
        json.dump(d, open(p, 'w'), separators=(',', ':'))
        n = os.path.getsize(p)
        print(f'data/{k}: {n:,} bytes' + ('  <-- over 240 KB' if n > 240000 else '') + ('  (unchanged)' if same else ''))
        if not same:
            w = {'op': 'set', 'collection': 'data', 'doc_id': k, 'file_path': os.path.abspath(p)}
            if vers and k in vers:
                w['if_version'] = vers[k]
            elif old is not None:
                w['if_version'] = 'VERSION'
            writes.append(w)
    # core last, so the app's "synced" time only moves once everything else is in
    writes.sort(key=lambda w: w['doc_id'] == 'core')
    tot = 0; batch = []; batches = [batch]
    for w in writes:
        s = os.path.getsize(w['file_path'])
        if batch and tot + s > 900000:
            batch = []; batches.append(batch); tot = 0
        batch.append(w); tot += s
    json.dump([b for b in batches if b], open(os.path.join(a.out, 'batches.json'), 'w'), indent=1)
    print(f'{len(writes)} doc(s) to write in {len([b for b in batches if b])} batch(es): {os.path.join(a.out, "batches.json")}')


if __name__ == '__main__':
    main()
