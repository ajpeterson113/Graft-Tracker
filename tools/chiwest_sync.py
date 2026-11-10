#!/usr/bin/env python3
"""Build the Chi West app's data (db collection `data`) from the sources the other apps already keep.

    python3 -I tools/chiwest_sync.py --gt GTDB --patients Chi_West_Patient_List.xlsx \
        [--pdb PDB | --overlay OVERLAY.json] --targets Chi_West_Surgeon_Targets.xlsx \
        --outreach TGTDB/outreach --cw CWDB OUT

GTDB      Graft Tracker app db download (or the build's OUT folder plus GTDB/todos): meta/snap.json,
          meta/forecast.json, meta/stock.json, months/*.json, todos/*.json
--patients  the patient list Excel the refresh just built (initials, notes, visits, log, day notes)
--pdb       Patients app db export (ArtifactData list with out_dir: PDB/patients/*.json, PDB/events/*.json).
            Adds what the Excel leaves out: names, who covers each visit, next-visit location, verify flags,
            and AJ's events. Preferred when the export is allowed.
--overlay   when the export isn't allowed: a JSON file {"patients":[{id,name,cover,status,apps,next,nextTime,
            nextLoc,verify,fuVerify,ivrDate,paExp}], "events":[{id,title,kind,place,date,time,end,note}]} written from an
            inline ArtifactData list of the same collections.
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
    'Alexian Brothers Medical Center': 'Alexian Main OR',  # the hospital's main OR, its own account (AJ 11/9)
}
CANON_ACCT = {'Alexian Brothers - Ambulatory': 'Alexian Ambulatory', 'Advocate Sherman Hospital': 'Sherman'}
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
def incoming_of(s):
    """Grafts on the way to a shelf: an OR shelf's Incoming units, or a restock already shipped to the Alexian WCC
    shelf (the Stock tab's On order column, carried in the size note as "On order: N ..."). AJ 11/6."""
    if s.get('incoming'):
        return int(s['incoming'])
    m = re.match(r'\s*On order:\s*(\d+)', str(s.get('note') or ''))
    return int(m.group(1)) if m else 0


def arriving_of(s):
    m = re.search(r'arriving ([A-Z][a-z]{2} \d{1,2}/\d{1,2})', str(s.get('note') or ''))
    return m.group(1) if m else ''


def alerts_with_incoming(loc):
    """AJ 11/6: no reorder or low-stock alert for a size that already has stock on the way (expiry alerts stay)."""
    out = []
    for a in loc.get('alerts', []) or []:
        if a.get('level') != 'bad' and any(incoming_of(s) and str(a.get('title', '')).startswith(f"{s.get('product')} {s.get('size')}:")
                                           for s in loc.get('sizes', []) or []):
            continue
        out.append(dict(a))
    return out


def stock_users(stock, patients, today=None):
    """Who is on each size of a wound-care-center shelf (AJ 11/6), from Chi West's own patient list: Alexian WCC and
    Good Sam WCC. Each size gets `users` [{ini, status, freq, qty, next}]; a size with no use per week from the tracker
    (Good Sam WCC) gets it from its grafting patients' booked visits: `useWeeks` = whole grafts in this week and the next
    four (weekly patients every week, every-2-week patients every other week from their next visit), and use per week =
    next week's grafts (AJ 11/6: never a fraction of a graft)."""
    t = dt.date.fromisoformat(str(today or dt.date.today().isoformat())[:10])
    ws = t - dt.timedelta(days=t.weekday())
    shelf = {'alexian': 'Alexian Brothers Wound Center', 'goodsamwcc': 'Advocate Good Samaritan Hospital'}
    fam = lambda p: ('OASIS Ultra' if re.search(r'ultra', p, re.I) else 'OASIS Burn' if re.search(r'burn', p, re.I) else
                     'OASIS Wound' if re.search(r'oasis', p, re.I) else 'Grafix' if re.search(r'grafix', p, re.I) else '')
    def parts(p):
        prod, out = fam(p.get('product') or ''), []
        for chunk in re.split(r'\s*\+\s*|,\s*|\band\b', str(p.get('size') or '')):
            m = re.search(r'(\d+(?:\.\d+)?\s*x\s*\d+(?:\.\d+)?|\d+\s*mm)', chunk, re.I)
            if not m:
                continue
            q = re.search(r'\((?:x\s*)?(\d+)(?:\s*sheets?)?\)|\b(\d+)x\s*(?=\d+\s*x)|^\s*\((\d+)\)', chunk, re.I)
            n = next((int(g) for g in (q.groups() if q else ()) if g), 1)
            f = fam(chunk) or prod
            out.append((f, re.sub(r'\s+', '', m.group(1).lower()), n))
        return out
    for loc in stock.get('locations', []):
        acct = shelf.get(loc.get('key'))
        if not acct:
            continue
        ppl = [p for p in patients.get('patients', []) if p.get('acct') == acct and p.get('status') in ('active', 'approved', 'pending')]
        for s in loc.get('sizes', []):
            sz = re.sub(r'\s+', '', str(s.get('size') or '').lower())
            users, weeks, who = [], [0] * 5, [[] for _ in range(5)]
            for p in ppl:
                for f, size, n in parts(p):
                    if f == s.get('product') and (sz == size or sz.startswith(size)):
                        step = 14 if re.search(r'2 weeks|biweekly', str(p.get('freq') or ''), re.I) else 7
                        users.append({'ini': p.get('ini'), 'status': p.get('status'), 'freq': p.get('freq') or '', 'qty': n, 'next': p.get('next') or ''})
                        try:
                            d = dt.date.fromisoformat(str(p.get('next'))[:10])
                        except ValueError:
                            continue
                        if p.get('status') != 'active':
                            continue
                        while d < t:
                            d += dt.timedelta(days=step)
                        while (d - ws).days < 35:
                            weeks[(d - ws).days // 7] += n
                            who[(d - ws).days // 7].append(f"{p.get('ini')}" + (f" x{n}" if n > 1 else ''))
                            d += dt.timedelta(days=step)
            s['users'] = users
            s['useWeeks'] = weeks
            s['useWho'] = who
            if s.get('usePerWeek') in (None, '') and loc.get('key') == 'goodsamwcc':
                s['usePerWeek'] = weeks[1]
                s['patients'] = sum(1 for u in users if u['status'] == 'active')
    return stock


def build_tracker(gt):
    snap = load(os.path.join(gt, 'meta', 'snap.json'))
    fc = load(os.path.join(gt, 'meta', 'forecast.json'))
    stock = load(os.path.join(gt, 'meta', 'stock.json'))
    months = sorted(loadall(os.path.join(gt, 'months')).values(), key=lambda m: m['ym'])
    todos = loadall(os.path.join(gt, 'todos'))
    if not snap or not months:
        sys.exit('STOP: --gt needs meta/snap.json and months/*.json')
    core = {k: snap.get(k) for k in ('updated', 'today', 'period', 'sales', 'forecast', 'accounts', 'year', 'months', 'periods', 'tabs', 'weekly')}
    core['todos'] = [dict(id=k, **{f: v.get(f) for f in ('text', 'detail', 'acct', 'ini', 'caseId', 'due', 'src', 'ref', 'created', 'status')})
                     for k, v in sorted(todos.items()) if v.get('status') != 'done']
    keep = ('id', 'date', 'ini', 'alx', 'acct', 'short', 'surgeon', 'prod', 'product', 'size', 'sz', 'qty', 'lots', 'price',
            'po', 'poDate', 'match', 'submitted', 'status', 'flags', 'outstanding', 'gh')
    cases = []
    for m in months[-3:]:
        for c in m['cases']:
            d = {k: c.get(k) for k in keep}
            d['notes'] = cut(c.get('notes'), 220)
            # the case's own files (case sheet, transport log) by file name; Chi West opens them from its copies listed in
            # its data/files doc (tools/chiwest_files.py keeps those in step with GT, AJ 11/5)
            d['files'] = [{'f': f.get('f'), 'kind': f.get('kind') or 'Case sheet'} for f in (c.get('files') or [])
                          if str(f.get('f') or '').lower().endswith('.jpg')]
            d['ym'], d['period'] = m['ym'], m.get('period')
            cases.append(d)
    items = []
    for i in (fc or {}).get('items', []):
        items.append({k: i.get(k) for k in ('acct', 'g', 'ini', 'surgeon', 'product', 'size', 'downSize', 'freq', 'apps', 'visits',
                                            'downVisits', 'up', 'down', 'start', 'ship', 'po')} | {'note': cut(i.get('note'), 180)})
    locs = []
    for l in (stock or {}).get('locations', []):
        locs.append({'name': l.get('name'), 'acct': l.get('acct'), 'kind': l.get('kind'), 'key': l.get('key'), 'counted': l.get('counted'),
                     'summary': l.get('summary'), 'about': cut(l.get('about'), 300), 'alerts': alerts_with_incoming(l),
                     'rules': cut(l.get('rules'), 300) if isinstance(l.get('rules'), str) else l.get('rules'),
                     'sizes': [{k: s.get(k) for k in ('product', 'size', 'onHand', 'min', 'max', 'usePerWeek', 'patients', 'runsOut',
                                                       'weeksLeft', 'order', 'status', 'pipeline', 'value', 'nextExp')} | {'note': cut(s.get('note'), 160), 'incoming': incoming_of(s), 'arriving': arriving_of(s)}
                               for s in l.get('sizes', [])],
                     'boxes': [{k: b.get(k) for k in ('product', 'size', 'lot', 'left', 'perBox', 'exp', 'flag')} for b in l.get('boxes', [])],
                     'incoming': l.get('incoming')})
    return core, {'cases': cases}, {'items': items, 'rules': (fc or {}).get('rules'), 'projectFrom': (fc or {}).get('projectFrom')}, {'locations': locs}


def safe_name(n):
    """Drop anything like an insurance or birth date tacked onto the name."""
    n = re.sub(r'\s*\(.*$', '', str(n or '')).strip()
    return re.sub(r'\s+', ' ', n)


# fields a patient needs when it is in the Patients app but not yet in the patient list Excel (a new patient AJ just
# added): overlay key -> Patients app field
NEW_FIELDS = {'acct': 'account', 'prov': 'provider', 'freq': 'freq', 'notes': 'notes', 'coverage': 'coverage',
              'last': 'last', 'caseDate': 'caseDate', 'pa': 'pa'}


def dnotes_of(d):
    """Notes AJ added on a patient in the Patients app (dnotes), text only (photo notes keep their caption)."""
    out = []
    for n in d.get('dnotes') or []:
        if isinstance(n, dict) and str(n.get('note') or '').strip():
            out.append({'date': n.get('date') or '', 'note': cut(n.get('note'), 300)})
    return out[-12:]


def labels_of(d):
    """AJ's label for a booked visit, by date (Patients app `visitLabels`, e.g. {"2026-11-13": "Stravix follow up"}):
    the schedule shows the label instead of an application number (AJ 11/6)."""
    v = d.get('visitLabels') or d.get('labels') or {}
    return {str(k)[:10]: cut(str(t), 60) for k, t in v.items() if str(t or '').strip()} if isinstance(v, dict) else {}


def load_extra(pdb, overlay):
    """Names, coverage, next-visit location, verify flags and events, by patient id."""
    ex, events = {}, []
    if pdb and os.path.isdir(os.path.join(pdb, 'patients')):
        for pid, d in loadall(os.path.join(pdb, 'patients')).items():
            d = d.get('data', d)
            ex[pid] = {'name': safe_name(d.get('name')), 'cover': d.get('cover') or '', 'status': d.get('status') or '',
                       'apps': d.get('apps'), 'next': d.get('next') or '', 'nextTime': d.get('nextTime') or '', 'nextLoc': d.get('nextLoc') or '',
                       'verify': cut(d.get('verify'), 120), 'fuVerify': cut(d.get('fuVerify'), 120), 'ini': d.get('initials') or '',
                       'size': d.get('size') or '', 'product': d.get('product') or '', 'log': d.get('log') or [],
                       'ivrDate': d.get('ivrDate') or '', 'paExp': d.get('paExp') or ''}
            ex[pid].update({k: d.get(src) or '' for k, src in NEW_FIELDS.items()})
            ex[pid]['dnotes'] = dnotes_of(d)
            ex[pid]['labels'] = labels_of(d)
        for eid, d in loadall(os.path.join(pdb, 'events')).items():
            d = d.get('data', d)
            events.append({'id': eid, 'title': d.get('title') or d.get('kind') or 'Event', 'kind': d.get('kind') or 'Other', 'place': d.get('place') or '',
                           'date': d.get('date') or '', 'time': d.get('time') or '', 'end': d.get('end') or '', 'note': cut(d.get('note'), 200)})
    elif overlay and os.path.exists(overlay):
        o = json.load(open(overlay))
        for d in o.get('patients', []):
            ex[d['id']] = {k: (safe_name(d.get(k)) if k == 'name' else d.get(k, '')) for k in ('name', 'cover', 'status', 'apps', 'next', 'nextTime', 'nextLoc', 'verify', 'fuVerify', 'ini', 'size', 'product', 'ivrDate', 'paExp')}
            ex[d['id']]['ini'] = ex[d['id']]['ini'] or d.get('initials') or ''
            ex[d['id']]['log'] = d.get('log') or []
            ex[d['id']].update({k: d.get(k) or d.get(src) or '' for k, src in NEW_FIELDS.items()})
            ex[d['id']]['dnotes'] = dnotes_of(d)
            ex[d['id']]['labels'] = labels_of(d)
        events = [{'id': e.get('id'), 'title': e.get('title') or 'Event', 'kind': e.get('kind') or 'Other', 'place': e.get('place') or '', 'date': e.get('date') or '',
                   'time': e.get('time') or '', 'end': e.get('end') or '', 'note': cut(e.get('note'), 200)} for e in o.get('events', [])]
    return ex, events


def build_patients(xlsx, today, extra=None, events=None):
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    extra = extra or {}
    pts = []
    for r in rows(wb['All Patients']):
        nd, nt = next_appt(r.get('Next Appointment'))
        acct = r.get('Account') or ''
        pid = r.get('ID') or slug(f"{acct}-{r.get('Initials')}")
        x = extra.get(pid, {})
        if x.get('next') or x.get('next') == '' and x:
            nd, nt = x.get('next') or '', x.get('nextTime') or ''
        pts.append({'id': pid, 'acct': acct, 'short': SHORT.get(acct, acct),
                    'status': x.get('status') or STATUS.get(r.get('Status'), 'active'), 'prov': r.get('Provider') or '', 'ini': x.get('ini') or r.get('Initials') or '',
                    'name': x.get('name') or '', 'cover': x.get('cover') if x.get('cover') in ('Caroline', 'Pete', 'Not covered') else '',
                    'coverage': cut(r.get('Coverage Status'), 140), 'pa': r.get('PA Required?') or '', 'product': x.get('product') or r.get('Product') or '',
                    'size': str(x.get('size') or r.get('Size') or ''), 'freq': r.get('Frequency') or '', 'apps': (x.get('apps') if x.get('apps') is not None else r.get('Application Count')) or 0,
                    'last': iso(r.get('Last Graft')), 'caseDate': iso(r.get('Case Date')), 'next': nd, 'time': nt, 'loc': x.get('nextLoc') or '',
                    'verify': x.get('verify') or '', 'fuVerify': x.get('fuVerify') or '', 'notes': cut(r.get('Notes'), 420),
                    'ivrDate': x.get('ivrDate') or '', 'paExp': x.get('paExp') or '', 'dnotes': x.get('dnotes') or [],
                    'labels': x.get('labels') or {}})
        # the Patients app is the source of truth: a field AJ edited there (or through a Chi West tap) wins over the patient
        # list Excel, which only changes when the morning update rebuilds it (AJ 11/6: what he changes shows and stays)
        for k, lim in (('prov', 0), ('freq', 0), ('coverage', 140), ('pa', 0), ('caseDate', 0), ('last', 0), ('notes', 420)):
            if x.get(k) not in (None, ''):
                pts[-1][k] = cut(x[k], lim) if lim else x[k]
    # patients in the Patients app that the patient list Excel doesn't have yet (AJ 11/5: a new patient he adds
    # shows up in Chi West at the next sync, not only after the Excel is rebuilt)
    have = {p['id'] for p in pts}
    for pid, x in extra.items():
        if pid in have or not x.get('acct'):
            continue
        acct = CANON_ACCT.get(x['acct'], x['acct'])
        pts.append({'id': pid, 'acct': acct, 'short': SHORT.get(acct, acct), 'status': x.get('status') or 'pending',
                    'prov': x.get('prov') or '', 'ini': x.get('ini') or '', 'name': x.get('name') or '',
                    'cover': x.get('cover') if x.get('cover') in ('Caroline', 'Pete', 'Not covered') else '',
                    'coverage': cut(x.get('coverage'), 140), 'pa': x.get('pa') or '', 'product': x.get('product') or '',
                    'size': str(x.get('size') or ''), 'freq': x.get('freq') or '', 'apps': x.get('apps') or 0,
                    'last': x.get('last') or '', 'caseDate': x.get('caseDate') or '', 'next': x.get('next') or '',
                    'time': x.get('nextTime') or '', 'loc': x.get('nextLoc') or '', 'verify': x.get('verify') or '',
                    'fuVerify': x.get('fuVerify') or '', 'notes': cut(x.get('notes'), 420),
                    'ivrDate': x.get('ivrDate') or '', 'paExp': x.get('paExp') or '', 'dnotes': x.get('dnotes') or [],
                    'labels': x.get('labels') or {}})
    t0 = (dt.date.fromisoformat(today) - dt.timedelta(days=63)).isoformat()
    acct_of = {p['id']: p['acct'] for p in pts}
    ini_of = {p['id']: p['ini'] for p in pts}
    dbl = {}  # (acct, initials, date) -> db log entry
    for pid, x in extra.items():
        for e in x.get('log') or []:
            # AJ 11/6: a result dated after today (a no-show logged on the next booked visit) doesn't close that visit
            if t0 <= (e.get('date') or '') <= today and pid in acct_of:
                dbl[(acct_of[pid], re.sub(r'[^A-Za-z]', '', ini_of[pid] or '').upper(), e['date'])] = (pid, e)
    visits = []
    for r in rows(wb['Follow-ups']):
        d = iso(r.get('Date'))
        if d >= t0:
            visits.append({'date': d, 'time': hhmm(r.get('Time')), 'loc': r.get('Location') or '', 'acct': r.get('Account') or '',
                           'ini': r.get('Initials') or '', 'prov': r.get('Surgeon') or '', 'product': r.get('Product') or '',
                           'size': str(r.get('Size') or ''), 'status': r.get('Status') or '', 'app': r.get('App #')})
    log = []
    seen = set()
    for r in rows(wb['Application Log']):
        d = iso(r.get('Date'))
        key = (r.get('Account') or '', re.sub(r'[^A-Za-z]', '', str(r.get('Initials') or '')).upper(), d)
        if t0 <= d <= today and key not in dbl:
            seen.add(key)
            log.append({'date': d, 'acct': r.get('Account') or '', 'ini': r.get('Initials') or '', 'result': r.get('Result') or '',
                        'size': str(r.get('Size') or ''), 'app': r.get('App #'), 'cover': r.get('Covered by') or '', 'note': cut(r.get('Note'), 200)})
    for (acct, _ini, d), (pid, e) in dbl.items():  # entries the Excel doesn't have yet (or has older)
        log.append({'date': d, 'acct': acct, 'ini': ini_of[pid], 'result': 'Grafted' if e.get('grafted') else 'Not grafted', 'size': str(e.get('size') or ''),
                    'app': e.get('app'), 'cover': e.get('cover') or '', 'note': cut(e.get('note'), 200), 'time': e.get('time') or '', 'loc': e.get('loc') or ''})
        st = ('Grafted' if e.get('grafted') else 'Not grafted') + (': ' + e['note'] if e.get('note') and not e.get('grafted') else '')
        hit = [v for v in visits if v['acct'] == acct and (re.sub(r'[^A-Za-z]', '', v['ini']).upper() == _ini or v['ini'] in ('', '??')) and v['date'] == d and (v['time'] == (e.get('time') or '') or not v['time'])]
        if hit:
            for v in hit:
                v.update({'status': st, 'ini': ini_of[pid], 'app': e.get('app') or v.get('app'), 'size': str(e.get('size') or v.get('size') or ''), 'time': e.get('time') or v.get('time') or ''})
        else:
            visits.append({'date': d, 'time': e.get('time') or '', 'loc': e.get('loc') or '', 'acct': acct, 'ini': ini_of[pid], 'prov': '', 'product': '',
                           'size': str(e.get('size') or ''), 'status': st, 'app': e.get('app')})
    # A# on the upcoming visits follows the Patients app's application count (AJ's A# taps and graft logs land there),
    # not the patient list Excel's App #, which only changes when the morning tracker update rebuilds it (AJ 11/5)
    norm = lambda s: re.sub(r'[^A-Za-z]', '', str(s or '')).upper()
    active = {}
    for p in pts:
        if p['status'] == 'active':
            active.setdefault((p['acct'], norm(p['ini'])), []).append(p)
    for (acct, ini), ps in active.items():
        x = extra.get(ps[0]['id'], {})
        if len(ps) > 1 or x.get('apps') in (None, ''):  # two grafting patients share account + initials: leave the Excel's numbers
            continue
        fut = sorted([v for v in visits if v['acct'] == acct and norm(v['ini']) == ini and v['date'] >= today
                      and not re.match(r'^(Grafted|Not grafted)', str(v.get('status') or ''))], key=lambda v: (v['date'], v['time'] or ''))
        for i, v in enumerate(fut):
            v['app'] = int(x['apps']) + 1 + i
    log.sort(key=lambda x: x['date'], reverse=True)
    notes = []
    for r in rows(wb['Day Notes']):
        d = iso(r.get('Date'))
        if d >= t0:
            notes.append({'date': d, 'acct': r.get('Account') or '', 'ini': r.get('Initials') or '', 'note': cut(r.get('Note'), 240)})
    return {'patients': pts, 'visits': visits, 'log': log, 'notes': notes, 'events': events or [], 'names': bool(extra)}


# VERSAJET (AJ 11/9): AJ also sells VERSAJET, for now only at Alexian Brothers (Elk Grove Village). Targets are general,
# vascular, podiatry, plastics and anyone else who debrides in the OR; it is also a way to get tissue in with the same
# provider. 'yes' = tied to Alexian Brothers, 'check' = in the St. Alexius / Alexian area, Alexian privileges not confirmed.
# Same rule as vjOf() in the Targets app page; when another account opens up, change both.
VJ_GROUPS = {'Podiatry / foot & ankle', 'Vascular surgery', 'Plastics / reconstruction', 'General / trauma / colorectal',
             'Orthopedics (trauma / F&A)', 'Wound center clinician'}
VJ_AT = re.compile(r'alexian|biesterfield|beisner|elk grove', re.I)


def versajet(r, hosp):
    if (r.get('Specialty group') or '') not in VJ_GROUPS:
        return ''
    where = ' '.join(str(r.get(k) or '') for k in ('Hospital affiliation (source)', 'Office(s)', 'Practice', 'Existing customer',
                                                   'Where to find / clinic days', 'Role / program'))
    if VJ_AT.search(where):
        return 'yes'
    return 'check' if 'Alexi' in f"{hosp} {r.get('Hospital') or ''}" else ''


# Hospital affiliation filter (AJ 11/9): each target's hospitals, from the workbook's "Hospital affiliation (source)" text.
# [short name, how sure]: "y" = the text names the hospital, "l" = named but hedged (verify / likely / presumed / claims),
# "a" = nothing named, so the hospitals of the target's area are assumed. A hospital the text rules out ("Good Sam NOT listed",
# "not GlenOaks") is left off, even as an area guess.
HOSPITALS = [('Good Sam', r'good sam'), ('GlenOaks', r'glen\s?oaks'), ('St. Alexius', r'(?:saint|st\.?)\s+alexius'),
             ('Alexian', r'alexian'), ('Elmhurst', r'elmhurst'), ('Sherman', r'sherman'),
             ('St. Joe Elgin', r'(?:saint|st\.?)\s+jo(?:seph|e)\b(?!\s*[-\u2013\u2014]?\s*\(?chicago)'),
             ('CDH', r'central dupage|\bcdh\b'), ('Delnor', r'delnor')]
_HEDGE = re.compile(r'verify|likely|presum|assum|suggest|claims|legacy|former|prior|probabl|unconfirm|inferred|appears', re.I)
_NEG_AFTER = re.compile(r'^[^;.]{0,45}?(?:\bNOT\b|\bnot (?:listed|found|verifi|confirmed|stated)|unverified|\b404)', re.I)
_NEG_BEFORE = re.compile(r"\b(?:no|not)\s+[\w\s/&'-]{0,22}$", re.I)


def hospitals(affil, cluster):
    """[[short, 'y'|'l'|'a'], ...] for one target; see HOSPITALS above."""
    text = str(affil or '')
    elgin = bool(re.search(r'elgin', f'{cluster} {text}', re.I))
    found, ruled_out = {}, set()
    for clause in re.split(r';|(?<!\bSt)(?<!\bDr)\.\s|\s[-\u2013\u2014]\s|\s->\s', text):
        for name, pat in HOSPITALS:
            for m in re.finditer(pat, clause, re.I):
                if name == 'St. Joe Elgin' and not (elgin or re.search(r'airlite|prime', clause, re.I)):
                    continue
                if _NEG_BEFORE.search(clause[:m.start()]) or _NEG_AFTER.search(clause[m.end():]):
                    ruled_out.add(name)
                    continue
                sure = 'l' if _HEDGE.search(clause) else 'y'
                if found.get(name) != 'y':
                    found[name] = sure
    for name in ruled_out:
        found.pop(name, None) if found.get(name) != 'y' else None
    if not found:
        area = [n for n, pat in HOSPITALS if re.search(pat, str(cluster or ''), re.I)]
        if re.search(r'elgin', str(cluster or ''), re.I):
            area = ['Sherman', 'St. Joe Elgin']
        found = {n: 'a' for n in area if n not in ruled_out}
    order = [n for n, _ in HOSPITALS]
    return [[n, found[n]] for n in order if n in found]


def build_targets(xlsx, outreach_dir, plans_dir=None):
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
                    'obj': cut(o.get('objection') or o.get('objections'), 200), 'hist': hist,
                    'lunch': o.get('lunch') or None, 'using': o.get('using') or [],
                    'fav': bool(o.get('fav')), 'assoc': bool(o.get('assoc'))})  # AJ 11/9: Set Target / Assign to Associates
        out[-1]['vj'] = versajet(r, hosp)
        out[-1]['aff'] = hospitals(r.get('Hospital affiliation (source)'), f"{hosp} {r.get('Hospital') or ''}")
        if (o.get('edits') or {}).get('office'):  # office corrected in the Targets app
            out[-1]['office'] = cut(o['edits']['office'], 160)
    for tid, o in orx.items():  # contacts added in the Targets app (e.g. cwl-*)
        if tid in seen:
            continue
        e = o.get('edits') or {}
        a = o.get('add') or {}  # a target added from a Chi West `newtarget` tap (AJ 11/9); the same fields are in TGT custom/<id>
        if a:
            out.append({'id': tid, 'name': a.get('name') or tid, 'cred': a.get('cred') or '', 'tier': str(a.get('tier') or '2'),
                        'hosp': a.get('hosp') or '', 'hospFull': a.get('hosp') or '', 'acct': HOSP_ACCT.get(a.get('hosp') or '', ''),
                        'grp': a.get('group') or 'Other', 'spec': a.get('spec') or '', 'practice': cut(a.get('practice'), 120),
                        'office': cut(e.get('office') or a.get('office'), 160), 'phone': a.get('phone') or '', 'email': a.get('email') or '',
                        'angle': cut(a.get('why'), 320), 'notes': cut(a.get('notes'), 200), 'npi': a.get('npi') or '',
                        'status': o.get('status') or 'none', 'followUp': o.get('followUp') or '', 'onote': cut(o.get('note'), 400),
                        'obj': cut(o.get('objection') or o.get('objections'), 200), 'lunch': o.get('lunch') or None, 'using': o.get('using') or [],
                        'fav': bool(o.get('fav')), 'assoc': bool(o.get('assoc')),
                        'vj': versajet({'Specialty group': a.get('group'), 'Hospital affiliation (source)': a.get('affil'), 'Office(s)': a.get('office'),
                                        'Practice': a.get('practice'), 'Hospital': a.get('hosp')}, a.get('hosp') or ''),
                        'aff': hospitals(a.get('affil'), a.get('hosp') or ''),
                        'hist': [{'d': (h.get('d') or '')[:10], 'a': h.get('a'), 'n': cut(h.get('n'), 200), 'who': h.get('who') or ''} for h in (o.get('history') or [])][-6:]})
            continue
        out.append({'id': tid, 'name': e.get('name') or tid.replace('cwl-', '').replace('-', ' ').title(), 'cred': '', 'tier': '',
                    'hosp': e.get('hospital') or 'NM CDH / Delnor expansion' if 'cdh' in tid or tid.startswith('cwl-') else '', 'acct': '',
                    'fav': bool(o.get('fav')), 'assoc': bool(o.get('assoc')),
                    'grp': 'Office / scheduling contact', 'spec': e.get('role') or '', 'office': cut(e.get('officeContact'), 160),
                    'status': o.get('status') or 'none', 'followUp': o.get('followUp') or '', 'onote': cut(o.get('note'), 400), 'obj': '',
                    'hist': [{'d': (h.get('d') or '')[:10], 'a': h.get('a'), 'n': cut(h.get('n'), 200), 'who': h.get('who') or ''} for h in (o.get('history') or [])][-6:]})
    for t in out:  # office contacts and anything else without a parsed affiliation: the hospitals of its area
        if 'aff' not in t:
            t['aff'] = hospitals('', t.get('hosp') or '')
    # route plans (manual stops / hidden drop-ins per day and person) from the Targets app's `plans` collection;
    # Chi West keeps its own copy in its `plans` collection and shows whichever is newer
    plans = {}
    for k, v in (loadall(plans_dir) if plans_dir and os.path.isdir(plans_dir) else {}).items():
        k = k.replace('@', '~')  # ArtifactData spells the "~" in a doc id as "@" in the file name
        if isinstance(v, dict) and (v.get('stops') or v.get('skip')):
            plans[k] = {'stops': v.get('stops') or [], 'skip': v.get('skip') or [], 'updated': v.get('updated') or ''}
    return {'targets': out, 'plans': plans}


# ---------------------------------------------------------------- writes
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--gt', required=True); ap.add_argument('--patients', required=True)
    ap.add_argument('--targets', required=True); ap.add_argument('--outreach'); ap.add_argument('--plans'); ap.add_argument('--cw')
    ap.add_argument('--pdb'); ap.add_argument('--overlay')
    ap.add_argument('out')
    a = ap.parse_args()
    core, cases, fc, stock = build_tracker(a.gt)
    today = core.get('today') or dt.date.today().isoformat()
    docs = {'core': core, 'cases': cases, 'forecast': fc, 'stock': stock,
            'patients': build_patients(a.patients, today, *load_extra(a.pdb, a.overlay)), 'targets': build_targets(a.targets, a.outreach, a.plans)}
    docs['stock'] = stock_users(docs['stock'], docs['patients'], today)
    if not docs['patients']['names']:
        print('NOTE: no --pdb or --overlay, so patients carry no names, coverage or events', file=sys.stderr)
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
