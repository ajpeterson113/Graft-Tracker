#!/usr/bin/env python3
"""Build the view-only "Chi West Schedule" page AJ shares with Caroline and Pete.

It is the Patients tab of Chi West (the week schedule and the patient list) with NO patient names and no notes:
initials, account, provider, product/size, frequency, application count, visit date/time/place, who covers the
visit, each patient's visit history (date, A#, grafted or not, size used; tap a visit or patient), and the same Verify / Confirm flag Chi West shows. No sales data. The data is baked into the page, so it works for anyone with the
link (no sign-in) and changes only when the page is republished.

Usage:
    python3 -I tools/assoc_view.py CW_PATIENTS_JSON OUT_HTML

CW_PATIENTS_JSON is Chi West's data/patients doc: the chiwest_sync.py output (OUT/data/patients.json) or an
ArtifactData export of Chi West's `data` collection. Then publish OUT_HTML to the Chi West Schedule artifact
(https://claude.ai/artifact/5V9C5K8zikma7b7AUjAoFz; Artifact publish with its url, after an Artifact read of it in that
session). The Chi West sync scheduled task does this after every refresh (its step C, AJ 11/5). The script stops without
writing the page if a patient name would be in it.
"""
import json, os, re, sys, datetime as dt

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, 'assoc_view.html')

SHORT = {'Alexian Brothers Wound Center': 'Alexian WCC', 'Advocate Good Samaritan Hospital': 'Good Sam',
         'Ascension Saint Alexius': 'St. Alexius', 'Saint Joseph Hospital - Elgin': 'St. Joseph Elgin',
         'Central DuPage Foot and Ankle': 'Central DuPage F&A', 'Northwest Suburban Foot and Ankle Clinic': 'NW Suburban F&A',
         'Suburban Foot and Ankle Center': 'Suburban F&A', 'Romano Orthopaedics - Lombard': 'Romano Lombard',
         'Elmhurst Hospital': 'Elmhurst', 'UChicago Medicine AdventHealth GlenOaks': 'GlenOaks',
         'Ankle and Foot Surgery LLC': 'Ankle & Foot Surgery', 'Alexian Ambulatory': 'Alexian Ambulatory',
         'Alexian Brothers - Ambulatory': 'Alexian Ambulatory', 'Sherman': 'Sherman'}
# same account order (and so the same account colors) as Chi West
ORDER = ['Advocate Good Samaritan Hospital', 'Alexian Brothers Wound Center', 'Alexian Ambulatory', 'Northwest Suburban Foot and Ankle Clinic',
         'Saint Joseph Hospital - Elgin', 'Sherman', 'Ascension Saint Alexius', 'Central DuPage Foot and Ankle', 'Elmhurst Hospital',
         'UChicago Medicine AdventHealth GlenOaks', 'Suburban Foot and Ankle Center', 'Romano Orthopaedics - Lombard', 'Ankle and Foot Surgery LLC']


def norm_ini(s):
    return re.sub(r'[^A-Z]', '', str(s or '').upper())


def fmt_ini(s):
    s = str(s or '').strip()
    if re.fullmatch(r'[A-Za-z]{2,3}', s):
        return '.'.join(s.upper()) + '.'
    return s or '?'


def safe_ini(ini, name):
    """Initials only. A field that holds a name (e.g. a surname typed where the initials go) is cut to its initials."""
    s = str(ini or '').strip()
    letters = re.sub(r'[^A-Za-z]', '', s)
    nm = str(name or '').lower()
    looks_like_name = len(letters) > 3 or (len(letters) > 1 and not s.isupper() and nm and s.lower() in nm)
    if looks_like_name:
        return ''.join(w[0].upper() + '.' for w in re.findall(r'[A-Za-z]+', s)) or '?'
    return fmt_ini(s)


def acct_key(a):
    a = a or ''
    if a == 'Alexian Brothers - Ambulatory':
        a = 'Alexian Ambulatory'
    return a


def build(src):
    d = json.load(open(src))
    d = d.get('data', d)
    today = dt.date.today()
    lo, hi = (today - dt.timedelta(days=63)).isoformat(), (today + dt.timedelta(days=42)).isoformat()  # 9 weeks back for each patient's history
    pts, index, src_rows = [], {}, []
    for p in d.get('patients') or []:
        a = acct_key(p.get('acct'))
        k = a + '|' + norm_ini(p.get('ini'))
        i = len(pts)
        index.setdefault(k, []).append(i)
        src_rows.append(p)
        pts.append({
            'acct': SHORT.get(a, a), 'ac': ORDER.index(a) % 12 if a in ORDER else 11,
            'prov': p.get('prov') or '', 'ini': safe_ini(p.get('ini'), p.get('name')), 'status': p.get('status') or '',
            'product': p.get('product') or '', 'size': p.get('size') or '', 'freq': p.get('freq') or '',
            'apps': p.get('apps') if isinstance(p.get('apps'), (int, float)) else None,
            'next': p.get('next') or '', 'time': p.get('time') or '', 'loc': p.get('loc') or '',
            # AJ's label for a booked visit (e.g. "Stravix follow up"), shown instead of the A# (AJ 11/6)
            'labels': {k: str(v)[:60] for k, v in (p.get('labels') or {}).items() if k >= today.isoformat()},
            'cover': p.get('cover') or '', # same flags as Chi West's schedule row (AJ 11/6): "Verify A#" (application count to check) and "Confirm Time" (time to confirm)
            'flag': ' · '.join((["Verify A#"] if str(p.get('verify') or '').startswith('App #') else [])
                                + (["Confirm Time"] if p.get('fuVerify') or (p.get('verify') and not str(p.get('verify')).startswith('App #')) else [])),
        })
    visits = {}

    def add(i, date, time='', loc='', app=None, res='', size=''):
        if not date or not (lo <= date <= hi):
            return
        v = visits.setdefault((i, date), {'p': i, 'date': date, 'time': '', 'loc': '', 'app': None, 'res': '', 'size': ''})
        v['size'] = v['size'] or (str(size) if size and res == 'Grafted' else '')
        v['time'] = v['time'] or time or ''
        v['loc'] = v['loc'] or loc or ''
        v['app'] = app if (res and app is not None) else (v['app'] if v['app'] is not None else app)  # a logged result's A# wins
        v['res'] = res or v['res']

    def res_of(s):
        s = str(s or '')
        return 'Grafted' if s.startswith('Grafted') else 'Not grafted' if s.startswith('Not grafted') else ''

    low = lambda s: str(s or '').strip().lower()

    def who(row):
        """The patient a visit / log row belongs to. Rows carry only account + initials, and two patients can share
        those (an office patient and their OR case, two people with the same initials), so pick the best fit."""
        c = index.get(acct_key(row.get('acct')) + '|' + norm_ini(row.get('ini'))) or []
        if len(c) < 2:
            return c[0] if c else None
        date, is_or = row.get('date') or '', 'OR' in str(row.get('status') or '')

        def score(i):
            p = src_rows[i]
            return ((p.get('next') == date) * 8 + (p.get('last') == date) * 4
                    + (bool(row.get('product')) and low(p.get('product')).startswith(low(row.get('product'))[:6])) * 2
                    + (bool(row.get('size')) and low(p.get('size')).startswith(low(row.get('size')))) * 1
                    + ((p.get('status') == 'or') == is_or) * 1)
        return max(c, key=score)  # ties keep the first

    for v in d.get('visits') or []:
        i = who(v)
        if i is not None:
            r = res_of(v.get('status'))
            if r and (v.get('date') or '') > today.isoformat():  # AJ 11/6: a result logged on a later booked visit (a no-show)
                r, v = '', dict(v, app=None)                        # doesn't close it; it stays a normal appointment
            add(i, v.get('date'), v.get('time'), v.get('loc'), v.get('app'), r, v.get('size'))
    for e in d.get('log') or []:
        i = who(e)
        if i is not None and e.get('date', '') <= today.isoformat():
            add(i, e.get('date'), '', '', e.get('app'), res_of(e.get('result')), e.get('size'))
    for i, p in enumerate(pts):
        if p['status'] in ('done', 'denied'):
            continue
        add(i, p['next'], p['time'], p['loc'])
    events = [{'title': e.get('title') or '', 'kind': e.get('kind') or '', 'place': e.get('place') or '',
               'date': e.get('date') or '', 'time': e.get('time') or '', 'end': e.get('end') or ''}
              for e in d.get('events') or [] if lo <= (e.get('date') or '') <= hi]
    return {'updated': d.get('synced') or dt.datetime.now().isoformat(timespec='minutes'), 'built': today.isoformat(),
            'patients': pts, 'visits': sorted(visits.values(), key=lambda v: (v['date'], v['time'])), 'events': events}


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    data = build(sys.argv[1])
    # nothing that could identify a patient beyond initials goes out
    blob = json.dumps(data, separators=(',', ':'), ensure_ascii=False)
    assert '"name"' not in blob and '"note' not in blob
    # last check: no patient's full name, and no part of a patient's name in that patient's own row
    src = json.load(open(sys.argv[1]))
    src = src.get('data', src)
    for p, q in zip(src.get('patients') or [], data['patients']):
        n = str(p.get('name') or '').strip()
        if not n:
            continue
        assert len(n) < 4 or n not in blob, 'a patient name would be published'
        own = ' '.join(str(q.get(k) or '') for k in ('ini', 'loc'))
        for w in re.findall(r'[A-Za-z]{3,}', n):
            assert not re.search(r'\b' + re.escape(w) + r'\b', own, re.I), 'part of a patient name would be published'
    html = open(TEMPLATE).read().replace('/*DATA*/null', blob.replace('</', '<\\/'))
    open(sys.argv[2], 'w').write(html)
    print(f"{sys.argv[2]}: {len(data['patients'])} patients, {len(data['visits'])} visits, {len(data['events'])} events, "
          f"{len(html)//1024} KB")


if __name__ == '__main__':
    main()
