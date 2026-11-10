#!/usr/bin/env python3
"""Build the view-only "Chi West Schedule" page AJ shares with Caroline and Pete.

It is the Patients tab of Chi West (the week schedule and the patient list) with NO patient names and no notes:
initials, account, provider, product/size, frequency, application count, visit date/time/place, who covers the
visit, each patient's visit history (date, A#, grafted or not, size used; tap a visit or patient), and the same Verify / Confirm flag Chi West shows. No sales data. The data is baked into the page, so it works for anyone with the
link (no sign-in) and changes only when the page is republished.

Targets tab (AJ 11/10): Chi West's Grow > Targets list, view only: each target's tier, stage, follow-up due, last touch,
specialty, hospitals, My targets / Associates / VERSAJET flags, office, phone, clinic days, why they are a target and the
research brief. Archived targets and current customers are left off. It never carries dollar figures (claims potential is
company data and the page is shared by link), AJ's private notes or touch text, or email addresses.

Usage:
    python3 -I tools/assoc_view.py CW_PATIENTS_JSON OUT_HTML [--targets CW_TARGETS_JSON] [--research CW_RESEARCH_JSON]

CW_TARGETS_JSON / CW_RESEARCH_JSON are Chi West's data/targets and data/research docs. Without --targets the page has
no Targets tab, so always pass it when republishing.

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
         'Alexian Brothers - Ambulatory': 'Alexian Ambulatory', 'Sherman': 'Sherman',
         'Alexian Brothers Medical Center': 'Alexian Main OR'}
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


GSHORT = {'Podiatry / foot & ankle': 'Podiatry', 'Vascular surgery': 'Vascular', 'Plastics / reconstruction': 'Plastics',
          'General / trauma / colorectal': 'General surgery', 'Orthopedics (trauma / F&A)': 'Ortho', 'Urology': 'Urology',
          'Wound center clinician': 'Wound center', 'Infectious disease (wound center)': 'Infectious disease',
          'Interventional radiology (referral)': 'IR', 'Office / scheduling contact': 'Office contact', 'Other': 'Other'}
MONEY = re.compile(r'\$\s?\d|\bdollars?\b|\bpotential\b', re.I)
EMAIL = re.compile(r'[\w.+-]+@[\w-]+(?:\.[\w-]+)+')


def _clean(text, limit=320):
    """Text that is safe on a link-shared page: sentences that carry money are dropped, email addresses removed."""
    t = str(text or '').replace('\u2026', '').strip()
    if not t:
        return ''
    parts = [x for x in re.split(r'(?<=[.;])\s+', t) if not MONEY.search(x)]
    t = EMAIL.sub('', ' '.join(parts)).strip()
    t = re.sub(r'\s{2,}', ' ', re.sub(r'\s+([,.;])', r'\1', t))
    return t[:limit].rstrip()


def build_targets(tfile, rfile):
    d = json.load(open(tfile))
    d = d.get('data', d)
    briefs = {}
    if rfile and os.path.exists(rfile):
        r = json.load(open(rfile))
        briefs = (r.get('data', r).get('briefs') or {})
    out = []
    for t in d.get('targets') or []:
        if t.get('arch') or t.get('cust') or t.get('archived'):
            continue
        hist = [h for h in (t.get('hist') or []) if h.get('a') in ('emailed', 'called', 'texted', 'visited', 'met', 'lunch')]
        last = max((str(h.get('d') or '')[:10] for h in hist), default='')
        grp = t.get('grp') or ''
        nm = str(t.get('name') or '').strip()
        cred = str(t.get('cred') or '')
        doc = bool(re.search(r'\b(MD|DO|DPM|MBBS)\b', cred)) and not nm.lower().startswith('dr')
        lunch = t.get('lunch') or {}
        b = briefs.get(t.get('id')) or {}
        out.append({
            'id': t.get('id'), 'n': ('Dr. ' if doc else '') + nm, 'cred': cred, 'tier': str(t.get('tier') or ''),
            'grp': GSHORT.get(grp, grp), 'spec': _clean(t.get('spec'), 90), 'aff': [[a[0], a[1]] for a in (t.get('aff') or [])][:4],
            'st': t.get('status') or 'none', 'due': str(t.get('followUp') or '')[:10], 'last': last,
            'fav': bool(t.get('fav')), 'assoc': bool(t.get('assoc')), 'vj': str(t.get('vj') or ''),
            'office': _clean(t.get('office'), 200), 'phone': _clean(t.get('phone'), 60), 'prac': _clean(t.get('practice'), 90),
            'where': _clean(t.get('where'), 160), 'why': _clean(t.get('angle'), 320), 'fit': _clean(t.get('fit'), 240),
            'obj': _clean(t.get('obj'), 200),
            'mtg': ({'d': lunch.get('d'), 't': lunch.get('t') or ''} if lunch.get('d') else None),
            'rd': str(b.get('date') or ''), 'res': [x for x in (_clean(i, 360) for i in (b.get('items') or [])) if x][:8],
            '_s': (0 if t.get('fav') else 1, {'1': 0, 'Customer \u2013 expand': 0, '2': 1, '3': 2}.get(str(t.get('tier') or ''), 3), -(t.get('score') or 0)),
        })
    out.sort(key=lambda x: x['_s'])
    for x in out:
        del x['_s']
    blob = json.dumps(out, ensure_ascii=False)
    assert not re.search(r'\$\s?\d', blob), 'a dollar figure would be published on the Targets tab'
    assert not EMAIL.search(blob), 'an email address would be published on the Targets tab'
    return out


def main():
    args = sys.argv[1:]
    opt = {}
    for k in ('--targets', '--research'):
        if k in args:
            i = args.index(k)
            opt[k] = args[i + 1]
            del args[i:i + 2]
    if len(args) != 2:
        sys.exit(__doc__)
    sys.argv = [sys.argv[0]] + args
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
    targets = build_targets(opt['--targets'], opt.get('--research')) if opt.get('--targets') else None
    if targets is not None:
        tblob = json.dumps(targets, ensure_ascii=False)
        # a patient's full name must not ride in on a target's text either
        for p in src.get('patients') or []:
            n = str(p.get('name') or '').strip()
            assert len(n) < 6 or ' ' not in n or n not in tblob, 'a patient name would be published on the Targets tab'
        data['targets'] = targets
        blob = json.dumps(data, separators=(',', ':'), ensure_ascii=False)
    html = open(TEMPLATE).read().replace('/*DATA*/null', blob.replace('</', '<\\/'))
    open(sys.argv[2], 'w').write(html)
    print(f"{sys.argv[2]}: {len(data['patients'])} patients, {len(data['visits'])} visits, {len(data['events'])} events, "
          f"{len(targets) if targets is not None else 'no'} targets, {len(html)//1024} KB")


if __name__ == '__main__':
    main()
