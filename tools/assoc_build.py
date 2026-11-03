#!/usr/bin/env python3
"""Build the data file each associate app shows (Chi West · Caroline, Chi West · Pete).

    python3 -I tools/assoc_build.py --patients PDB --tracker DB --targets TDB --targets-html targets.html \
        [--assoc Caroline=ADB_C --assoc Pete=ADB_P] [--today YYYY-MM-DD] OUT

PDB            Chi West Patients db export (collections `patients`, `events`)
DB             Graft Tracker app db export (collections `team`, `team-todos`)
TDB            Chi West Targets db export (collections `outreach`, `custom`, `plans`)
targets.html   the Chi West Targets page (its <script id="data"> JSON is the master target list)
--assoc        an associate app's own db export (collection `submissions`) — only exists once they sign in
OUT            folder; writes OUT/caroline/data.json and OUT/pete/data.json

What goes in (per associate):
- visits: Patients-app patients whose `cover` is them with a next visit this week or next, plus their
  logged visits (log entries with `cover` = them) this week. Initials only; no names, no prices.
- todos: Graft Tracker `team-todos` for them (open and done this week).
- targets: Targets-app targets on their plan stops (this week + next), anything they have touched
  (history entries with w = them), and any outreach doc with `assign` = them.
- submissions: Graft Tracker `team` docs for them (what AJ has approved / rejected / filed), so their
  app can show the status of what they sent.
Publish OUT/<name>/data.json as the file `data.json` of that associate's artifact.
"""
import argparse
import datetime as dt
import glob
import json
import os
import re

WHO = ['Caroline', 'Pete']
SHORT = {'Advocate Good Samaritan Hospital': 'Good Sam', 'Alexian Brothers Wound Center': 'Alexian WCC',
         'Alexian Ambulatory': 'Alexian Ambulatory', 'Alexian Brothers - Ambulatory': 'Alexian Ambulatory',
         'Northwest Suburban Foot and Ankle Clinic': 'NW Suburban F&A', 'Saint Joseph Hospital - Elgin': 'St. Joseph Elgin',
         'Ascension Saint Alexius': 'St. Alexius', 'Central DuPage Foot and Ankle': 'Central DuPage', 'Elmhurst Hospital': 'Elmhurst',
         'UChicago Medicine AdventHealth GlenOaks': 'GlenOaks', 'Suburban Foot and Ankle Center': 'Suburban F&A',
         'Romano Orthopaedics - Lombard': 'Romano Lombard', 'Ankle and Foot Surgery LLC': 'Ankle & Foot Surgery'}


def load(folder, coll):
    out = {}
    for f in glob.glob(os.path.join(folder, coll, '*.json')):
        try:
            out[os.path.basename(f)[:-5]] = json.load(open(f, encoding='utf-8'))
        except Exception:
            pass
    return out


def dots(ini):
    ini = (ini or '').strip()
    return '.'.join(ini.upper()) + '.' if re.fullmatch(r'[A-Za-z]{2,3}', ini) else ini


def monday(d):
    return d - dt.timedelta(days=d.weekday())


def iso(d):
    return d.strftime('%Y-%m-%d')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--patients', required=True)
    ap.add_argument('--tracker', required=True)
    ap.add_argument('--targets', required=True)
    ap.add_argument('--targets-html', required=True)
    ap.add_argument('--assoc', action='append', default=[])
    ap.add_argument('--today')
    ap.add_argument('out')
    a = ap.parse_args()
    today = dt.date.fromisoformat(a.today) if a.today else dt.date.today()
    ws = monday(today)
    we = ws + dt.timedelta(days=11)          # this week + next week (Fri)
    week = [iso(ws + dt.timedelta(days=i)) for i in range(5)]
    week2 = [iso(ws + dt.timedelta(days=7 + i)) for i in range(5)]

    pts = load(a.patients, 'patients')
    team = load(a.tracker, 'team')
    todos = load(a.tracker, 'team-todos')
    outreach = load(a.targets, 'outreach')
    custom = load(a.targets, 'custom')
    plans = load(a.targets, 'plans')
    html = open(a.targets_html, encoding='utf-8').read()
    m = re.search(r'<script id="data" type="application/json">(.*?)</script>', html, re.S)
    master = {t['id']: t for t in json.loads(m.group(1))} if m else {}
    for k, v in custom.items():
        master.setdefault(k, dict(v, id=k, added=True))
    assoc = {}
    for spec in a.assoc:
        who, folder = spec.split('=', 1)
        assoc[who] = load(folder, 'submissions')

    for who in WHO:
        # visits
        visits = []
        for pid, p in pts.items():
            if p.get('status') in ('done', 'denied'):
                continue
            nxt = p.get('next') or ''
            if p.get('cover') == who and ws <= dt.date.fromisoformat(nxt) <= we if re.fullmatch(r'\d{4}-\d{2}-\d{2}', nxt) else False:
                visits.append({'id': pid, 'd': nxt, 't': p.get('nextTime') or '', 'acct': p.get('account') or '', 'short': SHORT.get(p.get('account'), p.get('account') or ''),
                               'loc': p.get('nextLoc') or '', 'dr': p.get('provider') or '', 'ini': dots(p.get('initials')),
                               'app': (int(p.get('apps') or 0) + 1) if p.get('status') != 'or' else 0,
                               'prod': ' '.join(x for x in [p.get('product') or '', p.get('size') or ''] if x),
                               'note': (p.get('coverNote') or '')[:300], 'verify': p.get('verify') or '', 'logged': None})
            for e in p.get('log') or []:
                d = e.get('date') or ''
                if e.get('cover') == who and d in week:
                    visits.append({'id': pid, 'd': d, 't': e.get('time') or '', 'acct': p.get('account') or '', 'short': SHORT.get(p.get('account'), p.get('account') or ''),
                                   'loc': e.get('loc') or '', 'dr': p.get('provider') or '', 'ini': dots(p.get('initials')),
                                   'app': int(e.get('app') or 0), 'prod': ' '.join(x for x in [p.get('product') or '', e.get('size') or p.get('size') or ''] if x),
                                   'note': '', 'verify': '', 'logged': {'grafted': bool(e.get('grafted')), 'size': e.get('size') or '', 'note': e.get('note') or ''}})
        visits.sort(key=lambda v: (v['d'], v['t'] or '99:99', v['short']))
        # todos
        td = []
        for tid, t in todos.items():
            if t.get('who') != who:
                continue
            if t.get('status') == 'done' and (t.get('doneAt') or '')[:10] < iso(ws):
                continue
            td.append({'id': tid, 'text': t.get('text') or '', 'acct': t.get('acct') or '', 'due': t.get('due') or '',
                       'from': 'AJ' + (' · ' + t['created'][5:10].replace('-', '/').lstrip('0').replace('/0', '/') if t.get('created') else ''),
                       'status': t.get('status') or 'open'})
        td.sort(key=lambda t: (t['status'] != 'open', t['due'] or '9999'))
        # targets
        ids = set()
        for key, P in plans.items():
            if key.endswith('~' + who):
                day = key.split('~')[0]
                if day in week or day in week2:
                    ids.update(P.get('stops') or [])
        for tid, o in outreach.items():
            if o.get('assign') == who or any((h.get('w') == who) for h in (o.get('history') or [])):
                ids.add(tid)
            L = o.get('lunch') or {}
            if L.get('who') == who and (L.get('d') or '') >= iso(ws):
                ids.add(tid)
        targets = []
        for tid in ids:
            t = master.get(tid)
            if not t:
                continue
            o = outreach.get(tid) or {}
            hist = [h for h in (o.get('history') or []) if h.get('a') not in ('edit', 'stage')]
            last = hist[-1] if hist else None
            L = o.get('lunch') or {}
            targets.append({'id': tid, 'name': ', '.join(x for x in [t.get('name') or '', t.get('cred') or ''] if x), 'tier': str(t.get('tier') or '2'),
                            'hosp': t.get('hosp') or '', 'office': (t.get('office') or '').split(';')[0].strip(), 'phone': (t.get('phone') or '').split(';')[0].strip(),
                            'stage': (o.get('status') or 'none'), 'last': last and {'a': last.get('a'), 'd': (last.get('d') or '')[:10], 'w': last.get('w') or 'AJ', 'n': (last.get('n') or '')[:160]},
                            'why': (t.get('angle') or t.get('why') or '')[:220], 'followUp': o.get('followUp') or '',
                            'lunch': (L.get('d') + (' ' + L['t'] if L.get('t') else '')) if L.get('d') and L.get('who') == who else ''})
        targets.sort(key=lambda t: (t['tier'], t['name']))
        # submissions (status feedback)
        subs = []
        for sid, s in team.items():
            if s.get('who') != who:
                continue
            subs.append({'id': sid, 'kind': s.get('kind') or '', 'summary': s.get('summary') or '', 'at': s.get('at') or '', 'status': s.get('status') or 'pending',
                         'srcId': s.get('srcId') or ''})
        subs.sort(key=lambda s: s['at'], reverse=True)
        out = {'who': who, 'built': dt.datetime.now().astimezone().isoformat(timespec='minutes'), 'today': iso(today),
               'weeks': [week, week2], 'visits': visits, 'todos': td, 'targets': targets, 'submissions': subs[:60]}
        d = os.path.join(a.out, who.lower())
        os.makedirs(d, exist_ok=True)
        json.dump(out, open(os.path.join(d, 'data.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=0)
        print(f"{who}: {len(visits)} visits, {len(td)} to-dos, {len(targets)} targets, {len(subs)} submissions -> {d}/data.json")


if __name__ == '__main__':
    main()
