#!/usr/bin/env python3
"""Dollar potential by product, per surgeon and per account, from the claims export AJ keeps in the Project files.

    python3 -I tools/claims_potential.py --providers PROVIDERS.tsv --accounts ACCOUNTS.tsv \\
        --targets Chi_West_Surgeon_Targets.xlsx OUT/potential.json

PROVIDERS.tsv  the "Full Name / Brand / Potential / Volume" export (Project doc "Copy_of_AWM_Sales_Call (1).xlsx"),
               saved as tab-separated text
ACCOUNTS.tsv   the facility "Account Name / Brand / Potential / Volume" export for the territory
               (Project doc "Copy_of_AWM_Sales_Call (4).xlsx"), saved the same way
--targets      the targets workbook: its "Claims - all providers" sheet maps each provider name to an NPI and its
               "By hospital" sheet says who is already a target
OUT            one JSON document, written as is to Chi West `data/potential` and to the Targets app `intel/potential`

The document: {updated, src, floor, by: {<npi>: {n, s, g, v, o}}, off: [...], acct: [...]}. s / g / v / o are
[dollars a year, cases] for STRAVIX, GRAFIX, VERSAJET and OASIS. `off` lists the surgeons with real volume who are
not on the target list; `acct` the facilities. Volumes of 5, 10 or 15 are the export's floor for small counts, so
the apps mark anything under `floor` cases as thin.

AJ 11/9: the dollar figures are company data. They live only in the two private apps; never write them to this
repo (it is public) or to a page shared by link.
"""
import argparse, collections, csv, datetime as dt, json, re
import openpyxl

BR = {'Stravix': 's', 'Grafix': 'g', 'Versajet': 'v', 'Oasis': 'o'}
norm = lambda s: re.sub(r'[^a-z]', '', str(s or '').lower())


def tsv(path):
    with open(path, encoding='utf-8') as f:
        return [r for r in csv.DictReader(f, delimiter='\t')]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--providers', required=True); ap.add_argument('--accounts', required=True)
    ap.add_argument('--targets', required=True); ap.add_argument('--floor', type=int, default=20)
    ap.add_argument('--src', default='Claims export in the Project files (AWM Sales Call), added 11/9/2026')
    ap.add_argument('out')
    a = ap.parse_args()
    wb = openpyxl.load_workbook(a.targets, read_only=True)
    it = wb['Claims – all providers'].iter_rows(values_only=True); h = next(it)
    claims = {norm(d['Provider']): d for d in (dict(zip(h, r)) for r in it) if d.get('Provider')}
    it = wb['By hospital'].iter_rows(values_only=True); h = next(it)
    on = {str(d.get('NPI')).strip() for d in (dict(zip(h, r)) for r in it) if d.get('Name') and d.get('NPI')}
    pot = collections.defaultdict(dict)
    for r in tsv(a.providers):
        k = BR.get(r['Brand'])
        if k and float(r['Volume'] or 0):
            pot[r['Full Name']][k] = [round(float(r['Potential'])), round(float(r['Volume']))]
    by, off, miss = {}, [], []
    for name, b in pot.items():
        c = claims.get(norm(name))
        if not c:
            miss.append(name); continue
        npi = str(c['NPI'])
        by[npi] = dict(b, n=name)
        tissue = sum(b.get(k, [0])[0] for k in 'sgo')
        solid = any(b.get(k, [0, 0])[1] > 10 for k in 'sg')
        if npi not in on and solid:
            off.append(dict(b, npi=npi, n=name, t=tissue, spec=str(c.get('Specialty (claims file)') or '').title(),
                            addr=str(c.get('Address (claims file)') or '').title(), city=str(c.get('City') or '').title(),
                            hosp=c.get('Nearest hospital cluster') or '', cust=c.get('Existing customer') or ''))
    off.sort(key=lambda x: -x['t'])
    acc = collections.defaultdict(dict)
    for r in tsv(a.accounts):
        k = BR.get(r['Brand'])
        if k and float(r['Volume'] or 0):
            acc[r['Account Name']][k] = [round(float(r['Potential'])), round(float(r['Volume']))]
    acct = [dict(b, a=n, t=sum(b.get(k, [0])[0] for k in 'sgo')) for n, b in acc.items() if sum(v[0] for v in b.values()) >= 20000]
    acct.sort(key=lambda x: -x['t'])
    doc = {'updated': dt.date.today().isoformat(), 'src': a.src, 'floor': a.floor, 'by': by, 'off': off[:40], 'acct': acct}
    json.dump(doc, open(a.out, 'w'), ensure_ascii=False, separators=(',', ':'))
    print(f"{len(by)} surgeons with potential ({len([n for n in by if n in on])} on the target list), {len(off)} off the list above the floor "
          f"(top {len(doc['off'])} kept), {len(acct)} accounts; {len(miss)} names with no NPI match. {a.out}")


if __name__ == '__main__':
    main()
