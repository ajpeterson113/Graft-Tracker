#!/usr/bin/env python3
"""Add a new patient to the tracker's Projections sheet (Potential new starts section).

AJ 11/5: every new patient that comes in with a graft size goes into projections, whatever its category
(Pending included), from the run that adds the patient: a Chi West `newpatient` tap, the Patients app's
"+ New patient", a new patient AJ mentions, or a repo inbox/new-patients file. A new patient with no size
stays out until a size is given. Upside only, like every potential new start, unless AJ gives a downside
(--down-size/--down-price). Skip it if the patient already has a row (same account + initials) anywhere
on the sheet.

    python3 -I tools/proj_newstart.py MASTER.xlsx --acct "Advocate Good Samaritan Hospital" --prov "Dr. Ayesh" \
        --ini Z.G. --product GRAFIX --size 16mm --freq Biweekly --start 2026-11-19 --note "..." \
        [--price 425.5] [--units 1] [--down-size S --down-price P]

--start is the first visit (YYYY-MM-DD); with no date the row counts 0 visits until one is set.
--freq is Weekly, Biweekly or Once. --price defaults to the account's price for that size from the sheet's
"Downside lookups" table, then from other rows on the sheet; the script stops if it finds none.

It inserts a row at the end of the section (above the "Potential new starts" total), re-points every formula
in the workbook that refers to Projections rows below it, widens the section's sums and per-account SUMIFs
to take the new row, and copies the row formulas from the row above. Recalculate the Excel afterwards
(the xlsx skill's recalc.py), then build as usual.
"""
import argparse
import datetime as dt
import re
import sys

import openpyxl
from openpyxl.formula.tokenizer import Token, Tokenizer
from openpyxl.formula.translate import Translator

SHEET = 'Projections'
REF = re.compile(r'^(\$?)([A-Z]{1,3})(\$?)(\d+)$')
PRODUCT = {'GRAFIX': 'Grafix PL Prime', 'GRAFIX PRIME': 'Grafix PL Prime', 'GRAFIX PL': 'Grafix PL Prime',
           'OASIS WOUND': 'OASIS Wound Matrix', 'OASIS WOUND MATRIX': 'OASIS Wound Matrix', 'OASIS ULTRA': 'OASIS Ultra',
           'OASIS BURN': 'OASIS Burn Matrix', 'OASIS BURN MATRIX': 'OASIS Burn Matrix', 'STRAVIX': 'Stravix PL'}


def norm_size(s):
    s = str(s or '').strip()
    if re.fullmatch(r'16\s*mm', s, re.I):
        return '16mm disc'
    m = re.fullmatch(r'(\d+(?:\.\d+)?)\s*[xX×]\s*(\d+(?:\.\d+)?)', s)
    return f'{m[1]} x {m[2]}' if m else s


def find_row(ws, pred):
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and pred(v.strip()):
            return r
    return None


def shift_ref(ref, at, n, grow_from):
    """One range/cell reference without a sheet name. Rows >= at move down n; a range that starts inside the
    section (grow_from up to the insert point) and ends inside it grows to the section's new last row, so the
    total and the per-account SUMIFs take the new rows (some of those ranges stopped short of the last row)."""
    parts = ref.split(':')
    m = [REF.match(p) for p in parts]
    if not all(m):
        return ref
    rows = [int(x[4]) for x in m]
    new = [r + n if r >= at else r for r in rows]
    if len(rows) == 2 and grow_from <= rows[0] < at and rows[0] < rows[1] < at:
        new[1] = at - 1 + n
    return ':'.join(f'{x[1]}{x[2]}{x[3]}{r}' for x, r in zip(m, new))


def shift_formula(f, at, n, grow_from, same_sheet):
    tok = Tokenizer(f)
    changed = False
    for t in tok.items:
        if t.type != Token.OPERAND or t.subtype != Token.RANGE:
            continue
        v = t.value
        if '!' in v:
            sh, ref = v.rsplit('!', 1)
            if sh.strip("'") != SHEET:
                continue
            nv = sh + '!' + shift_ref(ref, at, n, grow_from)
        elif same_sheet:
            nv = shift_ref(v, at, n, grow_from)
        else:
            continue
        if nv != v:
            t.value, changed = nv, True
    return tok.render() if changed else f


def insert_rows(wb, at, n, grow_from):
    ws = wb[SHEET]
    for other in wb.worksheets:
        for row in other.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('='):
                    c.value = shift_formula(c.value, at, n, grow_from, other is ws)
    heights = {r: d.height for r, d in ws.row_dimensions.items() if r >= at and d.height}
    ws.insert_rows(at, n)
    for r in sorted(heights, reverse=True):
        ws.row_dimensions[r + n].height = heights[r]
        ws.row_dimensions[r].height = None
    for dv in ws.data_validations.dataValidation:
        dv.sqref = openpyxl.worksheet.cell_range.MultiCellRange(
            ' '.join(shift_ref(str(cr), at, n, grow_from) for cr in dv.sqref.ranges))
    assert not any(m.max_row >= at for m in ws.merged_cells.ranges), 'merged cells below the insert point'


def price_for(ws, acct, size, lookup_rows):
    for r in lookup_rows:
        if ws.cell(r, 1).value == acct and ws.cell(r, 2).value == size and isinstance(ws.cell(r, 3).value, (int, float)):
            return ws.cell(r, 3).value
    for r in range(40, ws.max_row + 1):  # any patient row at the same account and size
        if ws.cell(r, 1).value == acct and ws.cell(r, 5).value == size and isinstance(ws.cell(r, 11).value, (int, float)):
            return ws.cell(r, 11).value
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('xlsx')
    for k in ('acct', 'prov', 'ini', 'product', 'size', 'freq', 'start', 'note', 'down_size'):
        ap.add_argument('--' + k.replace('_', '-'), default='')
    ap.add_argument('--price', type=float)
    ap.add_argument('--down-price', type=float)
    ap.add_argument('--units', type=float, default=1)
    a = ap.parse_args()
    if not (a.acct and a.ini and a.size):
        sys.exit('--acct, --ini and --size are required (no size: the patient stays out of projections)')
    wb = openpyxl.load_workbook(a.xlsx)
    ws = wb[SHEET]
    ini = a.ini if '.' in a.ini else '.'.join(a.ini.upper()) + '.'
    for r in range(40, ws.max_row + 1):
        if ws.cell(r, 1).value == a.acct and str(ws.cell(r, 3).value or '').replace('.', '') == ini.replace('.', ''):
            sys.exit(f'{ini} at {a.acct} already has a row ({SHEET} row {r}); nothing added')
    t = find_row(ws, lambda v: v.startswith('Potential new starts ('))
    total = find_row(ws, lambda v: v == 'Potential new starts')
    assert t and total and total > t + 2, 'Potential new starts section not found'
    size = norm_size(a.size)
    lk = find_row(ws, lambda v: v == 'Downside lookups')
    price = a.price if a.price is not None else price_for(ws, a.acct, size, range(lk + 2, lk + 40) if lk else [])
    if price is None:
        sys.exit(f'no price on the sheet for {size} at {a.acct}: pass --price')
    at, prev = total, total - 1
    insert_rows(wb, at, 1, t + 2)
    for col in range(1, 18):
        src = ws.cell(prev, col)
        dst = ws.cell(at, col)
        dst._style = src._style
        dst.value = None
    f = lambda col: Translator(ws.cell(prev, col).value, f'{ws.cell(prev, col).column_letter}{prev}').translate_formula(
        f'{ws.cell(at, col).column_letter}{at}')
    start = dt.datetime.fromisoformat(a.start) if a.start else None
    prod = PRODUCT.get(a.product.strip().upper(), a.product.strip() or 'Grafix PL Prime')
    vals = {1: a.acct, 2: a.prov, 3: ini, 4: prod, 5: size, 6: a.freq or 'Weekly', 7: start, 8: 0,
            9: f(9), 10: a.units, 11: price, 12: f(12), 14: 0, 17: a.note}
    if a.down_size and a.down_price is not None:
        vals.update({13: norm_size(a.down_size), 15: a.down_price,
                     16: f'=IF(O{at}="",0,MAX(0,I{at}*J{at}-N{at})*O{at})'})
    else:
        vals[16] = 0
    for col, v in vals.items():
        ws.cell(at, col).value = v
    wb.save(a.xlsx)
    print(f'{SHEET} row {at}: {ini} {a.acct} {prod} {size} {a.freq} from {a.start or "TBD"} at ${price:,.2f} '
          f'({"downside " + norm_size(a.down_size) if a.down_size else "upside only"}); recalc before building')


if __name__ == '__main__':
    main()
